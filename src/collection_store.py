"""Persistent Jira collection checkpoints backed by Supabase PostgREST RPCs.

The publishable Supabase key is intentionally safe to ship with the server code.
Actual access to a user's collection is capability-scoped by a 256-bit owner key
that is derived from the server-side application login secret and never sent to
the browser.
"""
from __future__ import annotations

import hashlib
import os
import time
import json
import sqlite3
from pathlib import Path
from typing import Any

import requests


DEFAULT_SUPABASE_URL = "https://cqnjiwaqwtfvylvqpznk.supabase.co"
DEFAULT_SUPABASE_PUBLISHABLE_KEY = "sb_publishable_3SQeSug1wrgVixggCeVaYQ_92lXMMGD"


class PersistenceError(RuntimeError):
    def __init__(self, message: str, *, retryable: bool = True):
        super().__init__(message)
        self.retryable = retryable


def persistence_owner_key(settings) -> str:
    """Return a stable secret capability for this configured application account.

    Jira credential rotation does not invalidate stored checkpoints; changing the
    app login password intentionally does.
    """
    login_secret = settings.password_hash or settings.password
    material = f"jira-performance-collection-v1\0{settings.username}\0{login_secret}"
    return hashlib.sha256(material.encode()).hexdigest()


class CollectionStore:
    MAX_ATTEMPTS = 4

    def __init__(self, url: str, publishable_key: str, *, session=None,
                 timeout=(10, 45), sleep=time.sleep):
        self.url = url.rstrip("/")
        self.key = publishable_key
        self.session = session or requests.Session()
        self.timeout = timeout
        self.sleep = sleep
        self.session.headers.update({
            "apikey": self.key,
            "Authorization": f"Bearer {self.key}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        })

    @classmethod
    def configured(cls):
        # Local-only persistence by default. Jira payloads never leave this server.
        return LocalCollectionStore(os.getenv("COLLECTION_STATE_DB", ".collection_state.sqlite3"))

    def close(self):
        self.session.close()

    @staticmethod
    def _safe_error_detail(response):
        try:
            body = response.json()
            detail = body.get("message") or body.get("details") or body.get("hint")
        except ValueError:
            detail = None
        if detail and len(str(detail)) < 240:
            return f" ({detail})"
        return ""

    def _rpc(self, name: str, payload: dict[str, Any]):
        last_error = None
        for attempt in range(self.MAX_ATTEMPTS):
            try:
                response = self.session.post(
                    f"{self.url}/rest/v1/rpc/{name}", json=payload,
                    timeout=self.timeout, allow_redirects=False,
                )
            except requests.RequestException as exc:
                last_error = exc
                if attempt + 1 < self.MAX_ATTEMPTS:
                    self.sleep(min(2 ** attempt, 4))
                    continue
                raise PersistenceError(
                    "The persistent collection store could not be reached after repeated attempts. "
                    "Your Jira request was not marked complete."
                ) from exc

            status = response.status_code
            if status == 429 or 500 <= status <= 599:
                last_error = PersistenceError(
                    f"The persistent collection store is temporarily unavailable (HTTP {status}).",
                    retryable=True,
                )
                if attempt + 1 < self.MAX_ATTEMPTS:
                    try:
                        delay = float(response.headers.get("Retry-After", min(2 ** attempt, 4)))
                    except (TypeError, ValueError):
                        delay = float(min(2 ** attempt, 4))
                    self.sleep(max(0.0, min(delay, 8.0)))
                    continue
                raise last_error

            if status < 200 or status >= 300:
                raise PersistenceError(
                    f"The persistent collection store rejected the request (HTTP {status})."
                    f"{self._safe_error_detail(response)}",
                    retryable=False,
                )

            if not response.content:
                return None
            try:
                return response.json()
            except ValueError as exc:
                last_error = exc
                if attempt + 1 < self.MAX_ATTEMPTS:
                    self.sleep(min(2 ** attempt, 4))
                    continue
                raise PersistenceError(
                    "The persistent collection store returned an unreadable response after repeated attempts."
                ) from exc

        raise PersistenceError(
            "The persistent collection store could not complete the request."
        ) from last_error

    def begin(self, owner_key, job_id, fingerprint, space, query, definitions, cutoff):
        return self._rpc("collection_begin", {
            "p_owner_key": owner_key,
            "p_job_id": job_id,
            "p_fingerprint": fingerprint,
            "p_space": space,
            "p_query": query,
            "p_definitions": definitions,
            "p_cutoff": cutoff,
        })

    def seed_items(self, owner_key, job_id, items):
        return self._rpc("collection_seed_items", {
            "p_owner_key": owner_key,
            "p_job_id": job_id,
            "p_items": items,
        })

    def save_item(self, owner_key, job_id, issue_key, item, history):
        return self._rpc("collection_save_item", {
            "p_owner_key": owner_key,
            "p_job_id": job_id,
            "p_issue_key": issue_key,
            "p_item": item,
            "p_history": history,
        })

    def update_job(self, owner_key, job_id, status, stage, current_issue, message,
                   error_detail=None, result_meta=None):
        return self._rpc("collection_update_job", {
            "p_owner_key": owner_key,
            "p_job_id": job_id,
            "p_status": status,
            "p_stage": stage,
            "p_current_issue": current_issue,
            "p_message": message,
            "p_error_detail": error_detail,
            "p_result_meta": result_meta,
        })

    def load_job(self, owner_key, job_id):
        return self._rpc("collection_load_job", {
            "p_owner_key": owner_key,
            "p_job_id": job_id,
        })

    def latest_resumable(self, owner_key):
        return self._rpc("collection_latest_resumable", {"p_owner_key": owner_key})

    def latest_for_fingerprint(self, owner_key, fingerprint):
        return self._rpc("collection_latest_for_fingerprint", {
            "p_owner_key": owner_key,
            "p_fingerprint": fingerprint,
        })

class LocalCollectionStore:
    """SQLite checkpoint store kept on the application host."""
    def __init__(self, path):
        self.path = str(Path(path))
        parent = Path(self.path).parent
        if str(parent) not in ("", "."):
            parent.mkdir(parents=True, exist_ok=True)
        self.db = sqlite3.connect(self.path, check_same_thread=False)
        self.db.execute("CREATE TABLE IF NOT EXISTS collection_jobs (job_id TEXT PRIMARY KEY, owner_key TEXT NOT NULL, payload TEXT NOT NULL)")
        self.db.commit()

    def close(self):
        self.db.close()

    def _read(self, job_id):
        row = self.db.execute("SELECT owner_key,payload FROM collection_jobs WHERE job_id=?", (job_id,)).fetchone()
        if not row:
            return None
        payload = json.loads(row[1]); payload["owner_key"] = row[0]; return payload

    def _write(self, payload):
        owner = payload["owner_key"]; data = dict(payload); data.pop("owner_key", None)
        self.db.execute("INSERT INTO collection_jobs(job_id,owner_key,payload) VALUES(?,?,?) ON CONFLICT(job_id) DO UPDATE SET owner_key=excluded.owner_key,payload=excluded.payload",
                        (payload["job_id"], owner, json.dumps(data, default=str)))
        self.db.commit()

    def begin(self, owner_key, job_id, fingerprint, space, query, definitions, cutoff):
        existing = self._read(job_id)
        if existing:
            return {k:v for k,v in existing.items() if k != "owner_key"}
        payload = {"job_id":job_id,"owner_key":owner_key,"fingerprint":fingerprint,"space":space,"query":query,"definitions":definitions,"cutoff":cutoff,"status":"running","stage":"Ready","current_issue":None,"last_successful_issue":None,"message":"Persistent checkpoint created.","error_detail":None,"result_meta":None,"items":[],"total_count":0,"completed_count":0}
        self._write(payload); return {k:v for k,v in payload.items() if k != "owner_key"}

    def seed_items(self, owner_key, job_id, items):
        payload = self._read(job_id)
        if not payload or payload["owner_key"] != owner_key: raise RuntimeError("wrong owner")
        existing = {r["issue_key"]:r for r in payload.get("items",[])}; rows=[]
        for pos,item in enumerate(items):
            key=item["key"]; row=existing.get(key,{"position":pos,"issue_key":key,"seed_item":item,"item_data":None,"history_data":None,"completed":False})
            row["position"]=pos; row["seed_item"]=item; rows.append(row)
        payload["items"]=rows; payload["total_count"]=len(rows); self._write(payload); return {"total_count":len(rows)}

    def save_item(self, owner_key, job_id, issue_key, item, history):
        payload=self._read(job_id)
        if not payload or payload["owner_key"] != owner_key: raise RuntimeError("wrong owner")
        for row in payload["items"]:
            if row["issue_key"] == issue_key: row.update(item_data=item, history_data=history, completed=True); break
        payload["completed_count"]=sum(bool(r.get("completed")) for r in payload["items"]); payload["last_successful_issue"]=issue_key; payload["current_issue"]=None
        payload["message"]=f"Completed {payload['completed_count']} of {payload['total_count']} work items."; self._write(payload)
        return {"completed_count":payload["completed_count"],"total_count":payload["total_count"]}

    def update_job(self, owner_key, job_id, status, stage, current_issue, message, error_detail=None, result_meta=None):
        payload=self._read(job_id)
        if not payload or payload["owner_key"] != owner_key: raise RuntimeError("wrong owner")
        payload.update(status=status,stage=stage,current_issue=current_issue,message=message,error_detail=error_detail)
        if result_meta is not None: payload["result_meta"]=result_meta
        self._write(payload); return {"status":status,"completed_count":payload["completed_count"],"total_count":payload["total_count"]}

    def load_job(self, owner_key, job_id):
        payload=self._read(job_id)
        return None if not payload or payload["owner_key"] != owner_key else {k:v for k,v in payload.items() if k != "owner_key"}

    def _latest(self, owner_key, fingerprint=None):
        rows=self.db.execute("SELECT owner_key,payload FROM collection_jobs ORDER BY rowid").fetchall(); matches=[]
        for owner,raw in rows:
            if owner != owner_key: continue
            payload=json.loads(raw)
            if payload.get("status") not in {"running","error","complete"}: continue
            if fingerprint is not None and payload.get("fingerprint") != fingerprint: continue
            matches.append(payload)
        return matches[-1] if matches else None

    def latest_resumable(self, owner_key): return self._latest(owner_key)
    def latest_for_fingerprint(self, owner_key, fingerprint): return self._latest(owner_key, fingerprint)
