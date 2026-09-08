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
        url = os.getenv("SUPABASE_COLLECTION_URL", DEFAULT_SUPABASE_URL).strip()
        key = os.getenv("SUPABASE_COLLECTION_PUBLISHABLE_KEY", DEFAULT_SUPABASE_PUBLISHABLE_KEY).strip()
        if not url or not key:
            return None
        return cls(url, key)

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
