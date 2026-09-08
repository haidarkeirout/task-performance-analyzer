"""Persistent Jira collection checkpoints backed by Supabase PostgREST RPCs.

The publishable Supabase key is intentionally safe to ship with the server code.
Actual access to a user's collection is capability-scoped by a 256-bit owner key
that is derived from the server-side application login secret and never sent to
the browser.
"""
from __future__ import annotations

import hashlib
import os
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
    def __init__(self, url: str, publishable_key: str, *, session=None, timeout=(10, 45)):
        self.url = url.rstrip("/")
        self.key = publishable_key
        self.session = session or requests.Session()
        self.timeout = timeout
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

    def _rpc(self, name: str, payload: dict[str, Any]):
        try:
            response = self.session.post(
                f"{self.url}/rest/v1/rpc/{name}", json=payload,
                timeout=self.timeout, allow_redirects=False,
            )
        except requests.RequestException as exc:
            raise PersistenceError(
                "The persistent collection store could not be reached. Your Jira request was not marked complete."
            ) from exc

        if response.status_code < 200 or response.status_code >= 300:
            retryable = response.status_code == 429 or response.status_code >= 500
            try:
                body = response.json()
                detail = body.get("message") or body.get("details") or body.get("hint")
            except ValueError:
                detail = None
            safe = f" ({detail})" if detail and len(str(detail)) < 240 else ""
            raise PersistenceError(
                f"The persistent collection store rejected the request (HTTP {response.status_code}).{safe}",
                retryable=retryable,
            )
        if not response.content:
            return None
        try:
            return response.json()
        except ValueError as exc:
            raise PersistenceError("The persistent collection store returned an unreadable response.") from exc

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
