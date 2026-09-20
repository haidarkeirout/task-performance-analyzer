"""Synchronous facade over the durable Jira ``CollectionJob``.

The original Employee collector renders the job incrementally.  Project,
Department, and Company collection are synchronous Streamlit flows, so this
module drives the same persisted job to completion while preserving its
per-work-item checkpoints and automatic retries.
"""
from __future__ import annotations

import time
from datetime import datetime, timezone
from typing import Any, Callable

from collection_job import CollectionJob
from collection_store import CollectionStore, PersistenceError, persistence_owner_key
from jira_gateway import CollectionError, JiraGateway


Progress = Callable[[float, str], None]


def _recent_completed_payload(payload: dict[str, Any]) -> bool:
    if payload.get("status") != "complete":
        return False
    value = (payload.get("result_meta") or {}).get("collected_at")
    try:
        collected = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
        if collected.tzinfo is None:
            collected = collected.replace(tzinfo=timezone.utc)
    except (TypeError, ValueError):
        return False
    return (datetime.now(timezone.utc) - collected).total_seconds() <= 1800


def _persisted_issue_keys(payload: dict[str, Any] | None) -> set[str]:
    """Return the seeded issue keys stored with a persisted collection."""
    if not isinstance(payload, dict):
        return set()
    keys: set[str] = set()
    for item in payload.get("items") or ():
        if not isinstance(item, dict):
            continue
        seed = item.get("seed_item") or {}
        key = item.get("issue_key") or (seed.get("key") if isinstance(seed, dict) else None)
        if key:
            keys.add(str(key))
    return keys


def collect_jira_query(
    settings,
    *,
    space: dict[str, Any],
    query: str,
    fingerprint: str,
    cutoff: str | None = None,
    seed_issues: list[dict[str, Any]] | None = None,
    progress: Progress | None = None,
):
    """Collect one Jira query using durable per-task checkpoints.

    Retrying the same fingerprint restores completed work items from Supabase.
    ``space['id'] == '*'`` is a deliberate multi-project scope used by Employee
    collection; normal Project/Department jobs still verify project identity.
    """
    store = CollectionStore.configured()
    if store is None:
        raise CollectionError(
            "Persistent Jira collection storage is not configured. "
            "Configure the collection store before collecting analysis data."
        )
    owner_key = persistence_owner_key(settings)
    keep_store = True
    try:
        payload = store.latest_for_fingerprint(owner_key, fingerprint)
        # Employee collection supplies a freshly discovered issue list. Never
        # resume a persisted job whose seeded list belongs to a different
        # snapshot; doing so can make collection_save_item reject a valid issue
        # as "not part of this collection".
        if payload and seed_issues is not None:
            current_keys = {str(item.get("key")) for item in seed_issues if item.get("key")}
            persisted_keys = _persisted_issue_keys(payload)
            if persisted_keys != current_keys:
                payload = None
        job = None
        if payload and (
            payload.get("status") in {"running", "error", "paused"}
            or _recent_completed_payload(payload)
        ):
            job = CollectionJob.from_persisted(
                settings,
                payload,
                JiraGateway,
                store=store,
                owner_key=owner_key,
            )

        if job is None:
            gateway = JiraGateway(settings)
            try:
                definitions = gateway.fields()
            finally:
                gateway.close()
            job = CollectionJob(
                settings,
                space,
                query,
                fingerprint,
                definitions,
                JiraGateway,
                store=store,
                owner_key=owner_key,
            )
            job.checkpoint["cutoff"] = cutoff or datetime.now(timezone.utc).isoformat()
            if seed_issues is not None:
                job.checkpoint["issues"] = list(seed_issues)
                job.checkpoint["completed"] = {}

        job.start()
        if job.error:
            raise CollectionError(job.error, retryable=True)

        if seed_issues is not None:
            # Idempotently repair/reassert the ordered item list before the
            # first item is saved. This also repairs old jobs whose item rows
            # were only partially seeded before an interrupted request.
            store.seed_items(owner_key, job.id, list(seed_issues))

        while True:
            job.step()
            snapshot = job.snapshot()
            if progress:
                total = snapshot["total"]
                completed = snapshot["completed"]
                progress(completed / total if total else 0.0, snapshot["message"])
            if snapshot["result"] is not None:
                keep_store = False
                store.close()
                return snapshot["result"]
            if snapshot["error"]:
                detail = snapshot.get("error_detail") or {}
                reference = detail.get("reference") or snapshot.get("id")
                raise CollectionError(
                    f"{snapshot['error']} Saved collection reference: {reference}.",
                    status=detail.get("http_status"),
                    retryable=True,
                )
            if snapshot.get("retry_attempt"):
                time.sleep(0.25)
    except PersistenceError as exc:
        raise CollectionError(
            "The persistent collection checkpoint could not be used. "
            f"No saved work was discarded. {exc}",
            retryable=exc.retryable,
        ) from exc
    finally:
        if keep_store:
            store.close()
