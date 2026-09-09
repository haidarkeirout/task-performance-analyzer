"""Resumable Jira collection with persistent Supabase checkpoints."""
from __future__ import annotations

import logging
import time
import traceback
import uuid
from datetime import datetime, timezone

from collection_store import CollectionStore, PersistenceError, persistence_owner_key
from jira_export import PreparedData, _json, build_workbook
from jira_gateway import CollectionError


LOGGER = logging.getLogger("jira_collection")
LOGGER.setLevel(logging.INFO)
if not LOGGER.handlers:
    handler = logging.StreamHandler()
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    LOGGER.addHandler(handler)
LOGGER.propagate = False


class CollectionCancelled(Exception):
    pass


class CollectionJob:
    MAX_AUTO_RETRIES = 3

    def __init__(self, settings, space, query, fingerprint, definitions, gateway_factory,
                 *, store=None, owner_key=None, job_id=None, restored=False):
        self.settings = settings
        self.space = dict(space)
        self.query = query
        self.fingerprint = fingerprint
        self.definitions = list(definitions)
        self.gateway_factory = gateway_factory
        self.store = store or CollectionStore.configured()
        self.owner_key = owner_key or persistence_owner_key(settings)

        self.checkpoint = {}
        self.id = job_id or uuid.uuid4().hex[:12]
        self.cancelled = False
        self.running = False
        self.result = None
        self.error = None
        self.error_detail = None
        self.message = "Ready to collect."
        self.stage = "Ready"
        self.current_issue = None
        self.last_successful_issue = None
        self.started = None
        self.restored = restored
        self.persisted_status = "running" if not restored else "paused"

        self.retry_counts = {}
        self.retry_after = 0.0
        self.retry_key = None

    @classmethod
    def from_persisted(cls, settings, payload, gateway_factory, *, store=None, owner_key=None):
        if not payload:
            return None
        job = cls(
            settings,
            payload.get("space") or {},
            payload.get("query") or "",
            payload.get("fingerprint") or "",
            payload.get("definitions") or [],
            gateway_factory,
            store=store,
            owner_key=owner_key,
            job_id=payload.get("job_id"),
            restored=True,
        )
        job.checkpoint["cutoff"] = payload.get("cutoff")
        issues, completed = [], {}
        for row in sorted(payload.get("items") or [], key=lambda r: r.get("position", 0)):
            seed = row.get("seed_item") or {}
            issues.append(seed)
            if row.get("completed") and row.get("item_data") is not None and row.get("history_data") is not None:
                completed[row["issue_key"]] = (row["item_data"], row["history_data"])
        if issues:
            job.checkpoint["issues"] = issues
        job.checkpoint["completed"] = completed
        job.stage = payload.get("stage") or "Ready"
        job.current_issue = payload.get("current_issue")
        job.last_successful_issue = payload.get("last_successful_issue")
        job.message = payload.get("message") or "Previous collection restored."
        job.error_detail = payload.get("error_detail")
        job.persisted_status = payload.get("status") or "running"
        if job.persisted_status == "error":
            job.error = (job.error_detail or {}).get("message") or "The previous collection stopped with an error."
        elif job.persisted_status == "complete":
            job.message = "Completed persistent collection restored. Rebuilding the analysis source..."
        else:
            job.message = (
                f"Previous collection restored with {len(completed)} of {len(issues)} work items saved."
                if issues else "Previous collection restored."
            )
        return job

    def _persist_update(self, status=None, *, error_detail=None, result_meta=None):
        if self.store is None:
            raise PersistenceError("Persistent collection storage is not configured.", retryable=False)
        self.store.update_job(
            self.owner_key,
            self.id,
            status or ("running" if self.running else self.persisted_status),
            self.stage,
            self.current_issue,
            self.message,
            error_detail,
            result_meta,
        )
        self.persisted_status = status or self.persisted_status

    def cancel(self, persist=True):
        self.cancelled = True
        self.running = False
        self.persisted_status = "cancelled"
        if persist and self.store is not None:
            try:
                self._persist_update("cancelled")
            except Exception:
                LOGGER.exception("collection=%s failed_to_persist_cancel", self.id)
        LOGGER.info("collection=%s cancelled completed=%s total=%s", self.id,
                    len(self.checkpoint.get("completed", {})),
                    len(self.checkpoint.get("issues", [])))

    def snapshot(self):
        if self.result is not None:
            status = "complete"
        elif self.error:
            status = "error"
        elif self.running:
            status = "running"
        elif self.restored and self.persisted_status in {"running", "complete"}:
            status = "paused"
        elif self.cancelled:
            status = "cancelled"
        else:
            status = "idle"
        return {
            "status": status,
            "running": self.running,
            "result": self.result,
            "error": self.error,
            "error_detail": self.error_detail,
            "message": self.message,
            "stage": self.stage,
            "current_issue": self.current_issue,
            "last_successful_issue": self.last_successful_issue,
            "id": self.id,
            "completed": len(self.checkpoint.get("completed", {})),
            "total": len(self.checkpoint.get("issues", [])),
            "elapsed": int(time.monotonic() - self.started) if self.started else 0,
            "retry_attempt": self.retry_counts.get(self.retry_key, 0) if self.retry_key else 0,
            "max_auto_retries": self.MAX_AUTO_RETRIES,
            "restored": self.restored,
            "persisted_status": self.persisted_status,
        }

    def progress(self, message):
        if self.cancelled:
            raise CollectionCancelled()
        self.message = message
        LOGGER.info("collection=%s stage=%s issue=%s phase=%s", self.id, self.stage,
                    self.current_issue or "-", message)

    def _set_stage(self, stage, message, issue=None):
        self.stage = stage
        self.current_issue = issue
        self.progress(message)

    def _retry_identity(self):
        return self.current_issue or f"stage:{self.stage}"

    def _clear_retry(self):
        key = self._retry_identity()
        self.retry_counts.pop(key, None)
        if self.retry_key == key:
            self.retry_key = None
            self.retry_after = 0.0

    def _ensure_persistent_job(self):
        if self.store is None:
            raise PersistenceError("Persistent collection storage is not configured.", retryable=False)
        cutoff = self.checkpoint.setdefault("cutoff", datetime.now(timezone.utc).isoformat())
        self.store.begin(
            self.owner_key,
            self.id,
            self.fingerprint,
            self.space,
            self.query,
            self.definitions,
            cutoff,
        )

    def start(self):
        if self.running or self.cancelled or self.result is not None:
            return
        if self.error:
            key = self.current_issue or f"stage:{self.stage}"
            self.retry_counts.pop(key, None)
            self.retry_key = None
            self.retry_after = 0.0
        self.started = self.started or time.monotonic()
        self.error = None
        self.error_detail = None
        self.running = True
        self.restored = False
        self.persisted_status = "running"
        try:
            self._ensure_persistent_job()
            self._persist_update("running")
        except BaseException as exc:
            self.running = False
            self._fail(exc, persist=False)
            return
        LOGGER.info("collection=%s started_or_resumed completed=%s total=%s", self.id,
                    len(self.checkpoint.get("completed", {})),
                    len(self.checkpoint.get("issues", [])))

    @staticmethod
    def _is_retryable(exc):
        if isinstance(exc, PersistenceError):
            return exc.retryable
        if not isinstance(exc, CollectionError):
            return False
        if getattr(exc, "retryable", False):
            return True
        status = getattr(exc, "status", None)
        if status == 429 or (isinstance(status, int) and 500 <= status <= 599):
            return True
        text = str(exc).casefold()
        transient_markers = (
            "could not reach jira", "jira is busy", "limited requests", "try again later",
            "unreadable response", "changed while it was being collected",
            "ended before all records were collected", "repeated or skipped a page",
        )
        return any(marker in text for marker in transient_markers)

    def _schedule_auto_retry(self, exc):
        if not self._is_retryable(exc):
            return False
        key = self._retry_identity()
        attempt = self.retry_counts.get(key, 0) + 1
        if attempt > self.MAX_AUTO_RETRIES:
            return False
        self.retry_counts[key] = attempt
        self.retry_key = key
        delay = min(2 ** (attempt - 1), 8)
        self.retry_after = time.monotonic() + delay
        reason = str(exc).strip() or type(exc).__name__
        target = self.current_issue or self.stage
        self.message = (
            f"Temporary problem while collecting {target}. Automatic retry "
            f"{attempt} of {self.MAX_AUTO_RETRIES} in {delay} second(s). {reason}"
        )
        LOGGER.warning("collection=%s auto_retry=%s/%s stage=%s issue=%s delay=%ss error=%s",
                       self.id, attempt, self.MAX_AUTO_RETRIES, self.stage,
                       self.current_issue or "-", delay, reason)
        try:
            self._persist_update("running")
        except Exception:
            LOGGER.exception("collection=%s failed_to_persist_retry_state", self.id)
        return True

    def _fail(self, exc, *, persist=True):
        LOGGER.error("collection=%s stopped stage=%s issue=%s exception=%s", self.id,
                     self.stage, self.current_issue or "-", type(exc).__name__)
        location = " -> ".join(
            f"{frame.name}:{frame.lineno}" for frame in traceback.extract_tb(exc.__traceback__)
        )
        LOGGER.error("collection=%s locations=%s", self.id, location or "unknown")
        if isinstance(exc, (CollectionError, PersistenceError)):
            reason = str(exc).strip() or "The collection could not continue."
            status = getattr(exc, "status", None)
        else:
            reason = f"Unexpected internal error ({type(exc).__name__})."
            status = None
        completed = len(self.checkpoint.get("completed", {}))
        total = len(self.checkpoint.get("issues", []))
        self.error = reason
        self.error_detail = {
            "type": type(exc).__name__, "message": reason, "stage": self.stage,
            "failed_issue": self.current_issue, "completed": completed, "total": total,
            "http_status": status, "reference": self.id,
        }
        self.message = (
            f"Collection stopped at {self.stage}. {completed} of {total} work items are saved persistently."
            if total else f"Collection stopped at {self.stage}."
        )
        self.running = False
        self.persisted_status = "error"
        if persist and self.store is not None:
            try:
                self._persist_update("error", error_detail=self.error_detail)
            except Exception:
                LOGGER.exception("collection=%s failed_to_persist_error", self.id)

    def _prepare_result(self):
        issues = self.checkpoint.get("issues", [])
        completed = self.checkpoint.get("completed", {})
        full = [completed[item["key"]][0] for item in issues]
        histories = {item["key"]: completed[item["key"]][1] for item in issues}
        cutoff = self.checkpoint["cutoff"]
        collected_at = datetime.now(timezone.utc).isoformat()
        data = build_workbook(
            full, histories, self.definitions, cutoff=cutoff, collected_at=collected_at,
            query=self.query, space_name=self.space["name"],
            source_timezone=self.settings.source_timezone,
            preferred_start=self.settings.start_date_field,
        )
        filename = f"Jira_{self.space['key']}_{datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S')}.xlsx"
        self.result = PreparedData(
            data, _json(histories).encode(), cutoff, collected_at, self.query,
            self.fingerprint, len(full), filename, self.space["name"],
            self.settings.source_timezone,
        )
        self.running = False
        self.current_issue = None
        self.stage = "Complete"
        self.persisted_status = "complete"
        self.message = "Excel file prepared. All selected work items have been collected."
        self._persist_update("complete", result_meta={
            "filename": filename, "count": len(full), "collected_at": collected_at,
        })
        LOGGER.info("collection=%s complete total=%s", self.id, len(full))

    def step(self):
        if not self.running or self.cancelled or self.result is not None:
            return
        if self.retry_after and time.monotonic() < self.retry_after:
            return
        if self.retry_after:
            self.retry_after = 0.0

        gateway = None
        try:
            gateway = self.gateway_factory(self.settings)
            gateway.progress = self.progress
            cutoff = self.checkpoint.setdefault("cutoff", datetime.now(timezone.utc).isoformat())
            self.checkpoint.setdefault("fingerprint", self.fingerprint)
            self.checkpoint.setdefault("completed", {})

            if "issues" not in self.checkpoint:
                self._set_stage("Reading work item list", "Reading the Jira work item list...")
                issues = gateway.all_issues(self.query, progress=self.progress)
                keys = [item.get("key") for item in issues]
                if not issues:
                    raise CollectionError("No work items match these filters. Change the filters and start a new collection.")
                if any(not key for key in keys) or len(set(keys)) != len(keys):
                    raise CollectionError("The selected work item list is incomplete or contains duplicates.")
                self.store.seed_items(self.owner_key, self.id, issues)
                self.checkpoint["issues"] = issues
                self._clear_retry()
                self.stage = "Collecting work items"
                self.message = f"{len(issues)} work items are checkpointed persistently."
                self._persist_update("running")

            issues = self.checkpoint["issues"]
            completed = self.checkpoint["completed"]

            if "project" not in self.checkpoint:
                self._set_stage("Reading project details", "Reading project details...")
                self.checkpoint["project"] = (
                    gateway.project_details(self.space["id"])
                    if hasattr(gateway, "project_details") else self.space
                )
                self._clear_retry()

            pending = next((item for item in issues if item["key"] not in completed), None)
            if pending is not None:
                issue_key = pending["key"]
                self._set_stage(
                    "Reading work item",
                    f"Completed {len(completed)} of {len(issues)}; reading {issue_key}...",
                    issue=issue_key,
                )
                self._persist_update("running")
                item, history = gateway.complete_issue(pending)
                actual_project = str((item.get("fields", {}).get("project") or {}).get("id", ""))
                if actual_project != str(self.space["id"]) or item.get("key") != issue_key:
                    raise CollectionError("A work item changed spaces or keys during collection. Start a new collection.")
                if history.get("history_complete") is not True or not history.get("history_through"):
                    raise CollectionError("A task history is incomplete. Jira did not provide a complete history snapshot.")
                history_through = datetime.fromisoformat(history["history_through"].replace("Z", "+00:00"))
                cutoff_dt = datetime.fromisoformat(str(cutoff).replace("Z", "+00:00"))
                if history_through < cutoff_dt:
                    raise CollectionError("A task history does not cover the evaluation cutoff. Please retry the collection.")
                item["fields"]["project"] = {
                    **self.checkpoint["project"], **(item["fields"].get("project") or {}),
                }

                # Durable write happens before this item is acknowledged in RAM.
                self.store.save_item(self.owner_key, self.id, issue_key, item, history)
                completed[issue_key] = (item, history)
                self.last_successful_issue = issue_key
                self._clear_retry()
                self.current_issue = None
                self.stage = "Collecting work items"
                self.message = f"Completed {len(completed)} of {len(issues)} work items."
                return

            self._set_stage("Preparing Excel file", "Preparing your Excel file...")
            self._prepare_result()

        except CollectionCancelled:
            self.running = False
        except BaseException as exc:
            if not self._schedule_auto_retry(exc):
                self._fail(exc)
        finally:
            if gateway is not None:
                try:
                    gateway.close()
                except Exception:
                    pass
