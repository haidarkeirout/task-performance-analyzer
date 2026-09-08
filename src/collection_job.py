"""Resumable Jira collection with durable in-session checkpoints and automatic retries."""
from __future__ import annotations

import logging
import time
import traceback
import uuid
from datetime import datetime, timezone

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
    """Collect Jira data incrementally without losing completed work on reruns.

    A Streamlit fragment calls :meth:`step` repeatedly. Each successful work item is
    stored in ``checkpoint['completed']`` immediately, so ordinary Streamlit reruns
    and recoverable Jira failures never restart the collection from task 1.
    """

    MAX_AUTO_RETRIES = 3

    def __init__(self, settings, space, query, fingerprint, definitions, gateway_factory):
        self.settings = settings
        self.space = dict(space)
        self.query = query
        self.fingerprint = fingerprint
        self.definitions = list(definitions)
        self.gateway_factory = gateway_factory

        self.checkpoint = {}
        self.id = uuid.uuid4().hex[:12]
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

        self.retry_counts = {}
        self.retry_after = 0.0
        self.retry_key = None

    def cancel(self):
        self.cancelled = True
        self.running = False
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

    def start(self):
        """Start or manually resume the same checkpoint.

        Completed work is intentionally never expired by elapsed time. Large Jira
        spaces can take more than 30 minutes, and a manual retry must resume from
        the first unfinished item rather than clearing the checkpoint.
        """
        if self.running or self.cancelled or self.result is not None:
            return

        # A manual retry gets a fresh automatic-retry budget for the failed stage,
        # while the already completed Jira items remain untouched.
        if self.error:
            key = self.current_issue or f"stage:{self.stage}"
            self.retry_counts.pop(key, None)
            self.retry_key = None
            self.retry_after = 0.0

        self.started = self.started or time.monotonic()
        self.error = None
        self.error_detail = None
        self.running = True
        LOGGER.info("collection=%s started_or_resumed completed=%s total=%s", self.id,
                    len(self.checkpoint.get("completed", {})),
                    len(self.checkpoint.get("issues", [])))

    @staticmethod
    def _is_retryable(exc):
        if not isinstance(exc, CollectionError):
            return False
        if getattr(exc, "retryable", False):
            return True
        status = getattr(exc, "status", None)
        if status == 429 or (isinstance(status, int) and 500 <= status <= 599):
            return True
        text = str(exc).casefold()
        transient_markers = (
            "could not reach jira",
            "jira is busy",
            "limited requests",
            "try again later",
            "unreadable response",
            "changed while it was being collected",
            "ended before all records were collected",
            "repeated or skipped a page",
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
        LOGGER.warning(
            "collection=%s auto_retry=%s/%s stage=%s issue=%s delay=%ss error=%s",
            self.id, attempt, self.MAX_AUTO_RETRIES, self.stage,
            self.current_issue or "-", delay, reason,
        )
        return True

    def _fail(self, exc):
        LOGGER.error("collection=%s stopped stage=%s issue=%s exception=%s", self.id,
                     self.stage, self.current_issue or "-", type(exc).__name__)
        location = " -> ".join(
            f"{frame.name}:{frame.lineno}" for frame in traceback.extract_tb(exc.__traceback__)
        )
        LOGGER.error("collection=%s locations=%s", self.id, location or "unknown")

        if isinstance(exc, CollectionError):
            reason = str(exc).strip() or "Jira could not complete the collection request."
            status = getattr(exc, "status", None)
        else:
            reason = f"Unexpected internal error ({type(exc).__name__})."
            status = None

        completed = len(self.checkpoint.get("completed", {}))
        total = len(self.checkpoint.get("issues", []))
        self.error = reason
        self.error_detail = {
            "type": type(exc).__name__,
            "message": reason,
            "stage": self.stage,
            "failed_issue": self.current_issue,
            "completed": completed,
            "total": total,
            "http_status": status,
            "reference": self.id,
        }
        self.message = (
            f"Collection stopped at {self.stage}. {completed} of {total} work items are saved."
            if total else f"Collection stopped at {self.stage}."
        )
        self.running = False

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
                    raise CollectionError(
                        "No work items match these filters. Change the filters and start a new collection."
                    )
                if any(not key for key in keys) or len(set(keys)) != len(keys):
                    raise CollectionError(
                        "The selected work item list is incomplete or contains duplicates."
                    )
                self.checkpoint["issues"] = issues
                self._clear_retry()

            issues = self.checkpoint["issues"]

            if "project" not in self.checkpoint:
                self._set_stage("Reading project details", "Reading project details...")
                self.checkpoint["project"] = (
                    gateway.project_details(self.space["id"])
                    if hasattr(gateway, "project_details") else self.space
                )
                self._clear_retry()

            completed = self.checkpoint["completed"]
            pending = next((item for item in issues if item["key"] not in completed), None)

            if pending is not None:
                issue_key = pending["key"]
                self._set_stage(
                    "Reading work item",
                    f"Completed {len(completed)} of {len(issues)}; reading {issue_key}...",
                    issue=issue_key,
                )
                item, history = gateway.complete_issue(pending)

                actual_project = str((item.get("fields", {}).get("project") or {}).get("id", ""))
                if actual_project != str(self.space["id"]) or item.get("key") != issue_key:
                    raise CollectionError(
                        "A work item changed spaces or keys during collection. Start a new collection."
                    )
                if history.get("history_complete") is not True or not history.get("history_through"):
                    raise CollectionError(
                        "A task history is incomplete. Jira did not provide a complete history snapshot."
                    )
                history_through = datetime.fromisoformat(
                    history["history_through"].replace("Z", "+00:00")
                )
                cutoff_dt = datetime.fromisoformat(cutoff.replace("Z", "+00:00"))
                if history_through < cutoff_dt:
                    raise CollectionError(
                        "A task history does not cover the evaluation cutoff. Please retry the collection."
                    )

                item["fields"]["project"] = {
                    **self.checkpoint["project"],
                    **(item["fields"].get("project") or {}),
                }
                completed[issue_key] = (item, history)
                self.last_successful_issue = issue_key
                self._clear_retry()
                self.current_issue = None
                self.stage = "Collecting work items"
                self.progress(f"Completed {len(completed)} of {len(issues)} work items.")
                return

            self._set_stage("Preparing Excel file", "Preparing your Excel file...")
            full = [completed[item["key"]][0] for item in issues]
            histories = {item["key"]: completed[item["key"]][1] for item in issues}
            collected_at = datetime.now(timezone.utc).isoformat()
            data = build_workbook(
                full,
                histories,
                self.definitions,
                cutoff=cutoff,
                collected_at=collected_at,
                query=self.query,
                space_name=self.space["name"],
                source_timezone=self.settings.source_timezone,
                preferred_start=self.settings.start_date_field,
            )
            filename = (
                f"Jira_{self.space['key']}_"
                f"{datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S')}.xlsx"
            )
            self.result = PreparedData(
                data,
                _json(histories).encode(),
                cutoff,
                collected_at,
                self.query,
                self.fingerprint,
                len(full),
                filename,
                self.space["name"],
                self.settings.source_timezone,
            )
            self._clear_retry()
            self.running = False
            self.current_issue = None
            self.stage = "Complete"
            self.progress("Excel file prepared. All selected work items have been collected.")
            LOGGER.info("collection=%s complete total=%s", self.id, len(full))

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
