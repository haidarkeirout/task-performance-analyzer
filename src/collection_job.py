"""Resumable Jira collection; one bounded step per Streamlit rerun."""
import logging, time, traceback, uuid
from datetime import datetime, timezone
from jira_export import PreparedData, _json, build_workbook
from jira_gateway import CollectionError

LOGGER = logging.getLogger("jira_collection")
LOGGER.setLevel(logging.INFO)
if not LOGGER.handlers:
    h = logging.StreamHandler(); h.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s")); LOGGER.addHandler(h)
LOGGER.propagate = False

class CollectionCancelled(Exception): pass

class CollectionJob:
    def __init__(self, settings, space, query, fingerprint, definitions, gateway_factory, store=None, owner_key=None):
        self.settings, self.space, self.query = settings, dict(space), query
        self.fingerprint, self.definitions, self.gateway_factory = fingerprint, list(definitions), gateway_factory
        # Optional persistence hook; existing Jira behavior is unchanged when omitted.
        self.store = store
        self.owner_key = owner_key
        self.checkpoint, self.id = {}, uuid.uuid4().hex[:12]
        self.cancelled = False; self.running = False; self.result = None; self.error = None
        self.message = "Ready to collect."; self.started = None
    @classmethod
    def from_persisted(cls, settings, payload, gateway_factory, store=None, owner_key=None):
        job = cls(settings, payload["space"], payload["query"], payload["fingerprint"],
                  payload.get("definitions", []), gateway_factory, store=store, owner_key=owner_key)
        job.id = payload["job_id"]
        job.message = payload.get("message", "Restored collection.")
        job.checkpoint = {"cutoff": payload["cutoff"], "fingerprint": payload["fingerprint"],
                          "issues": [row["seed_item"] for row in payload.get("items", [])]}
        job.checkpoint["completed"] = {}
        for row in payload.get("items", []):
            if row.get("completed"):
                job.checkpoint["completed"][row["issue_key"]] = (row["item_data"], row["history_data"])
        return job

    def _persist_update(self, status=None, stage=None, current_issue=None, error_detail=None, result_meta=None):
        if self.store is not None and self.owner_key is not None:
            self.store.update_job(self.owner_key, self.id, status or ("running" if self.running else "error"),
                                  stage or self.message, current_issue, self.message, error_detail, result_meta)

    def cancel(self): self.cancelled = True; self.running = False
    def snapshot(self):
        return dict(running=self.running, result=self.result, error=self.error, message=self.message, id=self.id,
                    completed=len(self.checkpoint.get("completed", {})), total=len(self.checkpoint.get("issues", [])),
                    elapsed=int(time.monotonic()-self.started) if self.started else 0)
    def progress(self, message):
        if self.cancelled: raise CollectionCancelled()
        self.message = message; LOGGER.info("collection=%s phase=%s", self.id, message)
    def start(self):
        if self.running or self.cancelled or self.result is not None: return
        if self.started and time.monotonic()-self.started > 1800: self.checkpoint.clear(); self.started = None
        self.started = self.started or time.monotonic(); self.error = None; self.running = True
    def _fail(self, exc):
        LOGGER.error("collection=%s stopped phase=%s exception=%s", self.id, self.message, type(exc).__name__)
        loc = " -> ".join(f"{f.name}:{f.lineno}" for f in traceback.extract_tb(exc.__traceback__))
        LOGGER.error("collection=%s locations=%s", self.id, loc)
        reason = str(exc) if isinstance(exc, CollectionError) else "An unexpected collection error occurred."
        self.error = f"{reason} Stage: {self.message} Reference: {self.id}. Completed work is retained; click Done to retry."; self.running = False
    def step(self):
        if not self.running or self.cancelled or self.result is not None: return
        gateway = None
        try:
            gateway = self.gateway_factory(self.settings); gateway.progress = self.progress
            cutoff = self.checkpoint.setdefault("cutoff", datetime.now(timezone.utc).isoformat())
            self.checkpoint.setdefault("fingerprint", self.fingerprint); self.checkpoint.setdefault("completed", {})
            if "issues" not in self.checkpoint:
                self.progress("Reading the Jira work item list..."); issues = gateway.all_issues(self.query, progress=self.progress)
                keys = [i.get("key") for i in issues]
                if not issues: raise CollectionError("No work items match these filters. Change the filters and click Done again.")
                if any(not k for k in keys) or len(set(keys)) != len(keys): raise CollectionError("The selected work item list is incomplete or contains duplicates.")
                self.checkpoint["issues"] = issues
                if self.store is not None and self.owner_key is not None:
                    self.store.begin(self.owner_key, self.id, self.fingerprint, self.space, self.query, self.definitions, cutoff)
                    self.store.seed_items(self.owner_key, self.id, issues)
            issues = self.checkpoint["issues"]
            if "project" not in self.checkpoint:
                self.progress("Reading project details..."); self.checkpoint["project"] = gateway.project_details(self.space["id"]) if hasattr(gateway, "project_details") else self.space
            completed = self.checkpoint["completed"]; pending = next((i for i in issues if i["key"] not in completed), None)
            if pending is not None:
                self.progress(f"Completed {len(completed)} of {len(issues)}; reading {pending['key']}..."); item, history = gateway.complete_issue(pending)
                actual = str((item.get("fields", {}).get("project") or {}).get("id", ""))
                if actual != str(self.space["id"]) or item.get("key") != pending["key"]: raise CollectionError("A work item changed spaces or keys during collection. Start a new collection.")
                if history.get("history_complete") is not True or not history.get("history_through"): raise CollectionError("A task history is incomplete. No partial export was prepared.")
                if datetime.fromisoformat(history["history_through"].replace("Z", "+00:00")) < datetime.fromisoformat(cutoff.replace("Z", "+00:00")): raise CollectionError("A task history does not cover the evaluation cutoff.")
                item["fields"]["project"] = {**self.checkpoint["project"], **(item["fields"].get("project") or {})}; completed[pending["key"]] = (item, history)
                if self.store is not None and self.owner_key is not None:
                    self.store.save_item(self.owner_key, self.id, pending["key"], item, history)
                self.progress(f"Completed {len(completed)} of {len(issues)} work items."); return
            self.progress("Preparing your Excel file..."); full = [completed[i["key"]][0] for i in issues]; histories = {i["key"]: completed[i["key"]][1] for i in issues}; collected_at = datetime.now(timezone.utc).isoformat()
            data = build_workbook(full, histories, self.definitions, cutoff=cutoff, collected_at=collected_at, query=self.query, space_name=self.space["name"], source_timezone=self.settings.source_timezone, preferred_start=self.settings.start_date_field)
            filename = f"Jira_{self.space['key']}_{datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S')}.xlsx"
            self.result = PreparedData(data, _json(histories).encode(), cutoff, collected_at, self.query, self.fingerprint, len(full), filename, self.space["name"], self.settings.source_timezone); self.running = False; self.progress("Excel file prepared. All selected work items have been collected.")
        except CollectionCancelled: self.running = False
        except BaseException as exc: self._fail(exc)
        finally:
            if gateway is not None:
                try: gateway.close()
                except Exception: pass
