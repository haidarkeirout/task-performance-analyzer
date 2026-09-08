"""Session-owned collection worker. No Streamlit calls run in this thread."""
import logging
import threading
import time
import uuid
import traceback

from jira_export import collect_data
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
    def __init__(self, settings, space, query, fingerprint, definitions, gateway_factory):
        self.settings, self.space, self.query = settings, dict(space), query
        self.fingerprint, self.definitions = fingerprint, list(definitions)
        self.gateway_factory = gateway_factory
        self.checkpoint = {}
        self.id = uuid.uuid4().hex[:12]
        self.lock = threading.RLock()
        self.cancelled = threading.Event()
        self.thread = None
        self.running = False
        self.result = None
        self.error = None
        self.message = "Ready to collect."
        self.started = None

    def cancel(self):
        self.cancelled.set()

    def snapshot(self):
        with self.lock:
            return dict(running=self.running, result=self.result, error=self.error,
                        message=self.message, id=self.id,
                        completed=len(self.checkpoint.get("completed", {})),
                        total=len(self.checkpoint.get("issues", [])),
                        elapsed=int(time.monotonic() - self.started) if self.started else 0)

    def progress(self, message):
        if self.cancelled.is_set():
            raise CollectionCancelled()
        with self.lock:
            self.message = message
        LOGGER.info("collection=%s phase=%s", self.id, message)

    def start(self):
        with self.lock:
            if self.running or self.cancelled.is_set() or self.result is not None:
                return
            # Never reuse an old day's partial snapshots for a new evaluation.
            if self.started and time.monotonic() - self.started > 1800:
                self.checkpoint.clear()
                self.started = None
            self.started = self.started or time.monotonic()
            self.error = None
            self.running = True
            self.thread = threading.Thread(target=self._run, daemon=True, name="jira-collection-" + self.id)
            self.thread.start()

    def _run(self):
        gateway = None
        try:
            gateway = self.gateway_factory(self.settings)
            gateway.progress = self.progress
            result = collect_data(gateway, self.space, self.query, self.fingerprint,
                                  self.definitions, progress=self.progress, checkpoint=self.checkpoint)
            self.progress("Excel file prepared. All selected work items have been collected.")
            with self.lock:
                self.result = result
        except CollectionCancelled:
            LOGGER.info("collection=%s cancelled", self.id)
        except BaseException as exc:
            # Log type and our own stage only: HTTP bodies can contain private data.
            LOGGER.error("collection=%s stopped phase=%s exception=%s", self.id, self.message, type(exc).__name__)
            locations = [f"{frame.name}:{frame.lineno}" for frame in traceback.extract_tb(exc.__traceback__)]
            LOGGER.error("collection=%s locations=%s", self.id, " -> ".join(locations))
            with self.lock:
                reason = str(exc) if isinstance(exc, CollectionError) else "An unexpected collection error occurred."
                self.error = f"{reason} Stage: {self.message} Reference: {self.id}. Completed work is retained; click Done to retry."
        finally:
            if gateway is not None:
                try:
                    gateway.close()
                except Exception:
                    pass
            with self.lock:
                self.running = False
