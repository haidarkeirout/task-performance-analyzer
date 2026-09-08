"""Exercise the Streamlit flow with deterministic Jira and persistent-store fakes."""
import copy
from datetime import datetime, timezone
import unittest
from unittest.mock import patch

from streamlit.testing.v1 import AppTest

from test_automation import ROOT, FIELDS, example_issue, no_events
import automation_ui


class FakeStore:
    jobs = {}

    @classmethod
    def reset(cls):
        cls.jobs = {}

    def close(self):
        pass

    def begin(self, owner_key, job_id, fingerprint, space, query, definitions, cutoff):
        self.jobs.setdefault(job_id, {
            "job_id": job_id, "owner_key": owner_key, "fingerprint": fingerprint,
            "space": copy.deepcopy(space), "query": query,
            "definitions": copy.deepcopy(definitions), "cutoff": cutoff,
            "status": "running", "stage": "Ready", "current_issue": None,
            "last_successful_issue": None, "message": "Persistent checkpoint created.",
            "error_detail": None, "result_meta": None, "total_count": 0,
            "completed_count": 0, "items": [],
            "created_at": datetime.now(timezone.utc).isoformat(),
            "updated_at": datetime.now(timezone.utc).isoformat(),
        })
        return {"job_id": job_id, "status": "running"}

    def seed_items(self, owner_key, job_id, items):
        job = self.jobs[job_id]
        if not job["items"]:
            job["items"] = [
                {"position": i, "issue_key": item["key"], "seed_item": copy.deepcopy(item),
                 "item_data": None, "history_data": None, "completed": False,
                 "collected_at": None}
                for i, item in enumerate(items)
            ]
        job["total_count"] = len(job["items"])
        return {"total_count": job["total_count"]}

    def save_item(self, owner_key, job_id, issue_key, item, history):
        job = self.jobs[job_id]
        row = next(r for r in job["items"] if r["issue_key"] == issue_key)
        row.update(item_data=copy.deepcopy(item), history_data=copy.deepcopy(history),
                   completed=True, collected_at=datetime.now(timezone.utc).isoformat())
        job["completed_count"] = sum(1 for r in job["items"] if r["completed"])
        job["last_successful_issue"] = issue_key
        job["updated_at"] = datetime.now(timezone.utc).isoformat()
        return {"completed_count": job["completed_count"], "total_count": job["total_count"]}

    def update_job(self, owner_key, job_id, status, stage, current_issue, message,
                   error_detail=None, result_meta=None):
        job = self.jobs[job_id]
        job.update(status=status, stage=stage, current_issue=current_issue,
                   message=message, error_detail=copy.deepcopy(error_detail),
                   updated_at=datetime.now(timezone.utc).isoformat())
        if result_meta is not None:
            job["result_meta"] = copy.deepcopy(result_meta)
        job["completed_count"] = sum(1 for r in job["items"] if r["completed"])
        job["total_count"] = len(job["items"])
        return {"status": status, "completed_count": job["completed_count"],
                "total_count": job["total_count"]}

    def load_job(self, owner_key, job_id):
        job = self.jobs.get(job_id)
        return copy.deepcopy(job) if job and job["owner_key"] == owner_key else None

    def latest_resumable(self, owner_key):
        matches = [j for j in self.jobs.values()
                   if j["owner_key"] == owner_key and j["status"] in {"running", "error", "complete"}]
        return copy.deepcopy(matches[-1]) if matches else None

    def latest_for_fingerprint(self, owner_key, fingerprint):
        matches = [j for j in self.jobs.values()
                   if j["owner_key"] == owner_key and j["fingerprint"] == fingerprint
                   and j["status"] in {"running", "error", "complete"}]
        return copy.deepcopy(matches[-1]) if matches else None


class FakeJira:
    calls = []
    fail_collection = False

    def __init__(self, settings):
        self.settings = settings
        self.progress = None

    def close(self):
        pass

    def identity(self):
        self.calls.append("identity")
        return {"timeZone": "Asia/Damascus"}

    def spaces(self):
        return [{"id": "100", "key": "TEST", "name": "Test space"}]

    def fields(self):
        return copy.deepcopy(FIELDS)

    def filter_reference(self, project_id):
        return {"visibleFieldNames": [
            {"value": "status", "displayName": "Status", "operators": ["in", "not in", "is"], "auto": "true"},
            {"value": "priority", "displayName": "Priority", "operators": ["in", "not in"], "auto": "true"},
            {"value": "cf[12345]", "cfid": "cf[12345]", "displayName": "Business value", "operators": ["=", ">="], "types": ["java.lang.Double"]},
        ]}

    def suggestions(self, field, value=""):
        return [{"value": "Done", "displayName": "Done"}] if field == "status" else []

    def validate(self, query):
        self.calls.append(query)

    def search_page(self, query, token=None, all_fields=False):
        return {"issues": [example_issue(status="Done")], "isLast": True}

    def all_issues(self, query, progress=None):
        if self.fail_collection:
            raise automation_ui.CollectionError("Simulated permanent collection failure")
        return [example_issue(status="Done")]

    def project_details(self, project_id):
        return {"id": "100", "key": "TEST", "name": "Test space"}

    def complete_issue(self, issue):
        history = no_events()
        history.update(history_through=datetime.now(timezone.utc).isoformat(), snapshot={"current_status": "Done"})
        history["status_events"] = [
            {"from_status": "Idea", "to_status": "In Progress", "changed_at": "2026-09-01T08:00:00+00:00", "change_id": "1"},
            {"from_status": "In Progress", "to_status": "Done", "changed_at": "2026-09-02T10:00:00+00:00", "change_id": "2"},
        ]
        return copy.deepcopy(issue), history


def finish_collection(at, limit=20):
    for _ in range(limit):
        at.run()
        if "prepared_data" in at.session_state:
            return
        job = at.session_state.get("collection_job")
        if job is not None and job.snapshot()["error"]:
            return
    raise AssertionError("Test collection did not finish")


def button(at, label):
    return next(b for b in at.button if b.label == label)


class UserJourneyTests(unittest.TestCase):
    def setUp(self):
        FakeJira.calls = []
        FakeJira.fail_collection = False
        FakeStore.reset()
        self.mock_gateway = patch.object(automation_ui, "JiraGateway", FakeJira)
        self.mock_store = patch.object(automation_ui.CollectionStore, "configured", side_effect=lambda: FakeStore())
        self.mock_gateway.start()
        self.mock_store.start()
        self.addCleanup(self.mock_gateway.stop)
        self.addCleanup(self.mock_store.stop)

    def app(self, configured=True):
        at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=30)
        if configured:
            at.secrets.update(APP_USERNAME="test-admin", APP_PASSWORD="test-password",
                              JIRA_BASE_URL="https://test.atlassian.net",
                              JIRA_EMAIL="test@example.invalid", JIRA_API_TOKEN="test-token")
        at.run()
        self.assertEqual(len(at.exception), 0)
        return at

    def sign_in(self, at):
        at.text_input(key="login_username").input("test-admin")
        at.text_input(key="login_password").input("test-password")
        button(at, "Sign In").click().run()
        self.assertEqual(len(at.exception), 0)

    def select_space(self, at):
        at.selectbox(key="selected_space").set_value("100").run()
        self.assertEqual(len(at.exception), 0)

    def test_sign_in_filters_collection_analysis_and_new_collection(self):
        at = self.app()
        self.assertFalse(FakeJira.calls, "No Jira requests may run before sign-in")
        self.assertEqual(len(at.sidebar), 0)
        self.sign_in(at)
        self.select_space(at)
        self.assertTrue(button(at, "Run Analysis").disabled)
        at.multiselect(key="more_filters").select("cf[12345]").run()
        at.text_input(key="filter_cf[12345]_number").input("5").run()
        self.assertIn('(cf[12345] = 5)', at.session_state["preview_query"])
        button(at, "Done").click().run()
        finish_collection(at)
        self.assertEqual(len(at.exception), 0)
        self.assertEqual(at.session_state["prepared_data"].count, 1)
        self.assertFalse(button(at, "Run Analysis").disabled)
        self.assertTrue(any("ready for analysis" in s.value for s in at.success))
        button(at, "Run Analysis").click().run()
        self.assertEqual(len(at.exception), 0)
        self.assertEqual([t.label for t in at.tabs],
                         ["Executive Dashboard", "Process Analysis", "Individual Achievements", "Task Detail", "Data Quality"])
        self.assertEqual(len(at.session_state["task_metrics"]), 1)
        self.assertEqual(at.session_state["task_metrics"].iloc[0]["status_at_cutoff"], "Done")
        self.assertGreaterEqual(len(at.get("download_button")), 3)

        button(at, "Start New Collection").click().run()
        self.assertNotIn("prepared_data", at.session_state)
        self.assertNotIn("task_metrics", at.session_state)
        self.assertIsNotNone(at.selectbox(key="selected_space"))

        button(at, "Sign Out").click().run()
        self.assertEqual(len(at.exception), 0)
        self.assertIsNotNone(button(at, "Sign In"))
        self.assertNotIn("auth_revision", at.session_state)

    def test_wrong_password_and_missing_configuration_never_call_jira(self):
        at = self.app()
        at.text_input(key="login_username").input("test-admin")
        at.text_input(key="login_password").input("wrong")
        button(at, "Sign In").click().run()
        self.assertEqual(len(at.exception), 0)
        self.assertFalse(FakeJira.calls)
        self.assertTrue(any("Incorrect username" in e.value for e in at.error))
        unconfigured = self.app(configured=False)
        self.assertEqual(len(unconfigured.button), 0)
        self.assertTrue(any("being configured" in i.value for i in unconfigured.info))

    def test_failed_collection_retries_same_persistent_job(self):
        at = self.app()
        self.sign_in(at)
        self.select_space(at)
        FakeJira.fail_collection = True
        button(at, "Done").click().run()
        finish_collection(at)
        self.assertTrue(button(at, "Run Analysis").disabled)
        self.assertNotIn("prepared_data", at.session_state)
        job_id = at.session_state["collection_job"].id
        FakeJira.fail_collection = False
        button(at, "Retry Collection").click().run()
        finish_collection(at)
        self.assertIn("prepared_data", at.session_state)
        self.assertEqual(at.session_state["collection_job"].id, job_id)

    def test_persistent_job_is_offered_after_new_app_session(self):
        at = self.app()
        self.sign_in(at)
        self.select_space(at)
        button(at, "Done").click().run()
        # One full run is enough to persist at least the issue list / first item.
        at.run()
        job_id = at.session_state["collection_job"].id

        reopened = self.app()
        self.sign_in(reopened)
        self.assertIsNotNone(button(reopened, "Resume Previous Collection"))
        button(reopened, "Resume Previous Collection").click().run()
        self.assertEqual(reopened.session_state["collection_job"].id, job_id)
        self.assertGreaterEqual(reopened.session_state["collection_job"].snapshot()["completed"], 0)
        button(reopened, "Resume Collection").click().run()
        finish_collection(reopened)
        self.assertIn("prepared_data", reopened.session_state)


if __name__ == "__main__":
    unittest.main()
