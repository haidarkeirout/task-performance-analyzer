"""Exercise the real Streamlit flow with a deterministic Jira transport."""
import copy
from datetime import datetime, timezone
import unittest
from unittest.mock import patch

from streamlit.testing.v1 import AppTest

from test_automation import ROOT, FIELDS, example_issue, no_events
import automation_ui


class FakeJira:
    calls = []
    fail_collection = False

    def __init__(self, settings):
        self.settings = settings

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
            raise automation_ui.CollectionError("Jira is busy. Please try again.")
        return [example_issue(status="Done")]

    def complete_issue(self, issue):
        history = no_events()
        history.update(history_through=datetime.now(timezone.utc).isoformat(), snapshot={"current_status": "Done"})
        history["status_events"] = [
            {"from_status": "Idea", "to_status": "In Progress", "changed_at": "2026-09-01T08:00:00+00:00", "change_id": "1"},
            {"from_status": "In Progress", "to_status": "Done", "changed_at": "2026-09-02T10:00:00+00:00", "change_id": "2"},
        ]
        return copy.deepcopy(issue), history


def finish_collection(at):
    job = at.session_state["collection_job"]
    for _ in range(100):
        if not job.running:
            break
        job.step()
    assert not job.running, "Test collection did not finish"
    at.run()


def button(at, label):
    return next(b for b in at.button if b.label == label)


class UserJourneyTests(unittest.TestCase):
    def setUp(self):
        FakeJira.calls = []
        FakeJira.fail_collection = False
        self.mock_gateway = patch.object(automation_ui, "JiraGateway", FakeJira)
        self.mock_gateway.start()
        self.addCleanup(self.mock_gateway.stop)

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

    def test_sign_in_filters_collection_analysis_and_invalidation(self):
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
        self.assertEqual([t.label for t in at.tabs], ["Executive Dashboard", "Process Analysis", "Individual Achievements", "Task Detail", "Data Quality"])
        self.assertEqual(len(at.session_state["task_metrics"]), 1)
        self.assertEqual(at.session_state["task_metrics"].iloc[0]["status_at_cutoff"], "Done")
        self.assertGreaterEqual(len(at.get("download_button")), 3)
        at.multiselect(key="filter_status_values").select("Done").run()
        self.assertEqual(len(at.exception), 0)
        self.assertTrue(button(at, "Run Analysis").disabled)
        self.assertNotIn("prepared_data", at.session_state)
        self.assertNotIn("task_metrics", at.session_state)
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

    def test_failed_collection_and_changed_credentials_clear_data(self):
        at = self.app()
        self.sign_in(at)
        self.select_space(at)
        FakeJira.fail_collection = True
        button(at, "Done").click().run()
        finish_collection(at)
        self.assertEqual(len(at.exception), 0)
        self.assertTrue(button(at, "Run Analysis").disabled)
        self.assertNotIn("prepared_data", at.session_state)
        FakeJira.fail_collection = False
        button(at, "Done").click().run()
        finish_collection(at)
        self.assertIn("prepared_data", at.session_state)
        at.secrets["APP_PASSWORD"] = "new-password"
        at.run()
        self.assertEqual(len(at.exception), 0)
        self.assertNotIn("prepared_data", at.session_state)
        self.assertIsNotNone(button(at, "Sign In"))


if __name__ == "__main__":
    unittest.main()
