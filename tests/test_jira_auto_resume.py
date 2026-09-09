"""Regression tests for automatic continuation of active Jira collections."""
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

import automation_ui


class AutoResumeTests(unittest.TestCase):
    def test_running_persistent_job_rehydrates_and_starts_without_user_click(self):
        payload = {
            "job_id": "job-123",
            "status": "running",
            "space": {"id": "100", "key": "TEST", "name": "Test"},
            "query": "project = TEST",
            "fingerprint": "fp",
            "definitions": [],
            "items": [],
        }
        store = Mock()
        store.latest_resumable.return_value = payload
        job = Mock()
        state = {}

        with patch.object(automation_ui, "st", SimpleNamespace(session_state=state)), \
             patch.object(automation_ui.CollectionStore, "configured", return_value=store), \
             patch.object(automation_ui, "persistence_owner_key", return_value="owner-key"), \
             patch.object(automation_ui.CollectionJob, "from_persisted", return_value=job) as restore:
            automation_ui._auto_restore_running_jira_job(object())

        restore.assert_called_once()
        job.start.assert_called_once_with()
        self.assertIs(state["collection_job"], job)
        self.assertTrue(state["persistent_candidate_checked"])
        self.assertNotIn("persistent_candidate", state)
        store.close.assert_not_called()

    def test_error_or_complete_job_is_not_auto_started(self):
        for status in ("error", "complete"):
            with self.subTest(status=status):
                store = Mock()
                store.latest_resumable.return_value = {"job_id": "job-123", "status": status}
                state = {}
                with patch.object(automation_ui, "st", SimpleNamespace(session_state=state)), \
                     patch.object(automation_ui.CollectionStore, "configured", return_value=store), \
                     patch.object(automation_ui, "persistence_owner_key", return_value="owner-key"), \
                     patch.object(automation_ui.CollectionJob, "from_persisted") as restore:
                    automation_ui._auto_restore_running_jira_job(object())
                restore.assert_not_called()
                self.assertNotIn("collection_job", state)
                store.close.assert_called_once_with()


if __name__ == "__main__":
    unittest.main()
