import unittest
from types import SimpleNamespace
from unittest.mock import patch

from employee_ui import _load_employee_sources, _snapshot_visible


class EmployeeCrossSourceTests(unittest.TestCase):
    def test_loading_a_dual_source_record_calls_both_connectors(self):
        record = SimpleNamespace(
            name="Ada", jira_account_id="jira-1", clickup_user_id="clickup-1"
        )
        jira = {"source": "Jira", "spaces": {"jira-space": "Jira"}, "issues": []}
        clickup = {"source": "ClickUp", "spaces": {"clickup-space": "ClickUp"}, "tasks": []}
        with patch("employee_ui._load_jira", return_value=jira) as load_jira, patch(
            "employee_ui._load_clickup", return_value=clickup
        ) as load_clickup:
            snapshot = _load_employee_sources(record, object())

        load_jira.assert_called_once()
        load_clickup.assert_called_once()
        self.assertEqual(snapshot["source"], "Combined")
        self.assertEqual(set(snapshot["spaces"]), {"Jira:jira-space", "ClickUp:clickup-space"})

    def test_snapshot_preview_keeps_source_and_space_identity(self):
        snapshot = {
            "source": "Combined",
            "sources": {
                "Jira": {
                    "issues": [{
                        "key": "JRA-1",
                        "fields": {
                            "summary": "Jira task",
                            "project": {"id": "jira-space", "name": "Najm Jira"},
                            "status": {"name": "Done"},
                            "priority": {"name": "High"},
                            "assignee": {"displayName": "Ada"},
                        },
                    }],
                },
                "ClickUp": {
                    "tasks": [{
                        "id": "cu-1",
                        "name": "ClickUp task",
                        "employee_space_id": "clickup-space",
                        "employee_space_name": "Najm ClickUp",
                        "status": {"status": "in progress"},
                        "priority": {"priority": "urgent"},
                        "assignees": [{"username": "Ada"}],
                    }],
                },
            },
        }

        visible, rows = _snapshot_visible(snapshot, [])

        self.assertEqual({row["Source"] for row in rows}, {"Jira", "ClickUp"})
        self.assertEqual({row["Task"] for row in rows}, {"JRA-1", "cu-1"})
        self.assertEqual(len(visible), 2)

    def test_space_filter_is_source_qualified(self):
        snapshot = {
            "source": "Combined",
            "sources": {
                "Jira": {"issues": [{"key": "JRA-1", "fields": {"project": {"id": "same", "name": "Jira"}}}]},
                "ClickUp": {"tasks": [{"id": "CU-1", "name": "CU", "employee_space_id": "same", "employee_space_name": "ClickUp"}]},
            },
        }
        visible, rows = _snapshot_visible(snapshot, ["Jira:same"])

        self.assertEqual([row["Task"] for row in rows], ["JRA-1"])
        self.assertEqual(len(visible), 1)


if __name__ == "__main__":
    unittest.main()
