"""Focused tests for the isolated Jira/ClickUp Company Performance adapters."""

from datetime import date
from pathlib import Path
import sys
import unittest


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from company_performance.adapters import (
    adapt_clickup_collection,
    adapt_clickup_prepared,
    adapt_jira_collection,
    combine_company_sources,
)
from company_performance.models import ParentClassification


class JiraAdapterTests(unittest.TestCase):
    def test_jira_epic_is_container_and_epic_children_are_not_subtasks(self):
        epic = {
            "key": "TECH-EPIC",
            "fields": {
                "summary": "Website rollout", "issuetype": {"name": "Epic"},
                "project": {"name": "Platform"}, "status": {"name": "Done"},
            },
        }
        task = {
            "key": "TECH-1",
            "fields": {
                "summary": "Build landing page", "issuetype": {"name": "Task"},
                "project": {"name": "Platform"}, "status": {"name": "Done"},
                # Jira Cloud uses parent for an Epic relationship as well as
                # for native Sub-task relationships.
                "parent": {"key": "TECH-EPIC"},
            },
        }
        subtask = {
            "key": "TECH-2",
            "fields": {
                "summary": "Write copy", "issuetype": {"name": "Sub-task"},
                "project": {"name": "Platform"}, "status": {"name": "Done"},
                "parent": {"key": "TECH-1"},
            },
        }

        result = adapt_jira_collection([epic, task, subtask], {}, unified_project="P")
        records = {record.task_id: record for record in result.records}
        self.assertEqual(records["TECH-EPIC"].issue_type, "Epic")
        self.assertIs(records["TECH-EPIC"].parent_classification, ParentClassification.CONTAINER)
        self.assertIsNone(records["TECH-EPIC"].epic_name)
        self.assertEqual(records["TECH-1"].epic_name, "Website rollout")
        self.assertIs(records["TECH-1"].parent_classification, ParentClassification.STANDALONE)
        self.assertIs(records["TECH-2"].parent_classification, ParentClassification.SUBTASK)

    def test_jira_parent_and_subtask_relationships_are_classified(self):
        parent = {
            "key": "TECH-1",
            "fields": {
                "summary": "Parent task", "project": {"name": "Platform"},
                "status": {"name": "In Progress"}, "created": "2026-09-01T08:00:00Z",
            },
        }
        child = {
            "key": "TECH-2",
            "fields": {
                "summary": "Child task", "project": {"name": "Platform"},
                "status": {"name": "Done"}, "created": "2026-09-01T08:00:00Z",
                "parent": {"key": "TECH-1"},
            },
        }
        result = adapt_jira_collection([parent, child], {}, unified_project="P")
        records = {record.task_id: record for record in result.records}
        self.assertIs(records["TECH-1"].parent_classification, ParentClassification.INDEPENDENT)
        self.assertIs(records["TECH-2"].parent_classification, ParentClassification.SUBTASK)

    def test_jira_adapter_keeps_history_dates_and_raw_payloads(self):
        issue = {
            "key": "TECH-7",
            "fields": {
                "summary": "Ship adapter",
                "project": {"name": "Platform"},
                "status": {"name": "Done"},
                "priority": {"name": "Highest"},
                "assignee": {"displayName": "Maya"},
                "created": "2026-09-01T08:00:00Z",
                "duedate": "2026-09-10",
                "customfield_10015": "2026-09-02",
                "parent": {"key": "TECH-1"},
            },
        }
        histories = {
            "TECH-7": {
                "history_complete": True,
                "history_through": "2026-09-12T00:00:00Z",
                "initial_status": "To Do",
                "status_events": [
                    {"changed_at": "2026-09-05T10:00:00Z", "from_status": "In Progress", "to_status": "Done", "author_name": "A"},
                    {"changed_at": "2026-09-03T10:00:00Z", "from_status": "To Do", "to_status": "In Progress", "author_name": "B"},
                ],
            }
        }
        result = adapt_jira_collection(
            [issue], histories, unified_project="Company Platform",
            planned_start_field="customfield_10015", collection_timestamp="2026-09-12T11:00:00Z",
        )
        record = result.records[0]
        self.assertEqual(record.source_tool, "Jira")
        self.assertEqual(record.source_space, "Platform")
        self.assertEqual(record.department, "Tech Development")
        self.assertEqual(record.unified_project, "Company Platform")
        self.assertEqual(record.assignees, ("Maya",))
        self.assertEqual(record.created_date, date(2026, 9, 1))
        self.assertEqual(record.planned_start_date, date(2026, 9, 2))
        self.assertEqual(record.parent_id, "TECH-1")
        self.assertIs(record.parent_classification, ParentClassification.SUBTASK)
        self.assertTrue(record.history_complete)
        self.assertEqual([item.to_status for item in record.workflow_history], ["In Progress", "Done"])
        self.assertEqual(result.coverage.history_mode, "Complete")
        self.assertEqual(result.raw_tasks[record.unique_key]["key"], "TECH-7")
        self.assertEqual(result.raw_workflow[record.unique_key]["initial_status"], "To Do")

    def test_jira_failure_and_empty_source_remain_structured_coverage(self):
        failure = adapt_jira_collection(None, None, unified_project="P", source_space="S", failure_reason="Forbidden")
        empty = adapt_jira_collection([], {}, unified_project="P", source_space="S")
        self.assertFalse(failure.coverage.source_available)
        self.assertIn("Jira Collection Failed", failure.coverage.flags)
        self.assertTrue(empty.coverage.source_available)
        self.assertIn("Empty Source Result", empty.coverage.flags)


class ClickUpAdapterTests(unittest.TestCase):
    def test_clickup_parent_and_subtask_relationships_are_classified(self):
        tasks = [
            {"id": "parent", "name": "Parent", "status": {"status": "In Progress"}, "list": {"name": "Ops"}},
            {"id": "child", "name": "Child", "status": {"status": "Complete"}, "parent": "parent", "list": {"name": "Ops"}},
        ]
        result = adapt_clickup_collection(tasks, unified_project="P", source_space="S")
        records = {record.task_id: record for record in result.records}
        self.assertIs(records["parent"].parent_classification, ParentClassification.INDEPENDENT)
        self.assertIs(records["child"].parent_classification, ParentClassification.SUBTASK)

    def test_clickup_adapter_preserves_snapshot_and_marks_history_unavailable(self):
        task = {
            "id": "cu-1",
            "name": "Prepare launch",
            "status": {"status": "Review"},
            "priority": {"priority": "2"},
            "assignees": [{"username": "Haidar"}, {"email": "maya@example.com"}],
            "date_created": "1788256800000",  # 2026-09-01T10:00:00Z
            "start_date": "1788343200000",
            "due_date": "1789034400000",
            "parent": "cu-parent",
            "list": {"name": "Marketing"},
        }
        time_status = {"cu-1": {"current_status": {"status": "Review", "total_time": {"by_minute": 45}}}}
        result = adapt_clickup_collection(
            [task], time_status, unified_project="Launch", source_space="Growth Space",
            collection_timestamp="2026-09-12T12:00:00Z",
        )
        record = result.records[0]
        self.assertEqual(record.source_tool, "ClickUp")
        self.assertEqual(record.source_space, "Growth Space")
        self.assertEqual(record.department, "Marketing")
        self.assertEqual(record.assignees, ("Haidar", "maya@example.com"))
        self.assertEqual(record.parent_id, "cu-parent")
        self.assertFalse(record.history_complete)
        self.assertIn("ClickUp Chronological History Unavailable", record.data_quality_flags)
        self.assertNotIn("ClickUp Total Time in Status Unavailable", record.data_quality_flags)
        self.assertIn(record.unique_key, result.raw_time_in_status)
        self.assertEqual(result.coverage.history_mode, "Unavailable")

    def test_clickup_partial_time_status_and_prepared_payload(self):
        tasks = [
            {"id": "1", "name": "One", "status": {"status": "To Do"}, "list": {"name": "Ops"}},
            {"id": "2", "name": "Two", "status": {"status": "Complete"}, "list": {"name": "Ops"}},
        ]
        direct = adapt_clickup_collection(tasks, {"1": {"current_status": {}}}, unified_project="P", source_space="S")
        self.assertIn("Partial Total Time in Status Coverage", direct.coverage.flags)
        self.assertIn("ClickUp Total Time in Status Unavailable", direct.records[1].data_quality_flags)

        class Prepared:
            history_json = b'{"tasks":[{"id":"3","name":"Three","status":{"status":"Planning"},"list":{"name":"Ops"}}],"time_in_status":{}}'
            collected_at = "2026-09-12T12:00:00Z"
            space_name = "Prepared Space"

        result = adapt_clickup_prepared(Prepared(), unified_project="P")
        self.assertEqual(result.records[0].source_space, "Prepared Space")
        self.assertEqual(result.records[0].department, "Ops")

    def test_company_collection_keeps_sources_separate(self):
        jira = adapt_jira_collection([], {}, unified_project="P", source_space="J")
        clickup = adapt_clickup_collection([], unified_project="P", source_space="C")
        collection = combine_company_sources(jira, clickup)
        self.assertEqual(collection.records, ())
        self.assertEqual([coverage.source_tool for coverage in collection.coverages], ["Jira", "ClickUp"])


if __name__ == "__main__":
    unittest.main()
