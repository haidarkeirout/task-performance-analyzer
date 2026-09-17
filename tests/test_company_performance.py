"""Focused acceptance tests for the isolated Company Performance domain."""

from datetime import date, datetime, timezone
from pathlib import Path
import sys
import unittest


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from company_performance.kpis import average_days_in_status, calculate_core_kpis
from company_performance.models import (
    ParentClassification,
    StatusTransition,
    TaskRecord,
    UnifiedStatus,
)
from company_performance.normalization import (
    assignee_group,
    deduplicate_tasks,
    normalize_priority,
    normalize_status,
)
from company_performance.workflow import reconstruct_task


UTC = timezone.utc
START = date(2026, 9, 1)
END = date(2026, 9, 30)


def event(from_status, to_status, when):
    return StatusTransition(
        datetime.fromisoformat(when).replace(tzinfo=UTC), from_status, to_status
    )


def task(source="Jira", task_id="T-1", **values):
    base = {
        "source_tool": source,
        "task_id": task_id,
        "raw_status": "Done" if source == "Jira" else "Complete",
        "initial_status": "To Do",
        "created_date": date(2026, 9, 1),
        "history_complete": True,
        "history_through": datetime(2026, 10, 1, tzinfo=UTC),
        "collection_timestamp": datetime(2026, 10, 1, tzinfo=UTC),
    }
    base.update(values)
    return TaskRecord(**base)


class MappingAndAttributionTests(unittest.TestCase):
    def test_only_approved_statuses_map_and_priorities_follow_source_rules(self):
        self.assertEqual(normalize_status("ClickUp", "Planning"), UnifiedStatus.NOT_STARTED)
        self.assertEqual(normalize_status("ClickUp", "At Risk"), UnifiedStatus.AT_RISK)
        self.assertEqual(normalize_status("Jira", "In Triage"), UnifiedStatus.NOT_STARTED)
        self.assertEqual(normalize_status("Jira", "Rejected"), UnifiedStatus.REJECTED)
        self.assertEqual(normalize_status("Jira", "QA Testing"), UnifiedStatus.UNKNOWN)
        self.assertEqual(normalize_priority("Jira", "Highest"), "Critical")
        self.assertEqual(normalize_priority("ClickUp", "Urgent"), "Critical")
        self.assertEqual(normalize_priority("ClickUp", "3"), "Medium")
        self.assertEqual(normalize_priority("Jira", None), "Unknown")

    def test_deduplication_keeps_latest_backfills_and_flags_conflicts(self):
        older = task(
            task_id="same",
            task_name="Original name",
            due_date=date(2026, 9, 10),
            collection_timestamp=datetime(2026, 9, 10, tzinfo=UTC),
        )
        latest = task(
            task_id="same",
            task_name=None,
            due_date=date(2026, 9, 12),
            collection_timestamp=datetime(2026, 9, 11, tzinfo=UTC),
        )
        result = deduplicate_tasks([older, latest])
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0].task_name, "Original name")
        self.assertEqual(result[0].due_date, date(2026, 9, 12))
        self.assertIn("Duplicate Task Record", result[0].data_quality_flags)
        self.assertIn("Duplicate Conflicting Records", result[0].data_quality_flags)

    def test_assignee_and_parent_counting_rules_do_not_duplicate_tasks(self):
        unassigned = task(assignees=())
        multiple = task(task_id="two", assignees=("Maya", "Haidar"))
        container = task(task_id="parent", parent_classification=ParentClassification.CONTAINER)
        self.assertEqual(assignee_group(unassigned), "Unassigned")
        self.assertEqual(assignee_group(multiple), "Multiple Assignees")
        self.assertFalse(container.counted_in_kpis)
        self.assertTrue(
            task(parent_classification=ParentClassification.INDEPENDENT).counted_in_kpis
        )
        self.assertFalse(
            task(parent_classification=ParentClassification.SUBTASK).counted_in_kpis
        )

    def test_subtasks_remain_in_total_but_are_excluded_from_performance_kpis(self):
        parent = reconstruct_task(
            task(task_id="parent", raw_status="Done", workflow_history=(
                event("To Do", "Done", "2026-09-05T09:00:00"),
            )), START, END
        )
        child = reconstruct_task(
            task(task_id="child", parent_id="parent", parent_classification=ParentClassification.SUBTASK),
            START, END
        )
        metrics = calculate_core_kpis([parent, child])
        self.assertEqual(metrics.total_tasks, 2)
        self.assertEqual(metrics.completed_tasks, 1)
        self.assertEqual(metrics.completion_rate, 100.0)

    def test_unknown_status_is_excluded_from_completion_denominator(self):
        completed = reconstruct_task(
            task(
                task_id="known",
                raw_status="Done",
                workflow_history=(event("To Do", "Done", "2026-09-05T09:00:00"),),
            ),
            START,
            END,
        )
        unknown = reconstruct_task(
            task(task_id="unknown", history_complete=False, raw_status="Done"),
            START,
            END,
        )
        metrics = calculate_core_kpis([completed, unknown])
        self.assertEqual(metrics.total_tasks, 2)
        self.assertEqual(metrics.known_status_tasks, 1)
        self.assertEqual(metrics.unknown_status_tasks, 1)
        self.assertEqual(metrics.completion_rate, 100.0)

class WorkflowTests(unittest.TestCase):
    def test_reopened_task_is_visible_in_period_but_not_completed_at_period_end(self):
        record = task(
            created_date=date(2026, 8, 20),
            initial_status="To Do",
            workflow_history=(
                event("To Do", "In Progress", "2026-08-21T09:00:00"),
                event("In Progress", "Done", "2026-08-25T09:00:00"),
                event("Done", "In Progress", "2026-09-03T09:00:00"),
            ),
        )
        result = reconstruct_task(record, START, END)
        self.assertTrue(result.in_scope)
        self.assertEqual(result.status_at_period_end, UnifiedStatus.IN_EXECUTION)
        self.assertEqual(result.reopen_events, (date(2026, 9, 3),))
        self.assertIsNone(result.final_completion_date)

    def test_actual_start_is_first_in_progress_and_review_does_not_qualify(self):
        record = task(
            workflow_history=(
                event("To Do", "In Review", "2026-09-02T09:00:00"),
                event("In Review", "In Progress", "2026-09-04T09:00:00"),
                event("In Progress", "Done", "2026-09-06T09:00:00"),
            )
        )
        result = reconstruct_task(record, START, END)
        self.assertEqual(result.actual_start_date, date(2026, 9, 4))
        self.assertEqual(result.final_completion_date, date(2026, 9, 6))

    def test_direct_completion_does_not_invent_actual_start(self):
        record = task(
            workflow_history=(event("To Do", "Done", "2026-09-02T09:00:00"),)
        )
        result = reconstruct_task(record, START, END)
        self.assertEqual(result.status_at_period_end, UnifiedStatus.COMPLETED)
        self.assertIsNone(result.actual_start_date)
        self.assertIn("Missing Actual Start", result.data_quality_flags)

    def test_jira_exceptions_and_clickup_reactivation_are_separate(self):
        jira = task(
            workflow_history=(
                event("To Do", "In Review", "2026-09-02T09:00:00"),
                event("In Review", "In Progress", "2026-09-03T09:00:00"),
                event("In Progress", "In Review", "2026-09-04T09:00:00"),
                event("In Review", "To Do", "2026-09-05T09:00:00"),
                event("To Do", "In Review", "2026-09-06T09:00:00"),
                event("In Review", "In Triage", "2026-09-07T09:00:00"),
            )
        )
        clickup = task(
            source="ClickUp",
            task_id="C-1",
            raw_status="In Progress",
            initial_status="Cancelled",
            workflow_history=(event("Cancelled", "In Progress", "2026-09-03T09:00:00"),),
        )
        jira_result = reconstruct_task(jira, START, END)
        clickup_result = reconstruct_task(clickup, START, END)
        self.assertEqual(jira_result.exception_events, ("Rework", "Replanning", "Re-evaluation"))
        self.assertEqual(clickup_result.reactivation_events, (date(2026, 9, 3),))
        self.assertEqual(clickup_result.reopen_events, ())

    def test_missing_past_history_does_not_use_current_status(self):
        record = task(history_complete=False, raw_status="Done", workflow_history=())
        result = reconstruct_task(record, START, END, collection_date=date(2026, 10, 1))
        self.assertFalse(result.history_available)
        self.assertEqual(result.status_at_period_end, UnifiedStatus.UNKNOWN)
        self.assertIn("Missing Workflow History", result.data_quality_flags)

    def test_same_day_snapshot_does_not_replace_missing_historical_history(self):
        record = task(
            history_complete=False,
            raw_status="Done",
            workflow_history=(),
            collection_timestamp=datetime(2026, 9, 30, tzinfo=UTC),
        )
        result = reconstruct_task(record, START, END, collection_date=END)
        self.assertFalse(result.history_available)
        self.assertEqual(result.status_at_period_end, UnifiedStatus.UNKNOWN)
        self.assertNotIn("Snapshot Status Fallback", result.data_quality_flags)
        self.assertIn("Missing Workflow History", result.data_quality_flags)

    def test_status_intervals_are_clipped_to_the_selected_period(self):
        record = task(
            created_date=date(2026, 8, 28),
            workflow_history=(
                event("To Do", "In Progress", "2026-08-30T09:00:00"),
                event("In Progress", "In Review", "2026-09-10T09:00:00"),
            ),
        )
        result = reconstruct_task(record, START, END)
        execution = [x for x in result.status_intervals if x.status is UnifiedStatus.IN_EXECUTION]
        self.assertEqual(execution[0].start_date, START)
        self.assertEqual(execution[0].end_date, date(2026, 9, 10))
        self.assertEqual(execution[0].days, 9)
        self.assertEqual(average_days_in_status([result], UnifiedStatus.IN_EXECUTION), 9.0)


class KPItests(unittest.TestCase):
    def test_core_counts_rates_and_zero_denominators(self):
        complete = reconstruct_task(
            task(
                task_id="done",
                priority="High",
                due_date=date(2026, 9, 5),
                workflow_history=(
                    event("To Do", "In Progress", "2026-09-02T09:00:00"),
                    event("In Progress", "Done", "2026-09-05T09:00:00"),
                ),
            ),
            START,
            END,
        )
        overdue = reconstruct_task(
            task(task_id="open", priority="Highest", due_date=date(2026, 9, 10)), START, END
        )
        cancelled = reconstruct_task(
            task(
                source="ClickUp",
                task_id="cancelled",
                raw_status="Cancelled",
                initial_status="To Do",
                workflow_history=(event("To Do", "Cancelled", "2026-09-02T09:00:00"),),
            ),
            START,
            END,
        )
        container = reconstruct_task(
            task(
                task_id="container",
                parent_classification=ParentClassification.CONTAINER,
            ),
            START,
            END,
        )
        metrics = calculate_core_kpis([complete, overdue, cancelled, container])
        self.assertEqual(metrics.total_tasks, 3)
        self.assertEqual(metrics.completed_tasks, 1)
        self.assertEqual(metrics.open_tasks, 1)
        self.assertEqual(metrics.overdue_open_tasks, 1)
        self.assertEqual(metrics.high_priority_overdue_tasks, 1)
        self.assertEqual(metrics.completion_rate, 33.3)
        self.assertEqual(metrics.on_time_completion_rate, 100.0)
        self.assertEqual(metrics.average_time_to_start_days, 1.0)
        self.assertEqual(metrics.average_execution_duration_days, 3.0)
        only_cancelled = calculate_core_kpis([cancelled])
        self.assertEqual(only_cancelled.completion_rate, 0.0)


if __name__ == "__main__":
    unittest.main()
