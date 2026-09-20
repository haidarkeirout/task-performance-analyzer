"""Focused acceptance tests for Company Performance KPI evidence and advice."""

from datetime import date, datetime, timezone
from pathlib import Path
import sys
import unittest


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from company_performance.kpis import (
    calculate_core_kpis,
    calculate_status_metrics,
    generate_recommendations,
    identify_bottleneck_candidates,
)
from company_performance.models import StatusTransition, TaskRecord, UnifiedStatus
from company_performance.workflow import reconstruct_task


UTC = timezone.utc
START = date(2026, 9, 1)
END = date(2026, 9, 30)


def transition(before, after, value):
    return StatusTransition(
        changed_at=datetime.fromisoformat(value).replace(tzinfo=UTC),
        from_status=before,
        to_status=after,
    )


def snapshot(task_id, events=(), **values):
    fields = {
        "source_tool": "Jira",
        "task_id": task_id,
        "raw_status": "In Review",
        "initial_status": "To Do",
        "created_date": START,
        "history_complete": True,
        "history_through": datetime(2026, 10, 1, tzinfo=UTC),
        "collection_timestamp": datetime(2026, 10, 1, tzinfo=UTC),
        "workflow_history": events,
    }
    fields.update(values)
    return reconstruct_task(TaskRecord(**fields), START, END)


class CompanyKpiEvidenceTests(unittest.TestCase):
    def test_metrics_use_clipped_history_and_terminal_stages_are_excluded(self):
        review = snapshot(
            "review",
            (
                transition("To Do", "In Progress", "2026-09-02T09:00:00"),
                transition("In Progress", "In Review", "2026-09-04T09:00:00"),
            ),
            due_date=date(2026, 9, 10),
        )
        metrics = {item.status: item for item in calculate_status_metrics([review])}
        self.assertNotIn(UnifiedStatus.COMPLETED, metrics)
        self.assertEqual(metrics[UnifiedStatus.IN_REVIEW].tasks_passed_through, 1)
        self.assertEqual(metrics[UnifiedStatus.IN_REVIEW].average_days, 26.0)
        self.assertEqual(metrics[UnifiedStatus.IN_REVIEW].median_days, 26.0)
        self.assertEqual(metrics[UnifiedStatus.IN_REVIEW].open_tasks_now, 1)
        self.assertEqual(metrics[UnifiedStatus.IN_REVIEW].overdue_open_tasks, 1)

    def test_candidate_requires_multiple_evidence_signals_and_is_not_confirmed(self):
        one = snapshot(
            "one",
            (
                transition("To Do", "In Progress", "2026-09-02T09:00:00"),
                transition("In Progress", "In Review", "2026-09-04T09:00:00"),
                transition("In Review", "In Progress", "2026-09-08T09:00:00"),
                transition("In Progress", "In Review", "2026-09-09T09:00:00"),
            ),
            due_date=date(2026, 9, 12),
        )
        two = snapshot(
            "two",
            (
                transition("To Do", "In Progress", "2026-09-02T09:00:00"),
                transition("In Progress", "In Review", "2026-09-05T09:00:00"),
            ),
            due_date=date(2026, 9, 12),
        )
        candidates = identify_bottleneck_candidates([one, two])
        review = next(item for item in candidates if item.status is UnifiedStatus.IN_REVIEW)
        self.assertEqual(review.strength, "Strong Bottleneck Candidate")
        self.assertIn("Repeated workflow returns from review", review.evidence)
        self.assertIn("Open overdue tasks in status", review.evidence)
        self.assertNotIn("Confirmed", review.strength)

    def test_recommendations_use_thresholds_only_as_recommendations(self):
        late_done = snapshot(
            "late",
            (
                transition("To Do", "In Progress", "2026-09-02T09:00:00"),
                transition("In Progress", "Done", "2026-09-20T09:00:00"),
            ),
            raw_status="Done",
            due_date=date(2026, 9, 10),
        )
        overdue = snapshot(
            "overdue",
            (transition("To Do", "In Progress", "2026-09-02T09:00:00"),),
            due_date=date(2026, 9, 10),
            priority="Highest",
        )
        recommendations = {item.code: item for item in generate_recommendations([late_done, overdue])}
        self.assertIn("completion-rate", recommendations)
        self.assertIn("on-time-rate", recommendations)
        self.assertIn("priority-overdue", recommendations)
        self.assertIn("recommendation threshold", recommendations["completion-rate"].evidence)

    def test_rates_and_averages_are_presented_to_one_decimal(self):
        first = snapshot(
            "first",
            (
                transition("To Do", "In Progress", "2026-09-02T09:00:00"),
                transition("In Progress", "Done", "2026-09-04T09:00:00"),
            ),
            raw_status="Done",
        )
        second = snapshot(
            "second",
            (
                transition("To Do", "In Progress", "2026-09-03T09:00:00"),
                transition("In Progress", "Done", "2026-09-06T09:00:00"),
            ),
            raw_status="Done",
        )
        third = snapshot("third")
        core = calculate_core_kpis([first, second, third])
        self.assertEqual(core.completion_rate, 66.7)
        self.assertEqual(core.average_execution_duration_days, 2.5)


if __name__ == "__main__":
    unittest.main()
