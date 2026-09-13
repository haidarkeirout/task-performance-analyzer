"""Focused tests for the isolated Company Performance dashboard contract."""

from datetime import date, datetime, timezone
from pathlib import Path
import sys
import unittest


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from company_performance.dashboard import DashboardFilters, build_company_dashboard
from company_performance.models import StatusTransition, TaskRecord, UnifiedStatus
from company_performance.workflow import reconstruct_task


UTC = timezone.utc
START = date(2026, 9, 1)
END = date(2026, 9, 30)


def event(before, after, value):
    return StatusTransition(
        changed_at=datetime.fromisoformat(value).replace(tzinfo=UTC),
        from_status=before,
        to_status=after,
    )


def snapshot(task_id, events=(), **extra):
    values = {
        "source_tool": "Jira",
        "task_id": task_id,
        "task_name": task_id,
        "source_space": "Engineering",
        "unified_project": "Platform",
        "raw_status": "In Review",
        "initial_status": "To Do",
        "created_date": START,
        "history_complete": True,
        "workflow_history": events,
    }
    values.update(extra)
    return reconstruct_task(TaskRecord(**values), START, END)


class CompanyDashboardTests(unittest.TestCase):
    def test_dashboard_has_exactly_five_cards_and_three_charts(self):
        done = snapshot(
            "done",
            (event("To Do", "In Progress", "2026-09-02T09:00:00"),
             event("In Progress", "Done", "2026-09-05T09:00:00")),
            raw_status="Done",
            due_date=date(2026, 9, 10),
        )
        review = snapshot(
            "review",
            (event("To Do", "In Progress", "2026-09-03T09:00:00"),
             event("In Progress", "In Review", "2026-09-04T09:00:00")),
            due_date=date(2026, 9, 10),
        )
        model = build_company_dashboard([done, review])
        self.assertEqual(len(model.cards), 5)
        self.assertEqual(len(model.executive_charts), 3)
        self.assertEqual([card.key for card in model.cards], [
            "total-tasks", "completion-rate", "current-wip", "overdue-open", "on-time-rate",
        ])
        self.assertEqual(model.kpis.current_wip, 1)

    def test_filters_drive_task_details_and_kpis_without_reconstructing_history(self):
        platform = snapshot("platform", unified_project="Platform")
        mobile = snapshot("mobile", unified_project="Mobile", raw_status="Done",
                          workflow_history=(event("To Do", "Done", "2026-09-04T09:00:00"),))
        model = build_company_dashboard(
            [platform, mobile], filters=DashboardFilters(unified_projects=("Mobile",))
        )
        self.assertEqual(model.kpis.total_tasks, 1)
        self.assertEqual([item.task_id for item in model.task_details], ["mobile"])
        self.assertIn("Platform", model.filter_options.unified_projects)
        self.assertIn("Mobile", model.filter_options.unified_projects)

    def test_delivery_outcome_excludes_unknown_and_exposes_count_as_note_and_quality(self):
        unmapped = snapshot("unmapped", raw_status="Custom Status", initial_status="Custom Status")
        model = build_company_dashboard([unmapped])
        delivery = next(chart for chart in model.executive_charts if chart.key == "delivery-outcome")
        self.assertEqual(delivery.points, ())
        self.assertEqual(delivery.note, "Tasks not classified due to missing or unmapped data: 1")
        self.assertIn("Unmapped Status", [item.flag for item in model.data_quality])

    def test_task_details_keep_in_period_transition_and_exception_evidence(self):
        item = snapshot(
            "review-return",
            (event("To Do", "In Progress", "2026-09-02T09:00:00"),
             event("In Progress", "In Review", "2026-09-03T09:00:00"),
             event("In Review", "In Progress", "2026-09-04T09:00:00")),
        )
        detail = build_company_dashboard([item]).task_details[0]
        self.assertIn("Rework", detail.exception_events)
        self.assertIn("2026-09-04: In Review → In Progress", detail.workflow_events)


if __name__ == "__main__":
    unittest.main()
