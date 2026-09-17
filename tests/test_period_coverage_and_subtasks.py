from __future__ import annotations

from datetime import date, datetime, timezone
import json
from types import SimpleNamespace
import unittest

import pandas as pd

from company_performance.application import build_company_analysis
from company_performance.kpis import calculate_core_kpis
from company_performance.models import StatusTransition, TaskRecord, UnifiedStatus
from company_performance.workflow import reconstruct_task
from department_analysis import build_jira_department_result


UTC = timezone.utc


class PeriodCoverageTests(unittest.TestCase):
    def test_jira_status_is_unknown_beyond_verified_history_coverage(self):
        task = TaskRecord(
            source_tool="Jira",
            task_id="ENG-1",
            raw_status="Done",
            initial_status="To Do",
            created_date=date(2026, 9, 1),
            history_complete=True,
            history_through=datetime(2026, 9, 30, 20, 59, tzinfo=UTC),
            workflow_history=(StatusTransition(
                datetime(2026, 9, 20, 9, 0, tzinfo=UTC), "To Do", "Done"
            ),),
        )
        snapshot = reconstruct_task(
            task, date(2026, 9, 1), date(2026, 10, 1),
            collection_date=date(2026, 9, 30),
        )
        self.assertEqual(snapshot.status_at_period_end, UnifiedStatus.UNKNOWN)
        self.assertIn(
            "Analysis Period Exceeds History Coverage",
            snapshot.data_quality_flags,
        )

    def test_unknown_status_makes_completion_rate_unavailable(self):
        task = TaskRecord(
            source_tool="ClickUp",
            task_id="CU-1",
            raw_status="complete",
            created_date=date(2026, 9, 1),
            history_complete=False,
        )
        snapshot = reconstruct_task(
            task, date(2026, 9, 1), date(2026, 9, 30),
            collection_date=date(2026, 10, 1),
        )
        self.assertIsNone(calculate_core_kpis([snapshot]).completion_rate)

    def test_analysis_period_cannot_exceed_collected_clickup_coverage(self):
        prepared = SimpleNamespace(
            history_json=json.dumps({
                "tasks": [{
                    "id": "CU-1",
                    "name": "Task",
                    "status": {"status": "in progress"},
                    "date_created": "1788253200000",
                    "list": {"name": "Operations"},
                }],
                "time_in_status": {},
            }).encode(),
            collected_at="2026-09-30T10:00:00Z",
            space_name="Operations",
        )
        with self.assertRaisesRegex(ValueError, "cannot be later"):
            build_company_analysis(
                period_start=date(2026, 9, 1),
                period_end=date(2026, 10, 1),
                clickup_prepared=prepared,
                unified_project="Operations",
            )


class DepartmentSubtaskTests(unittest.TestCase):
    def test_jira_subtasks_do_not_change_department_completion_rate(self):
        frame = pd.DataFrame([
            {
                "issue_key": "ENG-1", "task_name": "Parent", "issue_type": "Task",
                "assignee_name": "Maya", "status_at_cutoff": "Done",
                "is_completed": True, "is_rejected": False, "is_open": False,
                "is_wip": False, "history_complete": True,
                "completed_at": "2026-09-10T09:00:00Z",
            },
            {
                "issue_key": "ENG-2", "task_name": "Child", "issue_type": "Sub-task",
                "assignee_name": "Maya", "status_at_cutoff": "To Do",
                "is_completed": False, "is_rejected": False, "is_open": True,
                "is_wip": False, "history_complete": True,
            },
            {
                "issue_key": "ENG-3", "task_name": "Unknown parent", "issue_type": "Task",
                "assignee_name": "Maya", "status_at_cutoff": "Unavailable",
                "is_completed": False, "is_rejected": False, "is_open": False,
                "is_wip": False, "history_complete": False, "status_known": False,
            },
        ])
        result = build_jira_department_result(
            frame, {}, "Tech", "2026-09-30T20:59:59Z"
        )
        kpis = result["kpis"].set_index("KPI")
        self.assertEqual(kpis.loc["Total Tasks", "Value"], 3)
        self.assertEqual(kpis.loc["Task Completion Rate (%)", "Value"], 100.0)


if __name__ == "__main__":
    unittest.main()
