"""Integration smoke tests for the Company Performance application boundary."""

from datetime import date
from io import BytesIO
import json
from pathlib import Path
import sys
import unittest

from openpyxl import Workbook


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from company_performance.application import build_company_analysis
from company_performance.dashboard import DashboardFilters
from company_performance.models import UnifiedStatus


def jira_workbook() -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Jira_Data"
    sheet.append(["Issue key", "Issue id", "Summary", "Project key", "Project name", "Status", "Priority", "Assignee", "Created", "Due date", "Custom field (Start date)"])
    sheet.append(["ENG-1", "1", "Ship feature", "ENG", "Engineering", "Done", "High", "Maya", "2026-09-01T08:00:00Z", "2026-09-10", "2026-09-02"])
    stream = BytesIO()
    workbook.save(stream)
    return stream.getvalue()


class JiraPrepared:
    xlsx = jira_workbook()
    space_name = "Engineering"
    collected_at = "2026-09-30T20:00:00Z"
    history_json = json.dumps({
        "ENG-1": {
            "history_complete": True,
            "history_through": "2026-09-30T20:00:00Z",
            "initial_status": "To Do",
            "status_events": [
                {"changed_at": "2026-09-02T08:00:00Z", "from_status": "To Do", "to_status": "In Progress"},
                {"changed_at": "2026-09-03T08:00:00Z", "from_status": "In Progress", "to_status": "Done"},
            ],
        }
    }).encode()


class ClickUpPrepared:
    space_name = "Growth"
    collected_at = "2026-09-30T20:00:00Z"
    history_json = json.dumps({
        "tasks": [{
            "id": "cu-1", "name": "Review campaign", "status": {"status": "Review"},
            "priority": {"priority": "2"}, "assignees": [{"username": "Zaher"}],
            "date_created": "1788256800000", "due_date": "1790683200000", "list": {"name": "Marketing"},
        }], "time_in_status": {}
    }).encode()


class CompanyApplicationTests(unittest.TestCase):
    def test_prepared_jira_and_clickup_build_one_company_analysis_without_api_calls(self):
        result = build_company_analysis(
            period_start=date(2026, 9, 1), period_end=date(2026, 9, 30),
            jira_prepared=JiraPrepared(), jira_project="Company Delivery",
            clickup_prepared=ClickUpPrepared(), clickup_project="Company Delivery",
        )
        self.assertEqual(result.model.kpis.total_tasks, 2)
        self.assertEqual({item.source_tool for item in result.model.source_coverage}, {"Jira", "ClickUp"})
        self.assertEqual(len(result.model.cards), 5)
        self.assertEqual(len(result.model.executive_charts), 3)
        self.assertIn("ClickUp Chronological History Unavailable", [
            flag.flag for flag in result.model.data_quality
        ])

    def test_filters_apply_after_the_prepared_sources_are_reconstructed(self):
        result = build_company_analysis(
            period_start=date(2026, 9, 1), period_end=date(2026, 9, 30),
            jira_prepared=JiraPrepared(), jira_project="Delivery",
        )
        filtered = build_company_analysis(
            period_start=date(2026, 9, 1), period_end=date(2026, 9, 30),
            jira_prepared=JiraPrepared(), jira_project="Delivery",
            filters=DashboardFilters(statuses=(UnifiedStatus.COMPLETED,)),
        )
        self.assertEqual(result.model.kpis.total_tasks, 1)
        self.assertEqual(filtered.model.kpis.completed_tasks, 1)


if __name__ == "__main__":
    unittest.main()
