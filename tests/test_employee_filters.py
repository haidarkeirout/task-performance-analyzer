import unittest
from unittest.mock import Mock

import pandas as pd
from datetime import date

from clickup_analysis import analyze_clickup, recalculate_clickup_analysis
from clickup_export import collect_data
from company_performance.dashboard import DashboardFilters, available_filters, filter_snapshots
from company_performance.models import TaskPeriodSnapshot, TaskRecord, UnifiedStatus
from employee_filters import apply_employee_filters, employee_filter_options, with_employee_filter_dimensions


class EmployeeFilterTests(unittest.TestCase):
    def setUp(self):
        self.jira = pd.DataFrame([
            {
                "issue_key": "A-1", "project_name": "Alpha Project", "project_key": "ALPHA",
                "company_name": "Acme", "status_at_cutoff": "Done", "issue_type": "Task",
                "priority": "High", "is_completed": True, "is_rejected": False, "is_open": False,
                "schedule_variance_days": 0, "overdue_days": 0,
            },
            {
                "issue_key": "B-1", "project_name": "Beta Project", "project_key": "BETA",
                "company_name": "Beta", "status_at_cutoff": "In Progress", "issue_type": "Bug",
                "priority": "Low", "is_completed": False, "is_rejected": False, "is_open": True,
                "schedule_variance_days": None, "overdue_days": 3,
            },
        ])

    def test_projects_are_dependent_on_company(self):
        dimensions = with_employee_filter_dimensions(self.jira, "Jira")
        options = employee_filter_options(dimensions, "Acme")
        self.assertEqual(options["Project / Space"], ["All", "Alpha Project"])
        filtered = apply_employee_filters(dimensions, {"Company": "Acme", "Project / Space": ["Alpha Project"]})
        self.assertEqual(filtered["issue_key"].tolist(), ["A-1"])

    def test_all_post_analysis_dimensions_filter_without_mutating_snapshot(self):
        dimensions = with_employee_filter_dimensions(self.jira, "Jira")
        original = dimensions.copy()
        filtered = apply_employee_filters(dimensions, {
            "Company": "Beta",
            "Project / Space": ["Beta Project"],
            "Status": ["In Progress"],
            "Task Type": ["Bug"],
            "Priority": ["Low"],
            "Due-Date State": ["Open overdue"],
        })
        self.assertEqual(filtered["issue_key"].tolist(), ["B-1"])
        pd.testing.assert_frame_equal(dimensions, original)

    def test_missing_company_is_explicit(self):
        frame = self.jira.drop(columns=["company_name"])
        dimensions = with_employee_filter_dimensions(frame, "Jira")
        self.assertEqual(set(dimensions["Company"]), {"Company not specified"})

    def test_clickup_recalculation_uses_filtered_snapshot(self):
        tasks = [
            {
                "id": "1", "name": "Acme task", "employee_space_name": "Alpha Space",
                "company": "Acme", "status": {"status": "COMPLETE", "type": "done"},
                "task_type": "Task", "priority": {"priority": "high"},
                "date_created": "2026-09-01T09:00:00Z", "date_closed": "2026-09-02T09:00:00Z",
            },
            {
                "id": "2", "name": "Beta task", "employee_space_name": "Beta Space",
                "company": "Beta", "status": {"status": "IN PROGRESS"},
                "task_type": "Bug", "priority": {"priority": "low"},
                "date_created": "2026-09-01T09:00:00Z",
            },
        ]
        prepared = collect_data(Mock(), tasks, "Employee: Test", "employee-filter", time_status_data={})
        result = analyze_clickup(prepared)
        dimensions = with_employee_filter_dimensions(result["tasks"], "ClickUp")
        filtered = apply_employee_filters(dimensions, {"Company": "Acme"})
        narrowed = recalculate_clickup_analysis(result, filtered)
        overall = narrowed["overall"].set_index("Metric")["Value"]
        self.assertEqual(len(narrowed["tasks"]), 1)
        self.assertEqual(overall["Total tasks"], 1)
        self.assertEqual(overall["Completed tasks"], 1)
        self.assertEqual(overall["Completion rate (%)"], 100.0)
        self.assertEqual(result["overall"].set_index("Metric").loc["Total tasks", "Value"], 2)

    def test_current_employee_dashboard_filters_recalculate_selected_scope(self):
        snapshots = (
            TaskPeriodSnapshot(
                task=TaskRecord(
                    source_tool="Jira", task_id="J-1", source_space="Alpha Space",
                    unified_project="Employee: Test", company_name="Acme",
                    issue_type="Bug", priority="High", due_date=date(2026, 9, 10),
                ), period_start=date(2026, 9, 1), period_end=date(2026, 9, 21),
                status_at_period_end=UnifiedStatus.COMPLETED, in_scope=True,
                history_available=True, final_completion_date=date(2026, 9, 9),
            ),
            TaskPeriodSnapshot(
                task=TaskRecord(
                    source_tool="ClickUp", task_id="C-1", source_space="Beta Space",
                    unified_project="Employee: Test", company_name="Beta",
                    issue_type="Task", priority="Low", due_date=date(2026, 9, 5),
                ), period_start=date(2026, 9, 1), period_end=date(2026, 9, 21),
                status_at_period_end=UnifiedStatus.IN_EXECUTION, in_scope=True,
                history_available=False,
            ),
        )
        options = available_filters(snapshots)
        self.assertEqual(options.companies, ("Acme", "Beta"))
        self.assertEqual(options.source_spaces, ("Alpha Space", "Beta Space"))
        selected = filter_snapshots(snapshots, DashboardFilters(
            companies=("Acme",), source_spaces=("Alpha Space",),
            task_types=("Bug",), priorities=("High",), due_states=("Completed on time",),
        ))
        self.assertEqual([(item.task.source_tool, item.task.task_id) for item in selected], [("Jira", "J-1")])


if __name__ == "__main__":
    unittest.main()
