"""Regression tests for Employee analysis date ranges."""

from datetime import date
import inspect
import json
from types import SimpleNamespace
from unittest import TestCase
from unittest.mock import patch

import employee_ui
from clickup_analysis import analyze_clickup
from clickup_export import ClickUpPreparedData


class EmployeeAnalysisPeriodTests(TestCase):
    def test_period_change_clears_analysis_but_keeps_prepared_snapshot(self):
        fake_streamlit = SimpleNamespace(
            session_state={
                "employee_period_start": date(2026, 9, 1),
                "employee_period_end": date(2026, 9, 21),
                "employee_prepared": object(),
                "employee_company_analysis": object(),
                "clickup_analysis": object(),
                "task_metrics": object(),
                "process_data": object(),
                "employee_output_cache": object(),
                "clickup_excel_report": object(),
            }
        )

        with patch.object(employee_ui, "st", fake_streamlit):
            employee_ui._on_employee_period_changed()

        self.assertEqual(fake_streamlit.session_state["employee_period_start"], date(2026, 9, 1))
        self.assertEqual(fake_streamlit.session_state["employee_period_end"], date(2026, 9, 21))
        self.assertIn("employee_prepared", fake_streamlit.session_state)
        for key in (
            "employee_company_analysis",
            "clickup_analysis",
            "task_metrics",
            "process_data",
            "employee_output_cache",
            "clickup_excel_report",
        ):
            self.assertNotIn(key, fake_streamlit.session_state)

    def test_employee_collection_runs_in_main_app_flow(self):
        source = inspect.getsource(employee_ui.render_employee_collection)

        self.assertNotIn("@st.fragment", source)

    def test_prepared_snapshot_is_kept_when_scope_marker_matches(self):
        fake_streamlit = SimpleNamespace(
            session_state={"employee_prepared_scope_fingerprint": "scope-123"}
        )
        prepared = object()

        with patch.object(employee_ui, "st", fake_streamlit):
            self.assertTrue(employee_ui._employee_prepared_matches_scope(prepared, "scope-123"))
            self.assertFalse(employee_ui._employee_prepared_matches_scope(prepared, "scope-456"))
            self.assertFalse(employee_ui._employee_prepared_matches_scope(None, "scope-123"))

    def test_clickup_analysis_uses_selected_period_without_recollection(self):
        tasks = [
            {
                "id": "in-period-completed",
                "name": "Completed in period",
                "status": {"status": "Complete", "type": "closed"},
                "assignees": [],
                "date_created": "2026-09-02T09:00:00+00:00",
                "date_updated": "2026-09-05T09:00:00+00:00",
                "date_closed": "2026-09-05T09:00:00+00:00",
            },
            {
                "id": "open-in-period",
                "name": "Open in period",
                "status": {"status": "In Progress", "type": "open"},
                "assignees": [],
                "date_created": "2026-08-01T09:00:00+00:00",
                "date_updated": "2026-09-10T09:00:00+00:00",
            },
            {
                "id": "completed-before-period",
                "name": "Completed before period",
                "status": {"status": "Complete", "type": "closed"},
                "assignees": [],
                "date_created": "2026-08-01T09:00:00+00:00",
                "date_updated": "2026-08-05T09:00:00+00:00",
                "date_closed": "2026-08-05T09:00:00+00:00",
            },
        ]
        prepared = ClickUpPreparedData(
            xlsx=b"",
            history_json=json.dumps({"tasks": tasks, "time_in_status": {}}).encode(),
            cutoff="2026-10-01T12:00:00+00:00",
            collected_at="2026-10-01T12:00:00+00:00",
            query="Employee: Test Employee",
            fingerprint="employee-test",
            count=3,
            filename="employee.xlsx",
            space_name="Test Space",
            source_timezone="Asia/Damascus",
        )

        result = analyze_clickup(
            prepared,
            period_start=date(2026, 9, 1),
            period_end=date(2026, 9, 21),
        )

        self.assertEqual(
            set(result["tasks"]["Task ID"]),
            {"in-period-completed", "open-in-period"},
        )
        self.assertEqual(result["period_start"], "2026-09-01")
        self.assertEqual(result["period_end"], "2026-09-21")
        self.assertEqual(result["analysis_context"].loc[
            result["analysis_context"]["Field"].eq("Evaluation Period Start"), "Value"
        ].iloc[0], "2026-09-01")
        self.assertEqual(result["analysis_context"].loc[
            result["analysis_context"]["Field"].eq("Evaluation Period End"), "Value"
        ].iloc[0], "2026-09-21")
