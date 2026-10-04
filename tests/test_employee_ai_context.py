import json
import unittest
from dataclasses import dataclass
from datetime import date

import pandas as pd

from employee_ai import (
    build_clickup_employee_context,
    build_combined_employee_context,
    build_employee_ai_context,
    build_jira_employee_context,
)
from employee_ai.sanitizer import assert_no_sensitive_fields
from employee_ai.schemas import ContextValidationError


SECRET = "test-only-subject-secret"
EMPLOYEE = {
    "display_name": "Test Employee",
    "department": "Technology",
    "position": "Developer",
    "subject_key": "internal-employee-key",
}
COLLECTION = {
    "collection_id": "collection-1",
    "collected_at": "2026-10-04T08:00:00+00:00",
    "analysis_completed_at": "2026-10-04T08:01:00+00:00",
}


class EmployeeAIContextSafetyTests(unittest.TestCase):
    def test_context_keeps_only_allowlisted_fields_and_redacts_embedded_email(self):
        frame = pd.DataFrame([{
            "issue_key": "SEC-1",
            "task_name": "Contact person@example.com",
            "status_at_cutoff": "In Progress",
            "priority": "High",
            "evaluation_cutoff": "2026-10-04",
            "status_known": True,
            "history_complete": True,
            "email": "person@example.com",
            "jira_account_id": "jira-sensitive-id",
            "access_token": "sensitive-token",
        }])
        context = build_jira_employee_context(
            employee=EMPLOYEE,
            collection=COLLECTION,
            dashboard_values={
                "total": 1, "completed": 0, "known_status_tasks": 1,
                "unknown_status_tasks": 0, "completion_rate": "0.0%",
                "on_time_rate": "N/A", "overdue": 0, "wip": 1,
                "on_time_tasks": 0, "on_time_valid_tasks": 0,
            },
            task_metrics=frame,
            subject_secret=SECRET,
        )
        payload = context.to_dict()
        encoded = json.dumps(payload)
        self.assertNotIn("person@example.com", encoded)
        self.assertNotIn("jira-sensitive-id", encoded)
        self.assertNotIn("sensitive-token", encoded)
        self.assertIn("[redacted-email]", encoded)
        assert_no_sensitive_fields(payload)

    def test_final_guard_fails_closed_on_a_forbidden_field(self):
        with self.assertRaises(ContextValidationError):
            assert_no_sensitive_fields({"analysis": {"clickup_user_id": "123"}})

    def test_fingerprint_changes_with_scope_but_not_with_question(self):
        common = dict(
            source_mode="jira",
            employee=EMPLOYEE,
            collection=COLLECTION,
            cards=[{"key": "total_tasks", "label": "Total Tasks", "display_value": "1"}],
            tasks=[{"source": "Jira", "task_key": "A-1"}],
            subject_secret=SECRET,
            request_id="fixed-request",
            sent_at="2026-10-04T08:02:00+00:00",
        )
        first = build_employee_ai_context(**common, question="First question").to_dict()
        second = build_employee_ai_context(**common, question="Second question").to_dict()
        filtered = build_employee_ai_context(
            **common,
            question="First question",
            analysis_filters={"company": "Acme"},
        ).to_dict()
        self.assertEqual(
            first["request"]["context_fingerprint"],
            second["request"]["context_fingerprint"],
        )
        self.assertNotEqual(
            first["request"]["context_fingerprint"],
            filtered["request"]["context_fingerprint"],
        )

    def test_fingerprint_changes_when_task_evidence_changes_but_count_does_not(self):
        common = dict(
            source_mode="jira",
            employee=EMPLOYEE,
            collection=COLLECTION,
            cards=[{"key": "total_tasks", "label": "Total Tasks", "display_value": "1"}],
            subject_secret=SECRET,
            request_id="fixed-request",
            sent_at="2026-10-04T08:02:00+00:00",
        )
        first = build_employee_ai_context(
            **common,
            tasks=[{"source": "Jira", "task_key": "A-1", "status": "In Progress"}],
        ).to_dict()
        changed = build_employee_ai_context(
            **common,
            tasks=[{"source": "Jira", "task_key": "A-1", "status": "Done"}],
        ).to_dict()
        self.assertNotEqual(
            first["request"]["context_fingerprint"],
            changed["request"]["context_fingerprint"],
        )

    def test_evidence_is_capped_and_marked_truncated(self):
        context = build_employee_ai_context(
            source_mode="clickup",
            employee=EMPLOYEE,
            collection=COLLECTION,
            cards=[],
            tasks=[{"source": "ClickUp", "task_key": str(index)} for index in range(4)],
            subject_secret=SECRET,
            max_evidence_rows=2,
        ).to_dict()
        self.assertEqual(context["evidence"]["task_count_in_scope"], 4)
        self.assertEqual(context["evidence"]["rows_included"], 2)
        self.assertTrue(context["evidence"]["truncated"])


class EmployeeAIContextAdapterTests(unittest.TestCase):
    def test_jira_uses_supplied_dashboard_values_without_recalculation(self):
        frame = pd.DataFrame([{
            "issue_key": "J-1", "task_name": "One row", "status_at_cutoff": "Done",
            "status_known": True, "history_complete": True,
            "evaluation_cutoff": "2026-10-04",
        }])
        context = build_jira_employee_context(
            employee=EMPLOYEE,
            collection=COLLECTION,
            dashboard_values={
                "total": 99, "completed": 40, "known_status_tasks": 80,
                "unknown_status_tasks": 19, "completion_rate": "50.0%",
                "on_time_rate": "75.0%", "overdue": 4, "wip": 9,
                "on_time_tasks": 30, "on_time_valid_tasks": 40,
            },
            task_metrics=frame,
            subject_secret=SECRET,
        ).to_dict()
        cards = {item["key"]: item for item in context["dashboard"]["cards"]}
        self.assertEqual(cards["total_tasks"]["numeric_value"], 99)
        self.assertEqual(cards["completion_rate"]["numeric_value"], 50)
        self.assertEqual(cards["completion_rate"]["components"]["denominator"], 80)

    def test_clickup_adapter_uses_calculated_overall_table(self):
        result = {
            "overall": pd.DataFrame([
                ("Total tasks", 3), ("Completed tasks", 2), ("Known status tasks", 2),
                ("Unknown status tasks", 1), ("Completion rate (%)", 100.0),
                ("On-time completion rate (%)", 50.0), ("Open overdue tasks", 0),
                ("WIP tasks", 0),
            ], columns=["Metric", "Value"]),
            "tasks": pd.DataFrame([{
                "Task ID": "C-1", "Task Name": "ClickUp task", "Current Status": "Complete",
                "Status Known?": True, "Priority": "High",
            }]),
            "period_start": "", "period_end": "", "cutoff": "2026-10-04",
            "source_timezone": "Asia/Damascus",
        }
        context = build_clickup_employee_context(
            employee=EMPLOYEE,
            collection=COLLECTION,
            analysis_result=result,
            subject_secret=SECRET,
        ).to_dict()
        self.assertEqual(context["analysis"]["source_mode"], "clickup")
        self.assertEqual(context["dashboard"]["cards"][2]["numeric_value"], 100)
        self.assertEqual(context["evidence"]["tasks"][0]["task_key"], "C-1")

    def test_combined_adapter_preserves_source_identity(self):
        @dataclass
        class Card:
            key: str
            title: str
            value: str
            supporting_text: str = ""

        @dataclass
        class Detail:
            source_tool: str
            source_space: str
            task_id: str
            task_name: str
            final_status: str
            priority: str
            due_date: date
            counted_in_kpis: bool
            exclusion_reason: str | None = None
            company_name: str | None = None
            data_quality_flags: tuple[str, ...] = ()

        @dataclass
        class Model:
            period_start: date
            period_end: date
            kpis: dict
            cards: tuple
            task_details: tuple
            source_coverage: tuple = ()
            data_quality: tuple = ()

        @dataclass
        class Result:
            model: Model
            bottlenecks: tuple = ()

        result = Result(Model(
            period_start=date(2026, 9, 1),
            period_end=date(2026, 10, 4),
            kpis={"total_tasks": 2},
            cards=(Card("total_tasks", "Total Tasks", "2"),),
            task_details=(
                Detail("Jira", "Jira Project", "J-1", "Jira task", "Completed", "High", date(2026, 9, 10), True),
                Detail("ClickUp", "ClickUp Space", "C-1", "ClickUp task", "In Execution", "Medium", date(2026, 10, 5), True),
            ),
        ))
        context = build_combined_employee_context(
            employee=EMPLOYEE,
            collection=COLLECTION,
            analysis_result=result,
            subject_secret=SECRET,
        ).to_dict()
        self.assertEqual(context["analysis"]["source_mode"], "combined")
        self.assertIn(
            "completion_population",
            {item["key"] for item in context["dashboard"]["cards"]},
        )
        self.assertEqual(
            {item["source"] for item in context["evidence"]["tasks"]},
            {"Jira", "ClickUp"},
        )


if __name__ == "__main__":
    unittest.main()
