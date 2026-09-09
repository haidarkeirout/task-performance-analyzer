"""Regression tests for cutoff-anchored weekly dashboard filters."""
from __future__ import annotations

import ast
from pathlib import Path
import sys
import unittest

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from process_analysis import weekly_flow_summary
from clickup_analysis import _weekly_flow


class TimelineWindowTests(unittest.TestCase):
    def test_weekly_flow_extends_through_cutoff_with_zero_activity_weeks(self):
        cutoff = "2026-09-09T11:00:00Z"
        frame = pd.DataFrame([
            {
                "issue_key": "T-1",
                "created_at": "2026-07-20T08:00:00Z",
                "completed_at": "2026-07-21T08:00:00Z",
                "evaluation_cutoff": cutoff,
                "work_calendar_timezone": "Asia/Damascus",
            },
            {
                "issue_key": "T-2",
                "created_at": "2026-08-17T08:00:00Z",
                "completed_at": None,
                "evaluation_cutoff": cutoff,
                "work_calendar_timezone": "Asia/Damascus",
            },
        ])

        weekly = weekly_flow_summary(frame)
        last_four = weekly.tail(4).reset_index(drop=True)

        self.assertEqual(
            last_four["week_start"].tolist(),
            [
                pd.Timestamp("2026-08-17"),
                pd.Timestamp("2026-08-24"),
                pd.Timestamp("2026-08-31"),
                pd.Timestamp("2026-09-07"),
            ],
        )
        self.assertEqual(last_four["tasks_opened"].tolist(), [1, 0, 0, 0])
        self.assertEqual(last_four["tasks_completed"].tolist(), [0, 0, 0, 0])

    def test_clickup_weekly_flow_extends_through_cutoff_with_zero_activity_weeks(self):
        cutoff = pd.Timestamp("2026-09-09T11:00:00Z")
        frame = pd.DataFrame([
            {
                "Task ID": "C-1",
                "Created": pd.Timestamp("2026-07-20T08:00:00Z"),
                "Completed": pd.Timestamp("2026-07-21T08:00:00Z"),
            },
            {
                "Task ID": "C-2",
                "Created": pd.Timestamp("2026-08-17T08:00:00Z"),
                "Completed": pd.NaT,
            },
        ])

        weekly = _weekly_flow(frame, cutoff)
        last_four = weekly.tail(4).reset_index(drop=True)
        expected_weeks = [
            pd.Timestamp("2026-08-17T00:00:00Z"),
            pd.Timestamp("2026-08-24T00:00:00Z"),
            pd.Timestamp("2026-08-31T00:00:00Z"),
            pd.Timestamp("2026-09-07T00:00:00Z"),
        ]
        self.assertEqual(last_four["Week Starting"].tolist(), expected_weeks)
        self.assertEqual(last_four["Tasks Created"].tolist(), [1, 0, 0, 0])
        self.assertEqual(last_four["Tasks Completed"].tolist(), [0, 0, 0, 0])

    def test_dashboard_supports_4_12_24_36_48_week_filters(self):
        source = ast.parse((ROOT / "app.py").read_text(encoding="utf-8"))
        function = next(
            node for node in source.body
            if isinstance(node, ast.FunctionDef) and node.name == "show_weekly_task_flow"
        )

        class FakeStreamlit:
            def __init__(self, selected):
                self.selected = selected
                self.options = None
                self.chart = None

            def subheader(self, *args, **kwargs):
                pass

            def caption(self, *args, **kwargs):
                pass

            def info(self, *args, **kwargs):
                pass

            def selectbox(self, label, options, **kwargs):
                self.options = list(options)
                if self.selected not in self.options:
                    raise AssertionError(f"Missing timeline option: {self.selected}")
                return self.selected

            def line_chart(self, data, **kwargs):
                self.chart = data.copy()

            def dataframe(self, *args, **kwargs):
                pass

        weeks = pd.date_range("2025-10-06", periods=48, freq="7D")
        weekly = pd.DataFrame({
            "week_start": weeks,
            "tasks_opened": range(48),
            "tasks_completed": range(48),
            "net_flow": [0] * 48,
            "cumulative_net_flow": [0] * 48,
        })

        expected = {
            "Last 4 weeks": 4,
            "Last 12 weeks": 12,
            "Last 24 weeks": 24,
            "Last 36 weeks": 36,
            "Last 48 weeks": 48,
        }

        for label, count in expected.items():
            fake = FakeStreamlit(label)
            namespace = {"st": fake, "pd": pd}
            exec(
                compile(ast.Module(body=[function], type_ignores=[]), "app.py", "exec"),
                namespace,
            )
            namespace["show_weekly_task_flow"]({"weekly_flow": weekly})
            self.assertIsNotNone(fake.chart)
            self.assertEqual(len(fake.chart), count, label)


if __name__ == "__main__":
    unittest.main()
