from __future__ import annotations

from datetime import date
import json
from types import SimpleNamespace
import unittest

import pandas as pd

from clickup_analysis import analyze_clickup
from department_analysis import filter_jira_department_period
from department_collection import (
    _analysis_cutoff, _period_bounds, _period_filter,
    _update_department_preparation_progress,
)
from jira_department_collection import _overall_progress_fraction, _period_query


def _milliseconds(value: str) -> str:
    return str(int(pd.Timestamp(value).timestamp() * 1000))


class _ProgressRecorder:
    def __init__(self):
        self.updates = []

    def progress(self, value, text):
        self.updates.append((value, text))


class DepartmentPeriodScopeTests(unittest.TestCase):
    def test_local_day_boundaries_are_converted_from_damascus_to_utc(self):
        start, end = _period_bounds(
            date(2026, 9, 1), date(2026, 9, 1), "Asia/Damascus"
        )
        self.assertEqual(start, pd.Timestamp("2026-08-31T21:00:00Z"))
        self.assertEqual(end.date(), date(2026, 9, 1))
        self.assertEqual(end.hour, 20)

    def test_task_just_after_local_midnight_stays_in_the_local_day(self):
        task = {
            "id": "local-midnight",
            "date_created": _milliseconds("2026-08-31T21:30:00Z"),
            "status": {"status": "in progress", "type": "custom"},
        }
        kept, missing = _period_filter(
            [task], date(2026, 9, 1), date(2026, 9, 1), "Asia/Damascus"
        )
        self.assertEqual([item["id"] for item in kept], ["local-midnight"])
        self.assertEqual(missing, 0)

    def test_open_task_created_before_period_remains_in_scope(self):
        task = {
            "id": "older-open",
            "date_created": _milliseconds("2026-08-10T09:00:00Z"),
            "status": {"status": "in progress", "type": "custom"},
        }
        kept, _ = _period_filter(
            [task], date(2026, 9, 1), date(2026, 9, 30), "Asia/Damascus"
        )
        self.assertEqual([item["id"] for item in kept], ["older-open"])

    def test_task_closed_before_period_is_excluded(self):
        task = {
            "id": "old-closed",
            "date_created": _milliseconds("2026-08-01T09:00:00Z"),
            "date_closed": _milliseconds("2026-08-20T09:00:00Z"),
            "status": {"status": "complete", "type": "closed"},
        }
        kept, _ = _period_filter(
            [task], date(2026, 9, 1), date(2026, 9, 30), "Asia/Damascus"
        )
        self.assertEqual(kept, [])

    def test_task_completed_after_period_remains_in_scope(self):
        task = {
            "id": "closed-later",
            "date_created": _milliseconds("2026-08-01T09:00:00Z"),
            "date_closed": _milliseconds("2026-10-05T09:00:00Z"),
            "status": {"status": "complete", "type": "closed"},
        }
        kept, _ = _period_filter(
            [task], date(2026, 9, 1), date(2026, 9, 30), "Asia/Damascus"
        )
        self.assertEqual([item["id"] for item in kept], ["closed-later"])

    def test_department_cutoff_never_extends_beyond_collection_time(self):
        cutoff = _analysis_cutoff(
            date(2026, 9, 1), date(2099, 9, 30), "Asia/Damascus"
        )
        self.assertLessEqual(cutoff, pd.Timestamp.now(tz="UTC"))

    def test_later_clickup_completion_is_not_counted_at_historical_cutoff(self):
        task = {
            "id": "closed-later",
            "name": "Closed after cutoff",
            "date_created": _milliseconds("2026-09-10T09:00:00Z"),
            "date_closed": _milliseconds("2026-10-05T09:00:00Z"),
            "status": {"status": "complete", "type": "closed"},
            "assignees": [],
        }
        prepared = SimpleNamespace(
            history_json=json.dumps({"tasks": [task], "time_in_status": {}}).encode(),
            cutoff="2026-09-30T20:59:59Z",
            collected_at="2026-10-10T09:00:00Z",
            source_timezone="Asia/Damascus",
            space_name="Operations",
            space_names=["Operations"],
            filter_summary="Department period",
            period_start="2026-09-01",
            period_end="2026-09-30",
            duplicate_count=0,
        )
        row = analyze_clickup(prepared)["tasks"].iloc[0]
        self.assertFalse(bool(row["Completed?"]))
        self.assertFalse(bool(row["Status Known?"]))
        self.assertEqual(row["Status at Cutoff"], "Unknown")

    def test_jira_query_does_not_discard_tasks_created_before_period(self):
        query = _period_query("ENG", date(2026, 9, 1), date(2026, 9, 30))
        self.assertIn('created <= "2026-09-30"', query)
        self.assertNotIn('created >= "2026-09-01"', query)
        self.assertIn('updated >= "2026-09-01"', query)

    def test_jira_scope_keeps_old_open_and_reopened_work(self):
        frame = pd.DataFrame([
            {
                "issue_key": "OLD-CLOSED",
                "created_at": "2026-08-01T09:00:00Z",
                "completed_at": "2026-08-20T09:00:00Z",
                "is_open": False,
            },
            {
                "issue_key": "OLD-OPEN",
                "created_at": "2026-08-01T09:00:00Z",
                "completed_at": None,
                "is_open": True,
            },
            {
                "issue_key": "REOPENED",
                "created_at": "2026-08-01T09:00:00Z",
                "completed_at": "2026-09-15T09:00:00Z",
                "is_open": False,
            },
        ])
        histories = {
            "REOPENED": {
                "status_events": [{
                    "changed_at": "2026-09-03T09:00:00Z",
                    "from_status": "Done",
                    "to_status": "In Progress",
                }]
            }
        }
        scoped = filter_jira_department_period(
            frame,
            histories,
            date(2026, 9, 1),
            "2026-09-30T20:59:59Z",
            "Asia/Damascus",
        )
        self.assertEqual(set(scoped["issue_key"]), {"OLD-OPEN", "REOPENED"})

    def test_jira_department_progress_maps_each_space_into_one_full_bar(self):
        self.assertEqual(_overall_progress_fraction(0, 2, 0.0), 0.0)
        self.assertEqual(_overall_progress_fraction(0, 2, 1.0), 0.5)
        self.assertEqual(_overall_progress_fraction(1, 2, 0.5), 0.75)
        self.assertEqual(_overall_progress_fraction(1, 2, 1.0), 1.0)

    def test_clickup_department_preparation_shows_completed_task_count(self):
        progress = _ProgressRecorder()

        _update_department_preparation_progress(
            progress, "Preparing ClickUp task 3 of 25...", 25
        )

        self.assertEqual(progress.updates, [(0.12, "Preparing department tasks: 3 / 25")])


if __name__ == "__main__":
    unittest.main()
