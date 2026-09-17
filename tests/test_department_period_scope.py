from __future__ import annotations

from datetime import date
import unittest

import pandas as pd

from department_analysis import filter_jira_department_period
from department_collection import _period_bounds, _period_filter
from jira_department_collection import _period_query


def _milliseconds(value: str) -> str:
    return str(int(pd.Timestamp(value).timestamp() * 1000))


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


if __name__ == "__main__":
    unittest.main()
