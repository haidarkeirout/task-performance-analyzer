"""Regression tests for date-only current-day Jira period handling."""

import unittest

from metrics_engine import WorkCalendar, calculate_task, parse_timestamp


class CurrentDayHistoryCoverageTests(unittest.TestCase):
    def setUp(self):
        self.calendar = WorkCalendar(timezone_name="Asia/Damascus")

    def _task(self):
        return {
            "issue_key": "TPA-1",
            "task_name": "Current-day task",
            "created_at": "2026-09-02T09:00:00+03:00",
            "current_status": "In Progress",
        }

    def test_history_from_same_calendar_day_is_accepted(self):
        cutoff = parse_timestamp(
            "2026-09-21T23:59:59.999999+03:00",
            self.calendar.timezone_name,
        )
        history = {
            "history_complete": True,
            "history_through": "2026-09-21T12:00:00+03:00",
            "initial_status": "In Progress",
            "snapshot": {"current_status": "In Progress"},
            "status_events": [],
        }

        result = calculate_task(
            self._task(),
            history,
            cutoff,
            self.calendar,
        )

        self.assertTrue(result["history_complete"])
        self.assertTrue(result["status_known"])
        self.assertEqual(result["status_at_cutoff"], "In Progress")
        self.assertTrue(result["is_open"])
        self.assertIsNotNone(result["task_age_elapsed_hours"])

    def test_history_from_previous_calendar_day_is_not_accepted(self):
        cutoff = parse_timestamp(
            "2026-09-21T23:59:59.999999+03:00",
            self.calendar.timezone_name,
        )
        history = {
            "history_complete": True,
            "history_through": "2026-09-20T23:59:59+03:00",
            "initial_status": "In Progress",
            "snapshot": {"current_status": "In Progress"},
            "status_events": [],
        }

        result = calculate_task(
            self._task(),
            history,
            cutoff,
            self.calendar,
        )

        self.assertFalse(result["history_complete"])
        self.assertFalse(result["status_known"])
        self.assertEqual(result["status_at_cutoff"], "Unavailable")


if __name__ == "__main__":
    unittest.main()
