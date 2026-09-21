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

    def test_same_day_snapshot_status_is_usable_when_history_chain_breaks(self):
        cutoff = parse_timestamp(
            "2026-09-21T23:59:59.999999+03:00",
            self.calendar.timezone_name,
        )
        history = {
            "history_complete": True,
            "history_through": "2026-09-21T12:00:00+03:00",
            "initial_status": "To Do",
            "snapshot": {"current_status": "Done"},
            "status_events": [
                {
                    "from_status": "To Do",
                    "to_status": "In Progress",
                    "changed_at": "2026-09-03T09:00:00+03:00",
                },
                {
                    "from_status": "Review",
                    "to_status": "Done",
                    "changed_at": "2026-09-04T09:00:00+03:00",
                },
            ],
        }

        result = calculate_task(
            {**self._task(), "current_status": "Done"},
            history,
            cutoff,
            self.calendar,
        )

        self.assertFalse(result["history_complete"])
        self.assertTrue(result["status_known"])
        self.assertEqual(result["status_at_cutoff"], "Done")
        self.assertTrue(result["is_completed"])
        self.assertFalse(result["is_open"])
        self.assertIsNone(result["completed_at"])
        self.assertIn("Current Jira status is usable", result["history_note"])

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
