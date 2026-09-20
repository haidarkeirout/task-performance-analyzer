import unittest
from types import SimpleNamespace
from unittest.mock import patch

import jira_ui


class _Streamlit:
    def __init__(self):
        self.session_state = {"login_username": "wrong", "login_password": "wrong"}


class LoginRateLimitTests(unittest.TestCase):
    def test_fifth_failure_locks_verification_for_five_minutes(self):
        fake_st = _Streamlit()
        settings = SimpleNamespace(revision="revision")
        with (
            patch.object(jira_ui, "st", fake_st),
            patch.object(jira_ui, "credentials_match", return_value=False) as match,
            patch.object(jira_ui.time, "time", return_value=1000.0),
        ):
            for _ in range(5):
                fake_st.session_state["login_password"] = "wrong"
                jira_ui._login(settings)
            jira_ui._login(settings)

        self.assertEqual(match.call_count, 5)
        self.assertEqual(fake_st.session_state["login_locked_until"], 1300.0)
        self.assertIn("Too many", fake_st.session_state["login_error"])


if __name__ == "__main__":
    unittest.main()
