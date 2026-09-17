"""Smoke tests for the current four-mode Streamlit entry point."""

import unittest
from pathlib import Path

from streamlit.testing.v1 import AppTest


ROOT = Path(__file__).resolve().parents[1]


def button(at, label):
    return next(item for item in at.button if item.label == label)


class CurrentAppSmokeTests(unittest.TestCase):
    def app(self):
        at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=30)
        at.secrets.update(
            APP_USERNAME="test-admin",
            APP_PASSWORD="test-password",
            JIRA_BASE_URL="https://test.atlassian.net",
            JIRA_EMAIL="test@example.invalid",
            JIRA_API_TOKEN="test-token",
        )
        at.run()
        self.assertEqual(len(at.exception), 0)
        return at

    def sign_in(self, at, password="test-password"):
        at.text_input(key="login_username").input("test-admin")
        at.text_input(key="login_password").input(password)
        button(at, "Sign In").click().run()
        self.assertEqual(len(at.exception), 0)

    def test_wrong_password_stays_before_any_data_collection(self):
        at = self.app()
        self.sign_in(at, password="wrong")
        self.assertTrue(any("Incorrect username" in item.value for item in at.error))
        self.assertNotIn("prepared_data", at.session_state)

    def test_company_mode_requires_explicit_company_choice_before_collection(self):
        at = self.app()
        self.sign_in(at)
        at.radio(key="analysis_type_selector").set_value("Company Performance").run()
        self.assertEqual(len(at.exception), 0)
        self.assertIsNotNone(at.selectbox(key="company_selected_company"))
        self.assertTrue(any("Choose a company" in item.value for item in at.info))


if __name__ == "__main__":
    unittest.main()
