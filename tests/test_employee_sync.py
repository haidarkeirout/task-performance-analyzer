import unittest
from types import SimpleNamespace
from unittest.mock import patch

from employee_sync import sync_employees
from jira_gateway import CollectionError


def _settings(*, jira=True, clickup=True):
    return SimpleNamespace(
        jira_url="https://example.atlassian.net" if jira else "",
        jira_email="admin@example.com" if jira else "",
        jira_token="jira-token" if jira else "",
        clickup_token="pk_clickup" if clickup else "",
        clickup_workspace_id="workspace-1",
    )


class EmployeeSyncTests(unittest.TestCase):
    def test_sync_returns_source_separate_records(self):
        jira_gateway = SimpleNamespace(
            users=lambda: [
                {"accountId": "jira-1", "displayName": "Ada", "active": True},
            ],
            close=lambda: None,
        )
        clickup_gateway = SimpleNamespace(
            workspace_members=lambda workspace_id: [
                {"user": {"id": 21, "username": "Ada", "status": "active"}},
                {"user": {"id": 22, "username": "Bea", "status": "active"}},
            ],
            close=lambda: None,
        )

        with patch("employee_sync.JiraGateway", return_value=jira_gateway), patch(
            "employee_sync.ClickUpGateway", return_value=clickup_gateway
        ):
            result = sync_employees(_settings())

        self.assertEqual(len(result.records), 3)
        self.assertEqual(
            [(record.name, record.primary_source) for record in result.records],
            [("Ada", "clickup"), ("Ada", "jira"), ("Bea", "clickup")],
        )
        self.assertEqual(result.records[0].clickup_user_id, "21")
        self.assertEqual(result.records[1].jira_account_id, "jira-1")
        self.assertEqual(result.warnings, ())

    def test_jira_app_accounts_are_not_employee_records(self):
        jira_gateway = SimpleNamespace(
            users=lambda: [
                {"accountId": "app-1", "displayName": "Atlas for Jira Cloud", "accountType": "app"},
                {"accountId": "human-1", "displayName": "Ada", "accountType": "atlassian"},
            ],
            close=lambda: None,
        )

        with patch("employee_sync.JiraGateway", return_value=jira_gateway):
            result = sync_employees(_settings(jira=True, clickup=False))

        self.assertEqual([record.name for record in result.records], ["Ada"])

    def test_clickup_member_payload_can_be_flat(self):
        clickup_gateway = SimpleNamespace(
            workspace_members=lambda workspace_id: [
                {"id": 7, "username": "Flat User", "status": "active"},
            ],
            close=lambda: None,
        )

        with patch("employee_sync.ClickUpGateway", return_value=clickup_gateway):
            result = sync_employees(_settings(jira=False))

        self.assertEqual(len(result.records), 1)
        self.assertEqual(result.records[0].name, "Flat User")
        self.assertEqual(result.records[0].clickup_user_id, "7")

    def test_source_failure_is_warning_when_other_source_succeeds(self):
        jira_gateway = SimpleNamespace(
            users=lambda: (_ for _ in ()).throw(CollectionError("jira unavailable")),
            close=lambda: None,
        )
        clickup_gateway = SimpleNamespace(
            workspace_members=lambda workspace_id: [
                {"user": {"id": 21, "username": "Bea", "status": "active"}},
            ],
            close=lambda: None,
        )

        with patch("employee_sync.JiraGateway", return_value=jira_gateway), patch(
            "employee_sync.ClickUpGateway", return_value=clickup_gateway
        ):
            with self.assertRaises(RuntimeError):
                sync_employees(_settings())

    def test_inactive_members_are_retained_but_not_active(self):
        clickup_gateway = SimpleNamespace(
            workspace_members=lambda workspace_id: [
                {"user": {"id": 21, "username": "Former", "status": "inactive"}},
            ],
            close=lambda: None,
        )

        with patch("employee_sync.ClickUpGateway", return_value=clickup_gateway):
            result = sync_employees(_settings(jira=False))

        self.assertEqual(len(result.records), 1)
        self.assertFalse(result.records[0].active)


if __name__ == "__main__":
    unittest.main()
