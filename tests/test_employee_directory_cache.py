import unittest
from unittest.mock import patch

import employee_directory as directory


ROWS = [{
    "employee_name": "Ada",
    "department": "Engineering",
    "primary_source": "jira",
    "jira_account_id": "ada-1",
    "clickup_user_id": "",
    "active": "yes",
}]


class EmployeeDirectoryCacheTests(unittest.TestCase):
    def setUp(self):
        directory._DIRECTORY_CACHE.clear()
        directory._LAST_GOOD_DIRECTORY.clear()

    def test_reuses_valid_cache_without_redownloading(self):
        with (
            patch.object(directory, "_directory_url", return_value="https://example.test/directory"),
            patch.object(directory, "_remote_rows", return_value=ROWS) as remote,
        ):
            first, first_warning = directory.load_employee_directory_with_status()
            second, second_warning = directory.load_employee_directory_with_status()

        self.assertEqual(first, second)
        self.assertIsNone(first_warning)
        self.assertIsNone(second_warning)
        remote.assert_called_once()

    def test_uses_last_validated_copy_after_temporary_failure(self):
        url = "https://example.test/directory"
        with (
            patch.object(directory, "_directory_url", return_value=url),
            patch.object(directory, "_remote_rows", return_value=ROWS),
        ):
            expected, _ = directory.load_employee_directory_with_status()

        directory._DIRECTORY_CACHE.clear()
        with (
            patch.object(directory, "_directory_url", return_value=url),
            patch.object(directory, "_remote_rows", side_effect=ValueError("offline")),
        ):
            actual, warning = directory.load_employee_directory_with_status()

        self.assertEqual(actual, expected)
        self.assertIn("last successfully validated", warning)

    def test_ignores_inactive_rows_with_incomplete_routing_fields(self):
        rows = [
            *ROWS,
            {
                "employee_name": "Former employee",
                "department": "Engineering",
                "primary_source": "",
                "jira_account_id": "",
                "clickup_user_id": "",
                "active": "FALSE",
            },
        ]

        records = directory._build_records(rows, "fixture")

        self.assertEqual([record.name for record in records], ["Ada"])

    def test_normalizes_common_source_labels(self):
        rows = [
            {
                "employee_name": "Ada",
                "department": "Engineering",
                "primary_source": "Atlassian Jira",
                "jira_account_id": "ada-1",
                "clickup_user_id": "",
                "active": "TRUE",
            },
            {
                "employee_name": "Bea",
                "department": "Operations",
                "primary_source": "Click Up",
                "jira_account_id": "",
                "clickup_user_id": "bea-1",
                "active": "TRUE",
            },
        ]

        records = directory._build_records(rows, "fixture")

        self.assertEqual([record.primary_source for record in records], ["jira", "clickup"])

    def test_invalid_source_error_includes_received_value(self):
        rows = [
            {
                "employee_name": "Ada",
                "department": "Engineering",
                "primary_source": "Jira / ClickUp",
                "jira_account_id": "ada-1",
                "clickup_user_id": "",
                "active": "TRUE",
            }
        ]

        with self.assertRaisesRegex(ValueError, r"row 2.*Jira / ClickUp.*jira.*clickup"):
            directory._build_records(rows, "fixture")


if __name__ == "__main__":
    unittest.main()
