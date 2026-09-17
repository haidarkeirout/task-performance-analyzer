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


if __name__ == "__main__":
    unittest.main()
