import unittest
from types import SimpleNamespace
from unittest.mock import patch

from jira_gateway import CollectionError
from company_performance.collection import _project_names, collect_company_spaces


class _Progress:
    def progress(self, *_args, **_kwargs):
        return None


class _Streamlit:
    def __init__(self):
        self.session_state = {}

    def progress(self, *_args, **_kwargs):
        return _Progress()


class CompanyCollectionCompletenessTests(unittest.TestCase):
    def test_cross_tool_match_is_kept_but_same_source_collision_is_not_merged(self):
        sources = [
            ("Jira", {"id": "1", "name": "Website Alpha"}),
            ("ClickUp", {"id": "2", "name": "Alpha"}),
        ]
        self.assertEqual(set(_project_names(sources).values()), {"Alpha"})

        ambiguous = [
            ("ClickUp", {"id": "1", "name": "Website Alpha"}),
            ("ClickUp", {"id": "2", "name": "Project Alpha"}),
        ]
        names = tuple(_project_names(ambiguous).values())
        self.assertEqual(len(set(names)), 2)
        self.assertTrue(all("[ClickUp:" in name for name in names))

    def test_failed_space_blocks_partial_dashboard_and_retry_reuses_success(self):
        fake_st = _Streamlit()
        spaces = {
            "one": {"id": "1", "name": "One"},
            "two": {"id": "2", "name": "Two"},
        }
        calls = []

        def collect(_settings, item, _fingerprint, progress=None):
            calls.append(item["id"])
            if item["id"] == "2" and calls.count("2") == 1:
                raise CollectionError("temporary failure")
            return SimpleNamespace(space_name=item["name"])

        with (
            patch("company_performance.collection.st", fake_st),
            patch("company_performance.collection._load_catalogs", return_value=(spaces, {})),
            patch("company_performance.collection._collect_jira_space", side_effect=collect),
        ):
            with self.assertRaisesRegex(CollectionError, "no partial dashboard"):
                collect_company_spaces(object())
            self.assertEqual(calls, ["1", "2"])
            self.assertEqual(len(fake_st.session_state["company_partial_prepared_items"]), 1)

            result = collect_company_spaces(object())

        self.assertEqual(calls, ["1", "2", "2"])
        self.assertEqual(len(result), 2)
        self.assertNotIn("company_partial_prepared_items", fake_st.session_state)


if __name__ == "__main__":
    unittest.main()
