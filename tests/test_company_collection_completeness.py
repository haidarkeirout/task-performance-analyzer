import unittest
from types import SimpleNamespace
from unittest.mock import patch

from jira_gateway import CollectionError
from company_performance.collection import (
    ALL_COMPANIES_ID,
    COMPANY_SOURCE_MAPPINGS_KEY,
    CompanyCatalogEntry,
    CompanySourceRef,
    _project_names,
    collect_company_spaces,
    discover_company_catalog,
    remember_company_source_mapping,
)


class _Progress:
    def progress(self, *_args, **_kwargs):
        return None


class _Streamlit:
    def __init__(self):
        self.session_state = {}

    def progress(self, *_args, **_kwargs):
        return _Progress()


class CompanyCollectionCompletenessTests(unittest.TestCase):
    def test_catalog_merges_same_company_across_sources_and_keeps_source_only_companies(self):
        fake_st = _Streamlit()
        jira = {
            "NAST": {"id": "jira-1", "key": "NAST", "name": "Najm Al-Shamal [TEST]"},
        }
        clickup = {
            "cu-1": {"id": "cu-1", "name": "Najm Al-Shamal [TEST]"},
            "cu-2": {"id": "cu-2", "name": "Marchent [TEST]"},
        }
        settings = SimpleNamespace(revision="catalog-v1")
        with (
            patch("company_performance.collection.st", fake_st),
            patch("company_performance.collection._load_catalogs", return_value=(jira, clickup)),
        ):
            catalog = discover_company_catalog(settings)

        self.assertEqual([entry.company_name for entry in catalog], ["Marchent [TEST]", "Najm Al-Shamal [TEST]"])
        najm = next(entry for entry in catalog if entry.company_name.startswith("Najm"))
        self.assertEqual({source.source_tool for source in najm.sources}, {"Jira", "ClickUp"})
        self.assertEqual(najm.source_summary, "Jira + ClickUp")
        marchent = next(entry for entry in catalog if entry.company_name.startswith("Marchent"))
        self.assertEqual(marchent.source_summary, "ClickUp")
        self.assertEqual(marchent.sources[0].source_id, "cu-2")

    def test_selected_company_collects_only_its_stable_source_ids(self):
        fake_st = _Streamlit()
        spaces = {
            "najm": {"id": "cu-1", "name": "Najm Al-Shamal [TEST]"},
            "marchent": {"id": "cu-2", "name": "Marchent [TEST]"},
        }
        calls = []

        def collect(_settings, item, _fingerprint, progress=None):
            calls.append(item["id"])
            return SimpleNamespace(space_name=item["name"], collected_at="2026-09-30")

        settings = SimpleNamespace(revision="selected-v1")
        with (
            patch("company_performance.collection.st", fake_st),
            patch("company_performance.collection._load_catalogs", return_value=({}, spaces)),
            patch("company_performance.collection._collect_clickup_space", side_effect=collect),
        ):
            catalog = discover_company_catalog(settings)
            selected = next(entry for entry in catalog if entry.company_name.startswith("Marchent"))
            result = collect_company_spaces(settings, selected.company_id)

        self.assertEqual(calls, ["cu-2"])
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0].company_name, "Marchent [TEST]")
        self.assertEqual(result[0].source_id, "cu-2")

    def test_all_companies_collects_every_catalog_entry(self):
        fake_st = _Streamlit()
        spaces = {
            "najm": {"id": "cu-1", "name": "Najm Al-Shamal [TEST]"},
            "marchent": {"id": "cu-2", "name": "Marchent [TEST]"},
        }
        calls = []

        def collect(_settings, item, _fingerprint, progress=None):
            calls.append(item["id"])
            return SimpleNamespace(space_name=item["name"], collected_at="2026-09-30")

        settings = SimpleNamespace(revision="all-v1")
        with (
            patch("company_performance.collection.st", fake_st),
            patch("company_performance.collection._load_catalogs", return_value=({}, spaces)),
            patch("company_performance.collection._collect_clickup_space", side_effect=collect),
        ):
            result = collect_company_spaces(settings, ALL_COMPANIES_ID)

        self.assertEqual(set(calls), {"cu-1", "cu-2"})
        self.assertEqual({item.company_name for item in result}, {"Najm Al-Shamal [TEST]", "Marchent [TEST]"})

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

    def test_explicit_mapping_joins_differently_named_jira_and_clickup_sources(self):
        fake_st = _Streamlit()
        jira = {
            "NAST": {"id": "jira-project-7", "key": "NAST", "name": "Najm Al-Shamal Website"},
        }
        clickup = {
            "cu-1": {"id": "clickup-space-9", "name": "Najm Al-Shamal"},
        }
        settings = SimpleNamespace(revision="mapping-v1")
        fake_st.session_state[COMPANY_SOURCE_MAPPINGS_KEY] = {
            "jira:jira-project-7": {
                "company_id": "najm-al-shamal",
                "company_name": "Najm Al-Shamal",
                "aliases": ["Najm Al-Shamal Website", "Najm Al-Shamal"],
            },
            "clickup:clickup-space-9": {
                "company_id": "najm-al-shamal",
                "company_name": "Najm Al-Shamal",
                "aliases": ["Najm Al-Shamal Website", "Najm Al-Shamal"],
            },
        }
        with (
            patch("company_performance.collection.st", fake_st),
            patch("company_performance.collection._load_catalogs", return_value=(jira, clickup)),
        ):
            catalog = discover_company_catalog(settings)

        self.assertEqual(len(catalog), 1)
        self.assertEqual(catalog[0].company_name, "Najm Al-Shamal")
        self.assertEqual(
            {(source.source_tool, source.source_id) for source in catalog[0].sources},
            {("Jira", "jira-project-7"), ("ClickUp", "clickup-space-9")},
        )

    def test_mapping_helper_persists_platform_qualified_ids_and_aliases(self):
        fake_st = _Streamlit()
        source = CompanySourceRef("Jira", "NAST", "Najm Al-Shamal Website", {"key": "NAST"})
        target_source = CompanySourceRef("ClickUp", "901234", "Najm Al-Shamal", {"id": "901234"})
        target = CompanyCatalogEntry("najm-al-shamal", "Najm Al-Shamal", (target_source,))
        source_entry = CompanyCatalogEntry("najm-al-shamal-website", "Najm Al-Shamal Website", (source,))
        remember_company_source_mapping(fake_st.session_state, source_entry, target)

        registry = fake_st.session_state[COMPANY_SOURCE_MAPPINGS_KEY]
        self.assertEqual(registry["jira:NAST"]["company_id"], "najm-al-shamal")
        self.assertEqual(registry["clickup:901234"]["company_id"], "najm-al-shamal")
        self.assertIn("Najm Al-Shamal Website", registry["jira:NAST"]["aliases"])

    def test_differently_named_sources_stay_separate_without_explicit_mapping(self):
        fake_st = _Streamlit()
        settings = SimpleNamespace(revision="unmapped-v1")
        with (
            patch("company_performance.collection.st", fake_st),
            patch(
                "company_performance.collection._load_catalogs",
                return_value=(
                    {"NAST": {"id": "jira-1", "name": "Najm Al-Shamal Website"}},
                    {"cu-1": {"id": "clickup-1", "name": "Najm Al-Shamal"}},
                ),
            ),
        ):
            catalog = discover_company_catalog(settings)

        self.assertEqual(len(catalog), 2)
        self.assertEqual({entry.source_summary for entry in catalog}, {"Jira", "ClickUp"})

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
