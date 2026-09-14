import unittest
from pathlib import Path


class IntegratedAnalysisModeTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.source = (Path(__file__).parents[1] / "app.py").read_text(encoding="utf-8")

    def test_launcher_exposes_exactly_three_top_level_analysis_paths(self):
        self.assertIn(
            '["Employee & Project Analysis", "Department Performance", "Company Performance"]',
            self.source,
        )

    def test_company_launcher_is_isolated_to_company_mode(self):
        self.assertIn('if analysis_mode == "company":\n    remember_prepared_source(', self.source)
        self.assertIn('if analysis_mode != "company" and run_button', self.source)

    def test_mode_change_clears_cross_analysis_prepared_payloads(self):
        self.assertIn("invalidate_selection()", self.source)
        self.assertIn('"clickup_prepared_data"', self.source)
        self.assertIn('"company_prepared_clickup_spaces"', self.source)


if __name__ == "__main__":
    unittest.main()
