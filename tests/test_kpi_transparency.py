import unittest
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from kpi_transparency import build_population


class CompletionRateTransparencyTests(unittest.TestCase):
    def test_population_keeps_rate_denominator_and_reasons_consistent(self):
        population = build_population(
            23,
            21,
            {"Subtask": 1, "Unknown status/history": 1},
        )
        self.assertEqual(population.total_tasks, 23)
        self.assertEqual(population.included_tasks, 21)
        self.assertEqual(population.excluded_tasks, 2)
        self.assertEqual(dict(population.reasons), {
            "Subtask": 1,
            "Unknown status/history": 1,
        })
        self.assertEqual(population.tooltip(), "2 task(s) excluded from Completion Rate")

    def test_unexplained_exclusions_are_visible_instead_of_disappearing(self):
        population = build_population(5, 3, {"Subtask": 1})
        self.assertEqual(dict(population.reasons), {
            "Other KPI exclusion": 1,
            "Subtask": 1,
        })


if __name__ == "__main__":
    unittest.main()
