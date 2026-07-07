import tempfile
import unittest
from pathlib import Path

from parser.rag_engine import analyze_project_health
from parser.xlsx_adapter import parse_project_xlsx
from monthly_synthesis import build_portfolio_summary, synthesize_folder


class MonthlySynthesisTests(unittest.TestCase):
    def test_builds_portfolio_summary_from_multiple_projects(self) -> None:
        project_a = parse_project_xlsx("d:/zycus/S2P Project (2).xlsx")
        project_b = parse_project_xlsx("d:/zycus/S2P Project (1).xlsx")
        analyses = [
            (project_a, analyze_project_health(project_a)),
            (project_b, analyze_project_health(project_b)),
        ]

        summary = build_portfolio_summary(analyses)

        self.assertIn("portfolio_health", summary)
        self.assertIn("health_distribution", summary)
        self.assertIn("cross_project_risk_themes", summary)
        self.assertIn("projects_requiring_escalation", summary)
        self.assertIn("executive_recommendations", summary)
        self.assertIn("projects", summary)

    def test_portfolio_health_uses_worst_project_rag_not_average_confidence(self) -> None:
        project_a = parse_project_xlsx("d:/zycus/S2P Project (2).xlsx")
        project_b = parse_project_xlsx("d:/zycus/S2P Project (1).xlsx")

        analyses = [
            (project_a, {"overall_rag": "Red", "confidence": 0.15, "signal_scores": {}, "evidence": [], "recommendations": []}),
            (project_b, {"overall_rag": "Green", "confidence": 0.95, "signal_scores": {}, "evidence": [], "recommendations": []}),
        ]

        summary = build_portfolio_summary(analyses)

        self.assertEqual(summary["portfolio_health"], "Red")
        self.assertEqual(len(summary["projects_requiring_escalation"]), 1)

    def test_synthesizes_folder_output(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            output_dir = Path(tmp_dir) / "monthly"
            summary = synthesize_folder(Path("d:/zycus"), output_dir=output_dir)
            self.assertIn("portfolio_health", summary)
            self.assertTrue((output_dir / "portfolio_summary.json").exists())


if __name__ == "__main__":
    unittest.main()
