import json
import tempfile
import unittest
from pathlib import Path

from ppt_generator import build_presentation, generate_presentation_from_json


class PPTGeneratorTests(unittest.TestCase):
    def test_build_presentation_from_summary(self) -> None:
        summary = {
            "portfolio_name": "Sample Portfolio",
            "portfolio_health": "Amber",
            "health_distribution": {"Green": 1, "Amber": 2, "Red": 0},
            "projects_requiring_escalation": [{"project_name": "Alpha", "client": "Client A", "project_manager": "PM A", "overall_rag": "Red", "confidence": 0.6}],
            "cross_project_risk_themes": ["Theme A", "Theme B"],
            "executive_recommendations": ["Recommendation 1", "Recommendation 2"],
            "projects": [{"project_name": "Alpha"}],
        }
        with tempfile.TemporaryDirectory() as tmp_dir:
            output_path = Path(tmp_dir) / "report.pptx"
            created = build_presentation(summary, output_path)
            self.assertTrue(created.exists())

    def test_generate_presentation_from_json(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            summary_path = Path(tmp_dir) / "portfolio_summary.json"
            summary_path.write_text(json.dumps({"portfolio_health": "Green", "health_distribution": {"Green": 1, "Amber": 0, "Red": 0}, "projects_requiring_escalation": [], "cross_project_risk_themes": ["None"], "executive_recommendations": ["Keep going"], "projects": [], "portfolio_name": "Demo"}), encoding="utf-8")
            output_path = generate_presentation_from_json(summary_path)
            self.assertTrue(output_path.exists())


if __name__ == "__main__":
    unittest.main()
