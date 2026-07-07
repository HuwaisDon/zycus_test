import tempfile
import unittest
from pathlib import Path

from parser.rag_engine import analyze_project_health
from parser.xlsx_adapter import parse_project_xlsx
from report_generator import generate_report_for_project, generate_reports_for_folder, render_markdown_report


class ReportGeneratorTests(unittest.TestCase):
    def test_renders_professional_markdown_report(self) -> None:
        project = parse_project_xlsx("d:/zycus/S2P Project (2).xlsx")
        analysis = analyze_project_health(project)

        report = render_markdown_report(project, analysis, report_date="2026-07-07")

        self.assertIn("# Weekly Project Health Report", report)
        self.assertIn("## Executive Summary", report)
        self.assertIn("## Signal Breakdown", report)
        self.assertIn("## Recommendations", report)
        self.assertIn("## Data Quality Notes", report)

    def test_generates_reports_for_folder(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            output_dir = Path(tmp_dir) / "reports"
            generated = generate_reports_for_folder(Path("d:/zycus"), output_dir=output_dir)
            self.assertTrue(generated)
            self.assertTrue(all(path.exists() for path in generated))


if __name__ == "__main__":
    unittest.main()
