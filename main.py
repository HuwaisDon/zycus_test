from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, List, Optional

from parser.rag_engine import analyze_project_health
from parser.xlsx_adapter import parse_project_xlsx
from monthly_synthesis import synthesize_folder
from ppt_generator import generate_presentation_from_json
from report_generator import generate_reports_for_folder


def run_pipeline(
    input_path: Path | str,
    weekly_output_dir: Optional[Path | str] = None,
    monthly_output_dir: Optional[Path | str] = None,
    ppt_output_path: Optional[Path | str] = None,
) -> Dict[str, Any]:
    input_path = Path(input_path)
    weekly_output_dir = Path(weekly_output_dir) if weekly_output_dir is not None else Path("outputs/weekly")
    monthly_output_dir = Path(monthly_output_dir) if monthly_output_dir is not None else Path("outputs/monthly")
    ppt_output_path = Path(ppt_output_path) if ppt_output_path is not None else monthly_output_dir / "Executive_Project_Health_Report.pptx"

    if not input_path.exists():
        raise FileNotFoundError(input_path)

    workbook_paths = sorted(
        [p for p in input_path.iterdir() if p.is_file() and p.suffix.lower() in {".xlsx", ".xlsm", ".xls"}]
    )
    if not workbook_paths:
        raise FileNotFoundError(f"No Excel workbooks found in {input_path}")

    weekly_output_dir.mkdir(parents=True, exist_ok=True)
    monthly_output_dir.mkdir(parents=True, exist_ok=True)

    generated_reports = generate_reports_for_folder(input_path, output_dir=weekly_output_dir)
    summary = synthesize_folder(input_path, output_dir=monthly_output_dir)
    summary_path = monthly_output_dir / "portfolio_summary.json"
    ppt_path = generate_presentation_from_json(summary_path, output_path=ppt_output_path)

    return {
        "input_path": str(input_path),
        "generated_reports": len(generated_reports),
        "summary_path": str(summary_path),
        "ppt_path": str(ppt_path),
        "portfolio_health": summary.get("portfolio_health"),
    }


if __name__ == "__main__":
    result = run_pipeline(Path("d:/zycus"))
    print(json.dumps(result, indent=2))
