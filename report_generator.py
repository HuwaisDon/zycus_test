from __future__ import annotations

import argparse
import os
import re
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional

from parser.rag_engine import analyze_project_health
from parser.xlsx_adapter import Project, parse_project_xlsx


def _safe_text(value: Optional[Any]) -> str:
    if value is None:
        return "N/A"
    text = str(value).strip()
    return text if text else "N/A"


def _format_date(value: Optional[Any]) -> str:
    if value is None:
        return "N/A"
    if hasattr(value, "strftime"):
        return value.strftime("%Y-%m-%d")
    return _safe_text(value)


def _build_signal_table(signal_scores: Dict[str, Any]) -> str:
    rows = [
        "| Signal | Status | Evidence |",
        "| --- | --- | --- |",
    ]
    for name, signal in signal_scores.items():
        label = name.replace("_", " ").title()
        status = signal.get("rag", "N/A")
        evidence = signal.get("evidence") or signal.get("reasons") or []
        if isinstance(evidence, list):
            evidence_text = "; ".join(str(item) for item in evidence[:3])
        else:
            evidence_text = str(evidence)
        rows.append(f"| {label} | {status} | {evidence_text} |")
    return "\n".join(rows)


def render_markdown_report(project: Project, analysis: Dict[str, Any], report_date: Optional[Any] = None) -> str:
    metadata = project.metadata
    report_date_value = report_date or date.today()
    confidence = analysis.get("confidence", 0.0)
    overall_rag = analysis.get("overall_rag", "N/A")

    signal_scores = analysis.get("signal_scores", {})
    recommendations = analysis.get("recommendations", [])
    evidence = analysis.get("evidence", [])

    lines: List[str] = []
    lines.append("# Weekly Project Health Report")
    lines.append("")
    lines.append("## Project")
    lines.append("")
    lines.append(f"- Project: {_safe_text(metadata.project_name)}")
    lines.append(f"- Client: {_safe_text(metadata.client_name)}")
    lines.append(f"- Project Manager: {_safe_text(metadata.project_manager)}")
    lines.append(f"- Date: {_format_date(report_date_value)}")
    lines.append("")
    lines.append("## Overall Status")
    lines.append("")
    lines.append(f"- Status: **{overall_rag}**")
    lines.append(f"- Confidence: **{confidence:.2f}**")
    lines.append("")
    lines.append("## Executive Summary")
    lines.append("")
    lines.append(
        f"This report summarizes the current delivery posture for {_safe_text(metadata.project_name)}. "
        f"The overall health is assessed as **{overall_rag}** with confidence **{confidence:.2f}**."
    )
    lines.append("")
    lines.append("## Signal Breakdown")
    lines.append("")
    lines.append(_build_signal_table(signal_scores))
    lines.append("")
    lines.append("### Schedule")
    lines.append("")
    schedule_signal = signal_scores.get("schedule_health", {})
    lines.append(f"- Status: {schedule_signal.get('rag', 'N/A')}")
    lines.append(f"- Evidence: {', '.join(str(item) for item in schedule_signal.get('evidence', []))}")
    lines.append("")
    lines.append("### Milestones")
    lines.append("")
    milestone_signal = signal_scores.get("milestone_health", {})
    lines.append(f"- Status: {milestone_signal.get('rag', 'N/A')}")
    lines.append(f"- Evidence: {', '.join(str(item) for item in milestone_signal.get('evidence', []))}")
    lines.append("")
    lines.append("### Blockers")
    lines.append("")
    blocker_signal = signal_scores.get("blockers", {})
    lines.append(f"- Status: {blocker_signal.get('rag', 'N/A')}")
    lines.append(f"- Evidence: {', '.join(str(item) for item in blocker_signal.get('evidence', []))}")
    lines.append("")
    lines.append("### Sentiment")
    lines.append("")
    sentiment_signal = signal_scores.get("stakeholder_sentiment", {})
    lines.append(f"- Status: {sentiment_signal.get('rag', 'N/A')}")
    lines.append(f"- Evidence: {', '.join(str(item) for item in sentiment_signal.get('evidence', []))}")
    lines.append("")
    lines.append("### Budget")
    lines.append("")
    budget_signal = signal_scores.get("budget", {})
    lines.append(f"- Status: {budget_signal.get('status', 'N/A')}")
    lines.append(f"- Confidence: {budget_signal.get('confidence', 0.0):.2f}")
    lines.append("")
    lines.append("## Key Evidence")
    lines.append("")
    for item in evidence:
        lines.append(f"- {item}")
    lines.append("")
    lines.append("## Recommendations")
    lines.append("")
    for item in recommendations:
        lines.append(f"- {item}")
    lines.append("")
    lines.append("## Next Actions")
    lines.append("")
    lines.append("- Review overdue work and recovery plans.")
    lines.append("- Confirm milestone dates and owners.")
    lines.append("- Resolve blocker themes surfaced in comments.")
    lines.append("- Validate budget visibility and forecast assumptions.")
    lines.append("")
    lines.append("## Data Quality Notes")
    lines.append("")
    lines.append("- Report content is based on the parsed workbook and deterministic health signals.")
    lines.append("- Budget information is reported as Not Available when the project model does not include it.")
    lines.append("")
    return "\n".join(lines)


def generate_report_for_project(project: Project, analysis: Dict[str, Any], output_dir: Optional[Path] = None, report_date: Optional[Any] = None) -> Path:
    output_dir = output_dir or Path("outputs/weekly")
    output_dir.mkdir(parents=True, exist_ok=True)
    project_name = _safe_text(project.metadata.project_name) or "project"
    safe_name = re.sub(r"[^A-Za-z0-9._-]+", "_", project_name).strip("._") or "project"
    report_path = output_dir / f"{safe_name}_weekly_report.md"
    report_path.write_text(render_markdown_report(project, analysis, report_date=report_date), encoding="utf-8")
    return report_path


def generate_reports_for_folder(folder: Path, output_dir: Optional[Path] = None) -> List[Path]:
    folder = Path(folder)
    if not folder.exists():
        raise FileNotFoundError(folder)

    workbook_paths = sorted(
        [p for p in folder.iterdir() if p.is_file() and p.suffix.lower() in {".xlsx", ".xlsm", ".xls"}]
    )
    generated_reports: List[Path] = []
    for workbook_path in workbook_paths:
        project = parse_project_xlsx(workbook_path)
        analysis = analyze_project_health(project)
        generated_reports.append(generate_report_for_project(project, analysis, output_dir=output_dir))
    return generated_reports


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate Markdown project health reports")
    parser.add_argument("input", help="Path to a workbook or a folder containing workbooks")
    parser.add_argument("--output-dir", default="outputs/weekly", help="Directory for generated markdown reports")
    args = parser.parse_args()

    input_path = Path(args.input)
    output_dir = Path(args.output_dir)

    if input_path.is_dir():
        generated = generate_reports_for_folder(input_path, output_dir=output_dir)
    else:
        project = parse_project_xlsx(input_path)
        analysis = analyze_project_health(project)
        generated = [generate_report_for_project(project, analysis, output_dir=output_dir)]

    print("Generated reports:")
    for path in generated:
        print(path)
