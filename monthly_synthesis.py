from __future__ import annotations

import argparse
import json
import os
import re
from collections import Counter
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

from parser.rag_engine import analyze_project_health
from parser.xlsx_adapter import Project, parse_project_xlsx


def _safe_text(value: Optional[Any]) -> str:
    if value is None:
        return "N/A"
    text = str(value).strip()
    return text if text else "N/A"


def _collect_projects(projects: Sequence[Project]) -> List[Dict[str, Any]]:
    return [
        {
            "project_name": _safe_text(project.metadata.project_name) or "Unnamed Project",
            "client": _safe_text(project.metadata.client_name) or "Unknown Client",
            "project_manager": _safe_text(project.metadata.project_manager) or "Unknown Manager",
            "overall_rag": analysis.get("overall_rag", "N/A"),
            "confidence": float(analysis.get("confidence", 0.0)),
            "signal_scores": analysis.get("signal_scores", {}),
            "evidence": analysis.get("evidence", []),
            "recommendations": analysis.get("recommendations", []),
        }
        for project, analysis in projects
    ]


def build_portfolio_summary(project_results: Sequence[Tuple[Project, Dict[str, Any]]]) -> Dict[str, Any]:
    if not project_results:
        return {
            "portfolio_health": "N/A",
            "health_distribution": {"Green": 0, "Amber": 0, "Red": 0},
            "cross_project_risk_themes": [],
            "projects_requiring_escalation": [],
            "executive_recommendations": [],
            "projects": [],
        }

    projects = _collect_projects(project_results)
    rag_counter = Counter(item["overall_rag"] for item in projects)
    health_distribution = {
        "Green": int(rag_counter.get("Green", 0)),
        "Amber": int(rag_counter.get("Amber", 0)),
        "Red": int(rag_counter.get("Red", 0)),
    }

    rag_priority = {"Red": 0, "Amber": 1, "Green": 2, "N/A": 3}
    portfolio_health = "Green"
    if projects:
        worst_rag = min((item["overall_rag"] for item in projects), key=lambda rag: rag_priority.get(rag, 99))
        portfolio_health = worst_rag if worst_rag in rag_priority else "N/A"

    cross_project_risk_themes: List[str] = []
    blocker_keywords = ["blocked", "pending", "dependency", "risk", "issue", "approval", "waiting", "mapping"]
    for item in projects:
        for evidence in item.get("evidence", []):
            text = evidence.lower()
            if any(keyword in text for keyword in blocker_keywords):
                cross_project_risk_themes.append(f"{item['project_name']}: {evidence}")

    if not cross_project_risk_themes:
        cross_project_risk_themes = ["No recurring risk themes detected from available evidence."]

    escalation_projects = [item for item in projects if item["overall_rag"] in {"Red", "Amber"} and item.get("confidence", 0.0) < 0.7]
    escalation_projects = sorted(escalation_projects, key=lambda item: item["project_name"].lower())

    recommendations = [
        "Prioritize recovery plans for red-rated projects and confirm accountable owners.",
        "Review recurring blocker themes across the portfolio for shared dependency risks.",
        "Increase governance cadence for projects with low confidence or delayed milestones.",
    ]

    return {
        "portfolio_health": portfolio_health,
        "health_distribution": health_distribution,
        "cross_project_risk_themes": cross_project_risk_themes[:10],
        "projects_requiring_escalation": [
            {
                "project_name": item["project_name"],
                "client": item["client"],
                "project_manager": item["project_manager"],
                "overall_rag": item["overall_rag"],
                "confidence": item["confidence"],
            }
            for item in escalation_projects
        ],
        "executive_recommendations": recommendations,
        "projects": projects,
    }


def write_json_summary(summary: Dict[str, Any], output_path: Path) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(summary, indent=2), encoding="utf-8")
    return output_path


def synthesize_folder(folder: Path, output_dir: Optional[Path] = None) -> Dict[str, Any]:
    folder = Path(folder)
    if not folder.exists():
        raise FileNotFoundError(folder)

    workbook_paths = sorted(
        [p for p in folder.iterdir() if p.is_file() and p.suffix.lower() in {".xlsx", ".xlsm", ".xls"}]
    )
    results: List[Tuple[Project, Dict[str, Any]]] = []
    for workbook_path in workbook_paths:
        project = parse_project_xlsx(workbook_path)
        results.append((project, analyze_project_health(project)))

    summary = build_portfolio_summary(results)
    if output_dir is not None:
        write_json_summary(summary, output_dir / "portfolio_summary.json")
    return summary


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Create deterministic portfolio synthesis for multiple projects")
    parser.add_argument("input", help="Path to a folder containing workbooks")
    parser.add_argument("--output-dir", default="outputs/monthly", help="Directory for generated portfolio summary JSON")
    args = parser.parse_args()

    output_dir = Path(args.output_dir)
    summary = synthesize_folder(Path(args.input), output_dir=output_dir)
    print(json.dumps(summary, indent=2))
