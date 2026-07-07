from __future__ import annotations

import os
import re
from datetime import date, datetime
from typing import Any, Dict, List, Optional, Tuple

from parser.xlsx_adapter import Project, TaskComment, TaskInfo, MilestoneInfo


RAG_ORDER = {"Red": 0, "Amber": 1, "Green": 2, "Not Available": 3}


class RAGEngineError(RuntimeError):
    pass


def _today() -> date:
    return date.today()


def _coerce_date(value: Any) -> Optional[date]:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        s = value.strip()
        if not s:
            return None
        try:
            return datetime.fromisoformat(s).date()
        except ValueError:
            return None
    return None


def _normalize_text(value: Optional[str]) -> str:
    if not value:
        return ""
    return re.sub(r"\s+", " ", value.strip().lower())


def _infer_schedule_rag(project: Project) -> Dict[str, Any]:
    today = _today()
    overdue_todos = [t for t in project.tasks if t.finish_date and t.finish_date < today and (t.status or "").lower() not in {"completed", "done"}]
    active_tasks = [t for t in project.tasks if (t.status or "").lower() not in {"completed", "done"}]
    variance_values = [t.variance for t in project.tasks if t.variance is not None]
    percentage_values = [t.percent_complete for t in project.tasks if t.percent_complete is not None]

    late_milestones = [m for m in project.milestones if m.due_date and m.due_date < today]
    completed_pct = 0.0
    if percentage_values:
        completed_pct = sum(percentage_values) / len(percentage_values)

    score = 0.0
    if overdue_todos:
        score += 0.35
    if variance_values and any(v is not None and v < 0 for v in variance_values):
        score += 0.25
    if late_milestones:
        score += 0.2
    if completed_pct < 50.0:
        score += 0.2
    elif completed_pct < 80.0:
        score += 0.1

    rag = "Green"
    if score >= 0.6:
        rag = "Red"
    elif score >= 0.3:
        rag = "Amber"

    return {
        "rag": rag,
        "overdue_tasks": len(overdue_todos),
        "variance": round(sum(v for v in variance_values if v is not None) / len(variance_values), 2) if variance_values else 0.0,
        "late_milestones": len(late_milestones),
        "percent_complete": round(completed_pct, 2),
        "active_tasks": len(active_tasks),
        "evidence": [
            f"{len(overdue_todos)} overdue active task(s)",
            f"{len(late_milestones)} milestone(s) past due",
            f"{round(completed_pct, 1)}% overall completion",
        ],
    }


def _infer_milestone_health(project: Project) -> Dict[str, Any]:
    today = _today()
    completed = 0
    overdue = 0
    upcoming = 0
    missed = 0

    for milestone in project.milestones:
        if not milestone.due_date:
            continue
        if milestone.due_date < today:
            overdue += 1
        elif milestone.due_date == today:
            upcoming += 1
        else:
            upcoming += 1

    if project.milestones:
        completed = sum(1 for m in project.milestones if m.due_date is None)

    missed = max(0, overdue - completed)

    rag = "Green"
    if missed > 0 or overdue > 0:
        rag = "Red"
    elif upcoming > 0:
        rag = "Amber"

    return {
        "rag": rag,
        "completed": completed,
        "overdue": overdue,
        "upcoming": upcoming,
        "missed": missed,
        "evidence": [
            f"{completed} milestone(s) completed",
            f"{overdue} overdue",
            f"{upcoming} upcoming",
        ],
    }


def _infer_blockers(project: Project, schedule_signal: Dict[str, Any]) -> Dict[str, Any]:
    blocker_keywords = [
        "blocked",
        "pending",
        "dependency",
        "waiting",
        "issue",
        "risk",
        "escalation",
        "approval",
        "mapping",
    ]

    blocked_comments: List[str] = []
    for comment in project.comments:
        text = _normalize_text(comment.comment)
        if any(keyword in text for keyword in blocker_keywords):
            blocked_comments.append(text)

    count = 0
    reasons: List[str] = []
    if schedule_signal["rag"] == "Red":
        count += 1
        reasons.append("schedule health is red")

    overdue_active_tasks = [t for t in project.tasks if t.finish_date and t.finish_date < _today() and (t.status or "").lower() not in {"completed", "done"}]
    if overdue_active_tasks:
        count += len(overdue_active_tasks)
        reasons.append(f"{len(overdue_active_tasks)} overdue active task(s)")

    if blocked_comments:
        count += len(blocked_comments)
        reasons.append(f"{len(blocked_comments)} comment(s) indicate blockers")

    rag = "Green"
    if count >= 3:
        rag = "Red"
    elif count >= 1:
        rag = "Amber"

    return {
        "rag": rag,
        "count": count,
        "reasons": reasons,
        "evidence": reasons[:3],
    }


def _infer_sentiment(project: Project) -> Dict[str, Any]:
    comments = [comment.comment for comment in project.comments if comment.comment]
    if not comments:
        return {"rag": "Amber", "score": 0.0, "evidence": ["No comments available"], "method": "keyword"}

    api_key = os.getenv("GEMINI_API_KEY")
    if api_key:
        return {
            "rag": "Amber",
            "score": 0.0,
            "evidence": ["Gemini integration is not enabled in this deterministic engine"],
            "method": "gemini",
        }

    positive_words = {"good", "great", "positive", "success", "progress", "on track", "complete", "completed", "support"}
    negative_words = {"blocked", "pending", "issue", "risk", "delay", "escalation", "problem", "concern", "approval", "dependency", "waiting"}

    score = 0.0
    for comment in comments:
        text = _normalize_text(comment)
        positive_hits = sum(1 for word in positive_words if word in text)
        negative_hits = sum(1 for word in negative_words if word in text)
        score += positive_hits - negative_hits

    avg_score = max(-1.0, min(1.0, score / max(1, len(comments))))
    if avg_score >= 0.2:
        rag = "Green"
    elif avg_score <= -0.2:
        rag = "Red"
    else:
        rag = "Amber"

    return {"rag": rag, "score": round(avg_score, 2), "evidence": comments[:3], "method": "keyword"}


def _infer_budget(project: Project) -> Dict[str, Any]:
    budget_fields = [
        ("budget", getattr(project.metadata, "budget", None)),
    ]
    for field_name, value in budget_fields:
        if value is not None:
            return {"status": "Available", "confidence": 0.95, "evidence": [f"{field_name} present"]}

    return {"status": "Not Available", "confidence": 0.35, "evidence": ["No budget fields were present in the project model"]}


def _compose_recommendations(schedule: Dict[str, Any], milestone: Dict[str, Any], blockers: Dict[str, Any], sentiment: Dict[str, Any], budget: Dict[str, Any]) -> List[str]:
    recommendations: List[str] = []
    if schedule["rag"] == "Red":
        recommendations.append("Prioritize overdue tasks and recover schedule variance immediately.")
    elif schedule["rag"] == "Amber":
        recommendations.append("Tighten tracking on pending work to prevent schedule slippage.")

    if milestone["rag"] == "Red":
        recommendations.append("Revisit milestone dates and escalate any missed deliverables.")
    elif milestone["rag"] == "Amber":
        recommendations.append("Monitor upcoming milestones and confirm owners are aligned.")

    if blockers["count"] > 0:
        recommendations.append("Address blocker themes in comments and dependency risks.")

    if sentiment["rag"] == "Red":
        recommendations.append("Engage stakeholders to reduce concerns and clarify next steps.")
    elif sentiment["rag"] == "Amber":
        recommendations.append("Maintain communication with stakeholders to preserve confidence.")

    if budget["status"] == "Not Available":
        recommendations.append("Collect budget status and forecast data to strengthen financial oversight.")

    return recommendations


def _overall_rag_from_scores(signal_scores: Dict[str, Any]) -> Tuple[str, float]:
    rag_values = [signal_scores["schedule_health"]["rag"], signal_scores["milestone_health"]["rag"], signal_scores["blockers"]["rag"], signal_scores["stakeholder_sentiment"]["rag"]]
    weights = {"Red": 1.0, "Amber": 0.5, "Green": 0.0, "Not Available": 0.25}
    weighted = sum(weights.get(rag, 0.0) for rag in rag_values)
    avg_score = weighted / len(rag_values)

    if avg_score >= 0.75:
        overall_rag = "Red"
    elif avg_score >= 0.25:
        overall_rag = "Amber"
    else:
        overall_rag = "Green"

    confidence = 0.75
    if signal_scores["budget"]["status"] == "Not Available":
        confidence -= 0.2
    if signal_scores["stakeholder_sentiment"]["method"] == "gemini":
        confidence += 0.05
    return overall_rag, round(max(0.1, min(0.95, confidence)), 2)


def analyze_project_health(project: Project) -> Dict[str, Any]:
    if not isinstance(project, Project):
        raise RAGEngineError("project must be a Project instance")

    schedule_signal = _infer_schedule_rag(project)
    milestone_signal = _infer_milestone_health(project)
    blockers_signal = _infer_blockers(project, schedule_signal)
    sentiment_signal = _infer_sentiment(project)
    budget_signal = _infer_budget(project)

    signal_scores = {
        "schedule_health": schedule_signal,
        "milestone_health": milestone_signal,
        "blockers": blockers_signal,
        "stakeholder_sentiment": sentiment_signal,
        "budget": budget_signal,
    }

    overall_rag, confidence = _overall_rag_from_scores(signal_scores)
    recommendations = _compose_recommendations(schedule_signal, milestone_signal, blockers_signal, sentiment_signal, budget_signal)

    return {
        "overall_rag": overall_rag,
        "confidence": confidence,
        "signal_scores": signal_scores,
        "evidence": [
            f"Schedule: {schedule_signal['rag']}",
            f"Milestones: {milestone_signal['rag']}",
            f"Blockers: {blockers_signal['rag']}",
            f"Sentiment: {sentiment_signal['rag']}",
            f"Budget: {budget_signal['status']}",
        ],
        "recommendations": recommendations,
    }
