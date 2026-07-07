from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import pandas as pd

logger = logging.getLogger(__name__)


# -----------------------------
# Normalized domain dataclasses
# -----------------------------


@dataclass(frozen=True)
class ProjectMetadata:
    project_name: Optional[str] = None
    client_name: Optional[str] = None
    project_manager: Optional[str] = None
    report_date: Optional[date] = None


@dataclass(frozen=True)
class TaskComment:
    task_uid: Optional[str] = None
    task_name: Optional[str] = None
    author: Optional[str] = None
    comment_date: Optional[date] = None
    comment: Optional[str] = None


@dataclass(frozen=True)
class TaskInfo:
    task_uid: Optional[str] = None
    task_name: Optional[str] = None

    status: Optional[str] = None
    start_date: Optional[date] = None
    finish_date: Optional[date] = None
    duration: Optional[float] = None
    percent_complete: Optional[float] = None

    schedule_health: Optional[str] = None
    total_float: Optional[float] = None
    critical: Optional[bool] = None
    hierarchy_level: Optional[int] = None
    variance: Optional[float] = None

    predecessors: List[str] = field(default_factory=list)
    successors: List[str] = field(default_factory=list)


@dataclass(frozen=True)
class PhaseInfo:
    phase_name: Optional[str] = None
    hierarchy_level: Optional[int] = None


@dataclass(frozen=True)
class MilestoneInfo:
    milestone_name: Optional[str] = None
    due_date: Optional[date] = None
    hierarchy_level: Optional[int] = None


@dataclass(frozen=True)
class SummaryStatistics:
    num_tasks: int = 0
    num_milestones: int = 0
    num_comments: int = 0

    num_phases: int = 0


@dataclass(frozen=True)
class Project:
    metadata: ProjectMetadata
    tasks: List[TaskInfo] = field(default_factory=list)
    phases: List[PhaseInfo] = field(default_factory=list)
    milestones: List[MilestoneInfo] = field(default_factory=list)
    comments: List[TaskComment] = field(default_factory=list)
    summary: SummaryStatistics = field(default_factory=SummaryStatistics)


# -----------------------------
# Helpers
# -----------------------------

_DATE_LIKE_RE = re.compile(r"\b(\d{4})[-/](\d{1,2})[-/](\d{1,2})\b")



def _coerce_date(value: Any) -> Optional[date]:
    """Coerce excel-ish date/time or strings into a date.

    Never raises.
    """
    if value is None:
        return None

    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    # pandas Timestamp
    try:
        if hasattr(value, "to_pydatetime"):
            dt = value.to_pydatetime()
            return dt.date()
    except Exception:
        pass

    if isinstance(value, (int, float)) and not pd.isna(value):
        # Excel serial date (best-effort). openpyxl/pandas often already converts,
        # but if we land here, try serial.
        try:
            origin = datetime(1899, 12, 30)
            return (origin + pd.to_timedelta(value, unit="D")).date()
        except Exception:
            return None

    if isinstance(value, str):
        s = value.strip()
        if not s or s.upper() in {"#N/A", "#NA", "#UNPARSEABLE", "N/A"}:
            return None

        # Try pandas fast-path
        try:
            parsed = pd.to_datetime(s, errors="coerce", dayfirst=False)
            if pd.isna(parsed):
                # Try dayfirst
                parsed = pd.to_datetime(s, errors="coerce", dayfirst=True)
            if not pd.isna(parsed):
                return parsed.date()
        except Exception:
            pass

        m = _DATE_LIKE_RE.search(s)
        if m:
            try:
                y, mo, d = map(int, m.groups())
                return date(y, mo, d)
            except Exception:
                return None

    return None


def _coerce_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, str):
        s = value.strip()
        if not s or s.upper() in {"#UNPARSEABLE", "#N/A", "N/A"}:
            return None
        # Handle percent strings like "34%"
        if s.endswith("%"):
            try:
                return float(s[:-1].strip())
            except Exception:
                return None

    try:
        if pd.isna(value):
            return None
    except Exception:
        pass

    try:
        return float(value)
    except Exception:
        return None


def _coerce_bool(value: Any) -> Optional[bool]:
    if value is None:
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)) and not pd.isna(value):
        if int(value) == 1:
            return True
        if int(value) == 0:
            return False
    if isinstance(value, str):
        s = value.strip().upper()
        if s in {"YES", "Y", "TRUE", "T"}:
            return True
        if s in {"NO", "N", "FALSE", "F"}:
            return False
        if s in {"1", "0"}:
            return s == "1"
    return None


def _normalize_colname(col: Any) -> str:
    s = str(col) if col is not None else ""
    s = s.strip().lower()
    s = re.sub(r"[\s\-_/]+", "_", s)
    s = re.sub(r"[^a-z0-9_]+", "", s)
    s = re.sub(r"_+", "_", s)
    return s


# Column synonyms by normalized key
_COL_SYNONYMS: Dict[str, Sequence[str]] = {
    "project_name": ["project_name", "project", "project_title"],
    # In some workbooks client/account may not be a dedicated column.
    # We'll also accept values from any "customer/account/client"-like column names.
    "client_name": ["client_name", "client", "customer", "account"],

    "project_manager": ["project_manager", "manager", "pm", "owner"],
    "report_date": ["report_date", "date", "report"],
    # task
    "task_name": ["task_name", "task", "name", "resource_name"],
    "task_uid": ["task_uid", "uid", "id", "task_id"],
    "status": ["status"],
    "start_date": ["start_date", "start", "startdate", "baseline_start", "baselinestart"],
    "finish_date": ["finish_date", "finish", "finishdate", "end", "end_date", "baseline_finish", "baselinefinish"],
    "duration": ["duration", "dur"],
    "percent_complete": ["percent_complete", "percent", "%_complete", "complete", "perc"],
    "schedule_health": ["schedule_health", "health", "schedule_status"],
    "total_float": ["total_float", "float", "totalfloat"],
    "critical": ["critical", "is_critical", "critical_"],
    "hierarchy_level": ["hierarchy_level", "level", "outline_level", "ancestors_level"],
    "variance": ["variance", "var"],
    "predecessors": ["predecessors", "predecessor", "precs"],
    "successors": ["successors", "successor", "succs"],
    "ancestors": ["ancestors", "ancestor"],
    # phase / milestone
    "phase_name": ["phase_name", "phase", "milestone_group"],
    "milestone_name": ["milestone_name", "milestone"],
    "due_date": ["due_date", "due", "milestone_due_date"],
}


def _build_normalized_column_map(columns: Iterable[Any]) -> Dict[str, str]:
    """Map from canonical normalized key -> actual normalized column name in sheet."""
    normalized = {_normalize_colname(c): c for c in columns}
    inv: Dict[str, str] = {}

    for canonical, syns in _COL_SYNONYMS.items():
        for syn in syns:
            syn_n = _normalize_colname(syn)
            if syn_n in normalized:
                inv[canonical] = syn_n
                break

    return inv


def _find_sheet(df_dict: Dict[str, pd.DataFrame], kind: str) -> Optional[str]:
    """Heuristic sheet selection without hardcoding names."""
    if not df_dict:
        return None

    kind = kind.lower()
    keywords = {
        "tasks": ["task", "schedule", "activities", "gantt"],
        "summary": ["summary", "overview", "stats", "report"],
        "comments": ["comment", "comments", "notes", "discussion"],
        "milestones": ["milestone"],
    }.get(kind, [kind])

    # Use order of preference: exact-ish keyword hit in name
    for name in df_dict.keys():
        nl = name.lower()
        if any(k in nl for k in keywords):
            return name

    # Fallback to column signature
    if kind == "comments":
        for name, df in df_dict.items():
            cols = {_normalize_colname(c) for c in df.columns}
            if {"comment"}.intersection(cols) or any("comment" in c for c in cols):
                return name

    if kind == "summary":
        for name, df in df_dict.items():
            cols = {_normalize_colname(c) for c in df.columns}
            if {"status"}.intersection(cols) and {"start_date", "finish_date"}.intersection(cols):
                return name

    if kind == "tasks":
        # look for outline/hierarchy columns
        for name, df in df_dict.items():
            cols = {_normalize_colname(c) for c in df.columns}
            if any(c in cols for c in ["level", "hierarchy_level", "outline_level", "ancestors"]) and any(
                c in cols for c in ["start_date", "finish_date", "percent_complete", "task_name", "task"]
            ):
                return name

    # Last fallback: first sheet
    return next(iter(df_dict.keys()))


def _split_list_cell(value: Any) -> List[str]:
    if value is None:
        return []
    if isinstance(value, list):
        return [str(x).strip() for x in value if str(x).strip()]
    s = str(value).strip()
    if not s or s.upper() in {"#UNPARSEABLE", "#N/A", "N/A"}:
        return []

    # Common patterns: "1FS, 2SS", "1, 2", etc.
    parts = re.split(r"[,;\n]+", s)
    out = []
    for p in parts:
        p = p.strip()
        if not p:
            continue
        # Remove duplicate whitespace
        out.append(re.sub(r"\s+", " ", p))
    return out


def _safe_text(v: Any) -> Optional[str]:
    if v is None:
        return None
    if isinstance(v, float) and pd.isna(v):
        return None
    s = str(v)
    s = s.strip()
    if not s:
        return None
    if s.upper() in {"#UNPARSEABLE", "#N/A", "N/A"}:
        return None
    return s


def _clean_percent(value: Any) -> Optional[float]:
    f = _coerce_float(value)
    if f is None:
        return None
    # sometimes percent_complete is 0..1
    if 0 <= f <= 1:
        return f * 100.0
    return f


# -----------------------------
# Core parser
# -----------------------------


def parse_project_xlsx(path: str | Path) -> Project:
    """Parse a Microsoft Project export workbook into a normalized Project dataclass.

    Never crashes on malformed cells; best-effort extraction.
    """
    path = Path(path)

    try:
        workbook_sheets = pd.read_excel(path, sheet_name=None, engine="openpyxl")
        xls: Dict[str, pd.DataFrame] = {}
        for sheet_name, df in workbook_sheets.items():
            if df is None:
                xls[sheet_name] = pd.DataFrame()
                continue
            sheet_lower = sheet_name.lower()
            if sheet_lower in {"summary", "comments"}:
                xls[sheet_name] = pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl", header=None)
            else:
                xls[sheet_name] = df.copy()
    except Exception as e:
        logger.exception("Failed to open workbook %s", path)
        # Return empty project with metadata unknown
        return Project(metadata=ProjectMetadata())

    # Pick candidate sheets
    tasks_sheet_name = _find_sheet(xls, "tasks")
    comments_sheet_name = _find_sheet(xls, "comments")
    summary_sheet_name = _find_sheet(xls, "summary")

    tasks_df = xls.get(tasks_sheet_name, pd.DataFrame())
    comments_df = xls.get(comments_sheet_name, pd.DataFrame()) if comments_sheet_name else pd.DataFrame()
    summary_df = xls.get(summary_sheet_name, pd.DataFrame()) if summary_sheet_name else pd.DataFrame()

    # Normalize columns for each df
    task_cols_map = _build_normalized_column_map(tasks_df.columns) if not tasks_df.empty else {}
    comments_cols_map = _build_normalized_column_map(comments_df.columns) if not comments_df.empty else {}
    summary_cols_map = _build_normalized_column_map(summary_df.columns) if not summary_df.empty else {}

    # Extract project metadata by scanning for key/value patterns in all sheets/columns.
    metadata = _parse_metadata_from_workbook(xls)

    tasks: List[TaskInfo] = _parse_tasks(tasks_df, task_cols_map)
    phases: List[PhaseInfo] = _parse_phases(tasks_df, task_cols_map)
    milestones: List[MilestoneInfo] = _parse_milestones(tasks_df, task_cols_map)
    comments: List[TaskComment] = _parse_comments(comments_df, comments_cols_map)

    summary = SummaryStatistics(
        num_tasks=len(tasks),
        num_milestones=len(milestones),
        num_comments=len(comments),
        num_phases=len(phases),
    )

    return Project(metadata=metadata, tasks=tasks, phases=phases, milestones=milestones, comments=comments, summary=summary)


def _parse_metadata_from_workbook(all_sheets: Dict[str, pd.DataFrame]) -> ProjectMetadata:
    found: Dict[str, Any] = {}

    for sheet_name, df in all_sheets.items():
        if df is None or df.empty:
            continue

        sheet_lower = sheet_name.lower()
        if sheet_lower == "summary":
            try:
                for row_idx in range(min(len(df), 30)):
                    row = df.iloc[row_idx]
                    if len(row) == 0:
                        continue
                    key = _safe_text(row.iloc[0]) if len(row) > 0 else None
                    if not key:
                        continue
                    key_l = key.lower()
                    value = _safe_text(row.iloc[1]) if len(row) > 1 else None
                    if key_l == "project name" and value and "project_name" not in found:
                        found["project_name"] = value
                    elif key_l == "project manager" and value and "project_manager" not in found:
                        found["project_manager"] = value
                    elif key_l == "project start date" and "report_date" not in found:
                        found["report_date"] = _coerce_date(row.iloc[1]) if len(row) > 1 else None
            except Exception:
                continue

    project_name = _safe_text(found.get("project_name"))
    client_name = _safe_text(found.get("client_name"))
    project_manager = _safe_text(found.get("project_manager"))

    if not project_name or not client_name:
        for sheet_name in all_sheets.keys():
            if sheet_name is None:
                continue
            normalized_sheet = sheet_name.strip().lower()
            if normalized_sheet in {"summary", "comments"}:
                continue
            if "-" in sheet_name:
                parts = [p.strip() for p in re.split(r"\s*-\s*", sheet_name, maxsplit=1)]
                if len(parts) == 2:
                    if not client_name and parts[0]:
                        client_name = parts[0]
                    if not project_name and parts[1]:
                        project_name = parts[1]
                    if project_name and client_name:
                        break
            elif not project_name:
                project_name = sheet_name.strip()
                break

    if not project_name:
        project_name = "Unknown Project"

    return ProjectMetadata(
        project_name=project_name,
        client_name=client_name,
        project_manager=project_manager,
        report_date=_coerce_date(found.get("report_date")),
    )



def _parse_tasks(tasks_df: pd.DataFrame, col_map: Dict[str, str]) -> List[TaskInfo]:
    if tasks_df is None or tasks_df.empty:
        return []

    # Create normalized columns view
    norm_to_actual = {_normalize_colname(c): c for c in tasks_df.columns}

    def get_col(canonical: str) -> Optional[pd.Series]:
        norm = col_map.get(canonical)
        if not norm:
            return None
        actual = norm_to_actual.get(norm)
        if not actual:
            return None
        return tasks_df[actual]

    # Identify hierarchy level / ancestors
    level_ser = get_col("hierarchy_level")
    if level_ser is None:
        level_ser = get_col("ancestors")

    task_name_ser = get_col("task_name")
    uid_ser = get_col("task_uid")

    if task_name_ser is None and uid_ser is None:
        # Try common fallback columns
        for candidate in ["task", "name", "resource_name", "outline_level"]:
            norm = _normalize_colname(candidate)
            if norm in norm_to_actual:
                task_name_ser = tasks_df[norm_to_actual[norm]]
                break

    # For hierarchical levels, handle strings like "1.2.3" or "Level 2"
    def coerce_level(v: Any) -> Optional[int]:
        if v is None:
            return None
        if isinstance(v, (int, float)) and not pd.isna(v):
            return int(v)
        s = _safe_text(v)
        if not s:
            return None
        s_u = s.upper()
        m = re.search(r"LEVEL\s*(\d+)", s_u)
        if m:
            try:
                return int(m.group(1))
            except Exception:
                return None
        # Ancestors "1,2,3" or "1.2.3"
        if re.match(r"^\d+(?:[\.,]\d+)*$", s.strip()):
            try:
                parts = re.split(r"[\.,]", s.strip())
                return len([p for p in parts if p.strip()])
            except Exception:
                return None
        # Outline level as leading number
        m = re.match(r"^(\d+)\b", s.strip())
        return int(m.group(1)) if m else None

    start_ser = get_col("start_date")
    finish_ser = get_col("finish_date")
    duration_ser = get_col("duration")
    pc_ser = get_col("percent_complete")
    status_ser = get_col("status")
    health_ser = get_col("schedule_health")
    float_ser = get_col("total_float")
    critical_ser = get_col("critical")
    variance_ser = get_col("variance")
    pred_ser = get_col("predecessors")
    succ_ser = get_col("successors")

    tasks: List[TaskInfo] = []

    # Determine iteration rows: use dataframe index
    n = len(tasks_df)
    for i in range(n):
        task_name = _safe_text(task_name_ser.iloc[i]) if task_name_ser is not None else None
        uid = _safe_text(uid_ser.iloc[i]) if uid_ser is not None else None

        # Skip empty rows
        if not task_name and not uid:
            continue

        t = TaskInfo(
            task_uid=uid,
            task_name=task_name,
            status=_safe_text(status_ser.iloc[i]) if status_ser is not None else None,
            start_date=_coerce_date(start_ser.iloc[i]) if start_ser is not None else None,
            finish_date=_coerce_date(finish_ser.iloc[i]) if finish_ser is not None else None,
            duration=_coerce_float(duration_ser.iloc[i]) if duration_ser is not None else None,
            percent_complete=_clean_percent(pc_ser.iloc[i]) if pc_ser is not None else None,
            schedule_health=_safe_text(health_ser.iloc[i]) if health_ser is not None else None,
            total_float=_coerce_float(float_ser.iloc[i]) if float_ser is not None else None,
            critical=_coerce_bool(critical_ser.iloc[i]) if critical_ser is not None else None,
            hierarchy_level=coerce_level(level_ser.iloc[i]) if level_ser is not None else None,
            variance=_coerce_float(variance_ser.iloc[i]) if variance_ser is not None else None,
            predecessors=_split_list_cell(pred_ser.iloc[i]) if pred_ser is not None else [],
            successors=_split_list_cell(succ_ser.iloc[i]) if succ_ser is not None else [],
        )
        tasks.append(t)

    return tasks


def _parse_phases(tasks_df: pd.DataFrame, col_map: Dict[str, str]) -> List[PhaseInfo]:
    # Best-effort: phases are usually summary rows (Level with long names) or have status like "Phase".
    if tasks_df is None or tasks_df.empty:
        return []

    norm_to_actual = {_normalize_colname(c): c for c in tasks_df.columns}

    def get_col(canonical: str) -> Optional[pd.Series]:
        norm = col_map.get(canonical)
        if not norm:
            return None
        actual = norm_to_actual.get(norm)
        if not actual:
            return None
        return tasks_df[actual]


    name_ser = get_col("task_name")
    level_ser = get_col("hierarchy_level")
    if level_ser is None:
        level_ser = get_col("ancestors")
    status_ser = get_col("status")


    phases: List[PhaseInfo] = []

    for i in range(len(tasks_df)):
        nm = _safe_text(name_ser.iloc[i]) if name_ser is not None else None
        if not nm:
            continue
        st = _safe_text(status_ser.iloc[i]) if status_ser is not None else None

        # Heuristic: phase rows often have status containing "phase" or are top-level (level==1)
        lvl = None
        if level_ser is not None:
            lvl = _coerce_float(level_ser.iloc[i])
            lvl = int(lvl) if lvl is not None else None

        if st and re.search(r"phase", st, re.I):
            phases.append(PhaseInfo(phase_name=nm, hierarchy_level=lvl))
        elif lvl == 1:
            # include top-level summaries
            phases.append(PhaseInfo(phase_name=nm, hierarchy_level=lvl))

    # Deduplicate by name+level
    uniq: Dict[Tuple[Optional[str], Optional[int]], PhaseInfo] = {}
    for p in phases:
        uniq[(p.phase_name, p.hierarchy_level)] = p
    return list(uniq.values())


def _parse_milestones(tasks_df: pd.DataFrame, col_map: Dict[str, str]) -> List[MilestoneInfo]:
    if tasks_df is None or tasks_df.empty:
        return []

    norm_to_actual = {_normalize_colname(c): c for c in tasks_df.columns}

    def get_col(canonical: str) -> Optional[pd.Series]:
        norm = col_map.get(canonical)
        if not norm:
            return None
        actual = norm_to_actual.get(norm)
        return tasks_df[actual] if actual else None

    name_ser = get_col("task_name")
    due_ser = get_col("finish_date")
    if due_ser is None:
        due_ser = get_col("start_date")
    level_ser = get_col("hierarchy_level")
    if level_ser is None:
        level_ser = get_col("ancestors")
    duration_ser = get_col("duration")
    start_ser = get_col("start_date")

    milestones: List[MilestoneInfo] = []

    for i in range(len(tasks_df)):
        nm = _safe_text(name_ser.iloc[i]) if name_ser is not None else None
        if not nm:
            continue

        duration = _coerce_float(duration_ser.iloc[i]) if duration_ser is not None else None
        start_date = _coerce_date(start_ser.iloc[i]) if start_ser is not None else None
        finish_date = _coerce_date(due_ser.iloc[i]) if due_ser is not None else None
        is_milestone = (duration is not None and duration <= 0.0) or (
            duration is None and start_date is not None and finish_date is not None and start_date == finish_date
        )
        if not is_milestone:
            continue

        lvl = None
        if level_ser is not None:
            lvl = _coerce_float(level_ser.iloc[i])
            lvl = int(lvl) if lvl is not None else None

        milestones.append(
            MilestoneInfo(
                milestone_name=nm,
                due_date=finish_date or start_date,
                hierarchy_level=lvl,
            )
        )

    uniq: Dict[Tuple[Optional[str], Optional[date], Optional[int]], MilestoneInfo] = {}
    for m in milestones:
        uniq[(m.milestone_name, m.due_date, m.hierarchy_level)] = m
    return list(uniq.values())


def _parse_comments(comments_df: pd.DataFrame, col_map: Dict[str, str]) -> List[TaskComment]:
    if comments_df is None or comments_df.empty:
        return []

    comments: List[TaskComment] = []
    for i in range(len(comments_df)):
        row = comments_df.iloc[i]
        if len(row) == 0:
            continue

        values = [_safe_text(row.iloc[j]) for j in range(min(4, len(row))) if j < len(row)]
        if not any(values):
            continue

        comment = _safe_text(row.iloc[1]) if len(row) > 1 else None
        author = _safe_text(row.iloc[2]) if len(row) > 2 else None
        comment_date = _coerce_date(row.iloc[3]) if len(row) > 3 else None

        if comment:
            comments.append(
                TaskComment(
                    task_uid=None,
                    task_name=None,
                    author=author,
                    comment_date=comment_date,
                    comment=comment,
                )
            )

    return comments

