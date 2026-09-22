"""Post-analysis filters for the Employee Performance view.

The employee is selected before collection, so these filters operate only on
the collected and analysed snapshot.  They never call Jira or ClickUp again.
"""
from __future__ import annotations

from collections.abc import Iterable, Mapping

import pandas as pd


ALL = "All"
COMPANY_UNKNOWN = "Company not specified"
PROJECT_UNKNOWN = "Project / Space not specified"
STATUS_UNKNOWN = "Status unavailable"
TASK_TYPE_UNKNOWN = "Task type unavailable"
PRIORITY_UNKNOWN = "Priority not specified"
DUE_STATE_UNKNOWN = "Due-date state unavailable"

FILTER_COLUMNS = (
    "Company",
    "Project / Space",
    "Status",
    "Task Type",
    "Priority",
    "Due-Date State",
)


def _text(value: object, fallback: str) -> str:
    if value is None:
        return fallback
    try:
        if pd.isna(value):
            return fallback
    except (TypeError, ValueError):
        pass
    text = str(value).strip()
    return text if text and text.casefold() not in {"nan", "none", "null"} else fallback


def _series(frame: pd.DataFrame, column: str, fallback: str) -> pd.Series:
    if column not in frame.columns:
        return pd.Series(fallback, index=frame.index, dtype="object")
    return frame[column].map(lambda value: _text(value, fallback))


def _jira_company(frame: pd.DataFrame) -> pd.Series:
    candidates = [
        column for column in frame.columns
        if any(token in str(column).casefold() for token in ("company", "client", "customer", "organization"))
    ]
    if not candidates:
        return pd.Series(COMPANY_UNKNOWN, index=frame.index, dtype="object")
    values = frame[candidates].copy()
    output = pd.Series(COMPANY_UNKNOWN, index=frame.index, dtype="object")
    for column in candidates:
        candidate = values[column].map(lambda value: _text(value, ""))
        output = output.mask(output.eq(COMPANY_UNKNOWN) & candidate.ne(""), candidate)
    return output


def _due_state_jira(frame: pd.DataFrame) -> pd.Series:
    def classify(row):
        if bool(row.get("is_rejected", False)):
            return "Cancelled / rejected"
        if bool(row.get("is_completed", False)):
            variance = row.get("schedule_variance_days")
            if pd.notna(variance):
                return "Completed late" if float(variance) > 0 else "Completed on time"
            return DUE_STATE_UNKNOWN
        if bool(row.get("is_open", False)):
            overdue = row.get("overdue_days")
            if pd.notna(overdue):
                return "Open overdue" if float(overdue) > 0 else "Open not due"
            return "Open without due date"
        return DUE_STATE_UNKNOWN

    return frame.apply(classify, axis=1)


def with_employee_filter_dimensions(frame: pd.DataFrame, source: str) -> pd.DataFrame:
    """Return a copy with common filter dimensions for Jira or ClickUp."""
    result = frame.copy()
    normalized_source = str(source or "").casefold()

    if normalized_source == "jira":
        result["Company"] = _jira_company(result)
        project = _series(result, "project_name", "")
        project_key = _series(result, "project_key", "")
        result["Project / Space"] = project.where(project.ne(""), project_key).replace("", PROJECT_UNKNOWN)
        result["Status"] = _series(result, "status_at_cutoff", STATUS_UNKNOWN)
        result["Task Type"] = _series(result, "issue_type", TASK_TYPE_UNKNOWN)
        result["Priority"] = _series(result, "priority", PRIORITY_UNKNOWN)
        result["Due-Date State"] = _due_state_jira(result)
    else:
        result["Company"] = _series(result, "Company", COMPANY_UNKNOWN)
        result["Project / Space"] = _series(result, "Project / Space", PROJECT_UNKNOWN)
        result["Status"] = _series(result, "Current Status", STATUS_UNKNOWN)
        result["Task Type"] = _series(result, "Task Type", TASK_TYPE_UNKNOWN)
        result["Priority"] = _series(result, "Priority", PRIORITY_UNKNOWN)
        result["Due-Date State"] = _series(result, "Due-Date State", DUE_STATE_UNKNOWN)

    return result


def _selected_values(values: Iterable[object] | None) -> set[str]:
    return {
        str(value).strip()
        for value in (values or [])
        if str(value).strip() and str(value).strip() != ALL
    }


def apply_employee_filters(frame: pd.DataFrame, selections: Mapping[str, Iterable[object] | object] | None) -> pd.DataFrame:
    """Apply selected post-analysis values without mutating the source frame."""
    result = frame.copy()
    for column in FILTER_COLUMNS:
        selected = selections.get(column) if selections else None
        if isinstance(selected, str):
            selected = [selected]
        selected_set = _selected_values(selected)
        if not selected_set or column not in result.columns:
            continue
        result = result[result[column].map(str).isin(selected_set)]
    result.attrs = dict(frame.attrs)
    return result.reset_index(drop=True)


def employee_filter_options(frame: pd.DataFrame, company: str = ALL) -> dict[str, list[str]]:
    """Return sorted options; Projects / Spaces depend on the Company choice."""
    options: dict[str, list[str]] = {}
    company_values = sorted(frame["Company"].dropna().astype(str).unique(), key=str.casefold) if "Company" in frame else []
    options["Company"] = [ALL, *company_values]

    scoped = frame
    if company and company != ALL and "Company" in scoped:
        scoped = scoped[scoped["Company"].astype(str).eq(company)]
    project_values = sorted(scoped["Project / Space"].dropna().astype(str).unique(), key=str.casefold) if "Project / Space" in scoped else []
    options["Project / Space"] = [ALL, *project_values]
    for column in FILTER_COLUMNS[2:]:
        values = sorted(frame[column].dropna().astype(str).unique(), key=str.casefold) if column in frame else []
        options[column] = values
    return options
