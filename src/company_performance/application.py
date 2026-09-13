"""Application-facing orchestration for Company Performance.

The source collectors remain the owners of authentication and API access.
This module only reads their already-prepared payloads, turns them into the
isolated company contract, and reconstructs a selected calendar period.
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import date
from io import BytesIO
from typing import Any, Mapping

from openpyxl import load_workbook

from .adapters import (
    CompanyCollection,
    SourceCoverage,
    adapt_clickup_prepared,
    adapt_jira_collection,
    combine_company_sources,
)
from .dashboard import CompanyDashboardModel, DashboardFilters, build_company_dashboard
from .kpis import BottleneckCandidate, Recommendation, generate_recommendations, identify_bottleneck_candidates
from .normalization import calendar_date, deduplicate_tasks
from .workflow import reconstruct_task
from .models import TaskPeriodSnapshot


@dataclass(frozen=True)
class CompanyAnalysisResult:
    """One complete, source-independent company analysis run."""

    collection: CompanyCollection
    snapshots: tuple[TaskPeriodSnapshot, ...]
    model: CompanyDashboardModel
    bottlenecks: tuple[BottleneckCandidate, ...]
    recommendations: tuple[Recommendation, ...]


def _value(row: Mapping[str, Any], name: str) -> Any:
    value = row.get(name)
    return None if value == "" else value


def _jira_export_issues(prepared_data: Any) -> list[dict[str, Any]]:
    """Restore the minimal Jira issue payload from the source XLSX.

    Jira's existing ``PreparedData`` deliberately persists a history JSON but
    not a second copy of every API issue object.  Its source workbook is the
    audited collector output, so rebuilding the adapter input from ``Jira_Data``
    keeps Company Performance additive and avoids a new API request.
    """
    try:
        workbook = load_workbook(BytesIO(prepared_data.xlsx), read_only=True, data_only=True)
        if "Jira_Data" not in workbook.sheetnames:
            return []
        sheet = workbook["Jira_Data"]
        rows = sheet.iter_rows(values_only=True)
        headers = [str(value) if value is not None else "" for value in next(rows)]
    except (AttributeError, KeyError, OSError, StopIteration, TypeError, ValueError):
        return []

    output: list[dict[str, Any]] = []
    for values in rows:
        row = dict(zip(headers, values))
        key = _value(row, "Issue key") or _value(row, "Issue id")
        if not key:
            continue
        assignee = _value(row, "Assignee")
        output.append({
            "key": str(key),
            "id": str(_value(row, "Issue id") or ""),
            "fields": {
                "summary": _value(row, "Summary"),
                "project": {"key": _value(row, "Project key"), "name": _value(row, "Project name")},
                "status": {"name": _value(row, "Status")},
                "priority": {"name": _value(row, "Priority")},
                "assignee": {"displayName": assignee} if assignee else None,
                "created": _value(row, "Created"),
                "duedate": _value(row, "Due date"),
                "start_date": _value(row, "Custom field (Start date)"),
            },
        })
    return output


def _jira_histories(prepared_data: Any) -> dict[str, Mapping[str, Any]]:
    try:
        raw = prepared_data.history_json
        payload = json.loads(raw.decode("utf-8") if isinstance(raw, bytes) else raw)
    except (AttributeError, TypeError, UnicodeDecodeError, json.JSONDecodeError):
        return {}
    return payload if isinstance(payload, dict) else {}


def _adapt_jira_prepared(prepared_data: Any, unified_project: str):
    return adapt_jira_collection(
        _jira_export_issues(prepared_data),
        _jira_histories(prepared_data),
        unified_project=unified_project,
        source_space=getattr(prepared_data, "space_name", None),
        collection_timestamp=getattr(prepared_data, "collected_at", None),
    )


def build_company_analysis(
    *,
    period_start: date,
    period_end: date,
    jira_prepared: Any | None = None,
    jira_project: str | None = None,
    clickup_prepared: Any | None = None,
    clickup_project: str | None = None,
    filters: DashboardFilters | None = None,
) -> CompanyAnalysisResult:
    """Build a company run from one or both existing prepared source results.

    The caller must explicitly provide a Unified Project name for every
    included source.  This reflects the approved manual mapping and avoids
    guessing that two similarly named Jira/ClickUp spaces are the same work.
    """
    if period_end < period_start:
        raise ValueError("Analysis period end must not be before its start.")
    sources = []
    if jira_prepared is not None:
        if not (jira_project or "").strip():
            raise ValueError("Enter a Unified Project name for the selected Jira space.")
        sources.append(_adapt_jira_prepared(jira_prepared, jira_project.strip()))
    if clickup_prepared is not None:
        if not (clickup_project or "").strip():
            raise ValueError("Enter a Unified Project name for the selected ClickUp space.")
        sources.append(adapt_clickup_prepared(clickup_prepared, unified_project=clickup_project.strip()))
    if not sources:
        raise ValueError("Collect at least one Jira or ClickUp space before running Company Performance.")

    collection = combine_company_sources(*sources)
    records = deduplicate_tasks(collection.records)
    snapshots = tuple(
        reconstruct_task(
            record,
            period_start,
            period_end,
            collection_date=calendar_date(record.collection_timestamp),
        )
        for record in records
    )
    if not snapshots:
        raise ValueError("The selected sources contain no usable task records.")
    model = build_company_dashboard(snapshots, coverages=collection.coverages, filters=filters)
    return CompanyAnalysisResult(
        collection=collection,
        snapshots=snapshots,
        model=model,
        bottlenecks=identify_bottleneck_candidates(snapshots),
        recommendations=generate_recommendations(snapshots),
    )
