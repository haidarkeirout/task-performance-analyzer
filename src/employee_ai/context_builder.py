"""Build Employee AI Assistant V1 context from already-calculated results."""

from __future__ import annotations

from dataclasses import asdict, is_dataclass
from datetime import datetime, timezone
import math
import re
from typing import Any, Iterable, Mapping
from uuid import uuid4

from .fingerprint import context_fingerprint, stable_subject_hash
from .sanitizer import (
    allowlisted_record,
    assert_no_sensitive_fields,
    transport_value,
)
from .schemas import EmployeeAIContext, SCHEMA_VERSION, ContextValidationError


_ALLOWED_ACTIONS = ["explain", "summarize", "compare_visible_metrics", "cite_evidence"]
_FORBIDDEN_ACTIONS = ["collect", "run_analysis", "recalculate", "update_task", "delete_task"]


def _records(value: Any) -> list[dict[str, Any]]:
    if value is None:
        return []
    if isinstance(value, Mapping):
        return [dict(value)]
    if hasattr(value, "to_dict"):
        try:
            return [dict(item) for item in value.to_dict(orient="records")]
        except TypeError:
            converted = value.to_dict()
            if isinstance(converted, Mapping):
                return [dict(converted)]
    output = []
    for item in value:
        if isinstance(item, Mapping):
            output.append(dict(item))
        elif is_dataclass(item):
            output.append(asdict(item))
        elif hasattr(item, "__dict__"):
            output.append(dict(vars(item)))
    return output


def _object_mapping(value: Any) -> dict[str, Any]:
    if value is None:
        return {}
    if isinstance(value, Mapping):
        return dict(value)
    if is_dataclass(value):
        return asdict(value)
    return dict(vars(value)) if hasattr(value, "__dict__") else {}


def _numeric(value: Any) -> int | float | None:
    value = transport_value(value)
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return value
    text = str(value).strip().replace(",", "")
    if text.endswith("%"):
        text = text[:-1]
    try:
        number = float(text)
    except ValueError:
        return None
    if math.isfinite(number) and number.is_integer():
        return int(number)
    return number if math.isfinite(number) else None


def _metric_key(value: Any) -> str:
    """Normalise renderer-specific card keys to the public contract style."""

    text = re.sub(r"[^a-z0-9]+", "_", str(value or "metric").casefold())
    return text.strip("_") or "metric"


def _card(
    key: str,
    label: str,
    display_value: Any,
    *,
    numeric_value: Any = None,
    unit: str = "count",
    components: Mapping[str, Any] | None = None,
    reason_unavailable: str | None = None,
) -> dict[str, Any]:
    numeric = _numeric(display_value if numeric_value is None else numeric_value)
    available = numeric is not None
    return {
        "key": key,
        "label": label,
        "display_value": "N/A" if display_value is None else transport_value(str(display_value)),
        "numeric_value": numeric,
        "unit": unit,
        "components": transport_value(dict(components or {})),
        "availability": "available" if available else "unavailable",
        "reason_unavailable": None if available else (reason_unavailable or "Not provided by the analysis result"),
    }


def _scope_filters(filters: Mapping[str, Any] | None) -> dict[str, Any]:
    filters = filters or {}
    return {
        "company": transport_value(filters.get("company", "All")),
        "projects_spaces": transport_value(filters.get("projects_spaces", [])),
        "statuses": transport_value(filters.get("statuses", [])),
        "task_types": transport_value(filters.get("task_types", [])),
        "priorities": transport_value(filters.get("priorities", [])),
        "due_date_states": transport_value(filters.get("due_date_states", [])),
    }


def build_employee_ai_context(
    *,
    source_mode: str,
    employee: Mapping[str, Any],
    collection: Mapping[str, Any],
    cards: Iterable[Mapping[str, Any]],
    tasks: Iterable[Mapping[str, Any]] = (),
    analysis_filters: Mapping[str, Any] | None = None,
    collection_spaces: Iterable[str] = (),
    period: Mapping[str, Any] | None = None,
    summaries: Mapping[str, Any] | None = None,
    data_quality: Mapping[str, Any] | None = None,
    question: str = "",
    locale: str = "en",
    view_state: Mapping[str, Any] | None = None,
    subject_secret: str,
    max_evidence_rows: int = 25,
    request_id: str | None = None,
    sent_at: str | None = None,
) -> EmployeeAIContext:
    """Build a safe context. Metrics must already be calculated by the system."""

    mode = str(source_mode).strip().casefold()
    if mode not in {"jira", "clickup", "combined"}:
        raise ContextValidationError(f"Unsupported Employee source mode: {source_mode}")
    if max_evidence_rows < 0:
        raise ContextValidationError("max_evidence_rows cannot be negative")

    display_name = str(employee.get("display_name") or "").strip()
    if not display_name:
        raise ContextValidationError("Employee display name is required")
    subject_key = str(employee.get("subject_key") or display_name)
    employee_public = {
        "display_name": transport_value(display_name),
        "department": transport_value(employee.get("department")),
        "position": transport_value(employee.get("position")),
        "stable_subject_hash": stable_subject_hash(subject_key, subject_secret),
    }

    filters = _scope_filters(analysis_filters)
    scope = {
        "collection_spaces": transport_value(list(collection_spaces)),
        "analysis_filters": filters,
        "view_state": {
            "active_tab": transport_value((view_state or {}).get("active_tab")),
            "local_filters": {},
        },
    }
    period_value = {
        "mode": transport_value((period or {}).get("mode", "full_snapshot")),
        "start": transport_value((period or {}).get("start")),
        "end": transport_value((period or {}).get("end")),
        "timezone": transport_value((period or {}).get("timezone", "UTC")),
    }
    collection_public = {
        "collection_id": transport_value(collection.get("collection_id")),
        "collected_at": transport_value(collection.get("collected_at")),
        "analysis_completed_at": transport_value(collection.get("analysis_completed_at")),
    }
    card_values = [transport_value(dict(item)) for item in cards]
    all_tasks = [transport_value(dict(item)) for item in tasks]
    selected_tasks = all_tasks[:max_evidence_rows]
    quality = {
        "overall_state": transport_value((data_quality or {}).get("overall_state", "good")),
        "history_complete": transport_value((data_quality or {}).get("history_complete")),
        "source_coverage": transport_value((data_quality or {}).get("source_coverage", [])),
        "warnings": transport_value((data_quality or {}).get("warnings", [])),
        "unavailable_metrics": transport_value((data_quality or {}).get("unavailable_metrics", [])),
    }

    fingerprint_identity = {
        "schema_version": SCHEMA_VERSION,
        "analysis_type": "employee",
        "source_mode": mode,
        "employee_hash": employee_public["stable_subject_hash"],
        "collection": collection_public,
        "period": period_value,
        "scope": {"collection_spaces": scope["collection_spaces"], "analysis_filters": filters},
        "cards": card_values,
        "task_count_in_scope": len(all_tasks),
        "evidence_fingerprint": context_fingerprint(all_tasks),
    }
    fingerprint = context_fingerprint(fingerprint_identity)
    context = EmployeeAIContext(
        schema_version=SCHEMA_VERSION,
        request={
            "request_id": request_id or str(uuid4()),
            "sent_at": sent_at or datetime.now(timezone.utc).isoformat(),
            "question": transport_value(question),
            "locale": transport_value(locale),
            "context_fingerprint": fingerprint,
        },
        analysis={
            "analysis_type": "employee",
            "source_mode": mode,
            "employee": employee_public,
            "collection": collection_public,
            "period": period_value,
            "scope": scope,
        },
        dashboard={
            "cards": card_values,
            "summaries": transport_value(dict(summaries or {})),
        },
        evidence={
            "mode": "inline" if selected_tasks else "none",
            "task_count_in_scope": len(all_tasks),
            "rows_included": len(selected_tasks),
            "truncated": len(selected_tasks) < len(all_tasks),
            "selection_reason": "current_filtered_employee_scope",
            "tasks": selected_tasks,
        },
        permissions={
            "mode": "read_only",
            "allowed_actions": list(_ALLOWED_ACTIONS),
            "forbidden_actions": list(_FORBIDDEN_ACTIONS),
            "source_access": "none",
        },
        data_quality=quality,
    )
    payload = context.to_dict()
    assert_no_sensitive_fields(payload)
    return context


def _jira_tasks(task_metrics: Any) -> list[dict[str, Any]]:
    mapping = {
        "source": "source_tool",
        "task_key": "issue_key",
        "task_name": "task_name",
        "company": "Company",
        "project_space": "Project / Space",
        "status": "status_at_cutoff",
        "priority": "priority",
        "due_date": "due_date",
        "due_state": "Due-Date State",
        "overdue_days": "overdue_days",
        "counted_in_kpis": "status_known",
        "exclusion_reason": "exclusion_reason",
        "quality_flags": "data_quality_flags",
    }
    rows = []
    for source in _records(task_metrics):
        item = allowlisted_record(source, mapping)
        item.setdefault("source", "Jira")
        rows.append(item)
    return rows


def build_jira_employee_context(
    *,
    employee: Mapping[str, Any],
    collection: Mapping[str, Any],
    dashboard_values: Mapping[str, Any],
    task_metrics: Any,
    analysis_filters: Mapping[str, Any] | None = None,
    collection_spaces: Iterable[str] = (),
    validation_warnings: Iterable[Any] = (),
    question: str = "",
    subject_secret: str,
    max_evidence_rows: int = 25,
) -> EmployeeAIContext:
    """Adapt Jira's already-rendered Employee metrics without recalculating them."""

    values = dict(dashboard_values)
    known = _numeric(values.get("known_status_tasks"))
    unknown = _numeric(values.get("unknown_status_tasks"))
    cards = [
        _card("total_tasks", "Total Tasks", values.get("total")),
        _card("completion_population", "Tasks Included in Completion Rate", known),
        _card(
            "completion_rate", "Completion Rate", values.get("completion_rate"), unit="percent",
            components={"numerator": values.get("completed"), "denominator": known, "excluded": unknown},
        ),
        _card("completed_tasks", "Completed", values.get("completed")),
        _card(
            "on_time_rate", "On-Time Rate", values.get("on_time_rate"), unit="percent",
            components={"numerator": values.get("on_time_tasks"), "denominator": values.get("on_time_valid_tasks")},
        ),
        _card("open_overdue", "Open Overdue", values.get("overdue")),
        _card("wip_tasks", "WIP Tasks", values.get("wip")),
    ]
    rows = _records(task_metrics)
    history_values = [row.get("history_complete") for row in rows if "history_complete" in row]
    history_complete = all(bool(value) for value in history_values) if history_values else None
    cutoff = next((row.get("evaluation_cutoff") for row in rows if row.get("evaluation_cutoff")), None)
    warnings = list(validation_warnings)
    return build_employee_ai_context(
        source_mode="jira",
        employee=employee,
        collection=collection,
        cards=cards,
        tasks=_jira_tasks(task_metrics),
        analysis_filters=analysis_filters,
        collection_spaces=collection_spaces,
        period={"mode": "full_snapshot", "end": cutoff, "timezone": "Asia/Damascus"},
        data_quality={
            "overall_state": "limited" if warnings or history_complete is False else "good",
            "history_complete": history_complete,
            "warnings": warnings,
        },
        question=question,
        subject_secret=subject_secret,
        max_evidence_rows=max_evidence_rows,
    )


def _metric_map(overall: Any) -> dict[str, Any]:
    output: dict[str, Any] = {}
    for row in _records(overall):
        name = row.get("Metric")
        if name is not None:
            output[str(name)] = row.get("Value")
    return output


def _clickup_tasks(tasks: Any) -> list[dict[str, Any]]:
    mapping = {
        "source": "Source",
        "task_key": "Task ID",
        "task_name": "Task Name",
        "company": "Company",
        "project_space": "Project / Space",
        "status": "Current Status",
        "priority": "Priority",
        "due_date": "Due Date",
        "due_state": "Due-Date State",
        "overdue_days": "Overdue Days",
        "counted_in_kpis": "Status Known?",
        "exclusion_reason": "Exclusion Reason",
        "quality_flags": "Timing Data Status",
    }
    rows = []
    for source in _records(tasks):
        item = allowlisted_record(source, mapping)
        item.setdefault("source", "ClickUp")
        rows.append(item)
    return rows


def build_clickup_employee_context(
    *,
    employee: Mapping[str, Any],
    collection: Mapping[str, Any],
    analysis_result: Mapping[str, Any],
    analysis_filters: Mapping[str, Any] | None = None,
    collection_spaces: Iterable[str] = (),
    question: str = "",
    subject_secret: str,
    max_evidence_rows: int = 25,
) -> EmployeeAIContext:
    """Adapt the already-calculated ClickUp Employee analysis result."""

    metrics = _metric_map(analysis_result.get("overall"))
    known = metrics.get("Known status tasks")
    unknown = metrics.get("Unknown status tasks")
    cards = [
        _card("total_tasks", "Total Tasks", metrics.get("Total tasks")),
        _card("completion_population", "Tasks Included in Completion Rate", known),
        _card(
            "completion_rate", "Completion Rate", metrics.get("Completion rate (%)"), unit="percent",
            components={"numerator": metrics.get("Completed tasks"), "denominator": known, "excluded": unknown},
        ),
        _card("completed_tasks", "Completed", metrics.get("Completed tasks")),
        _card("on_time_rate", "On-Time Rate", metrics.get("On-time completion rate (%)"), unit="percent"),
        _card("open_overdue", "Open Overdue", metrics.get("Open overdue tasks")),
        _card("wip_tasks", "WIP Tasks", metrics.get("WIP tasks")),
    ]
    tasks = analysis_result.get("tasks")
    task_rows = _records(tasks)
    known_values = [row.get("Status Known?") for row in task_rows if "Status Known?" in row]
    history_complete = all(bool(value) for value in known_values) if known_values else None
    quality_rows = _records(analysis_result.get("quality"))
    warnings = [row for row in quality_rows if str(row.get("Status", "")).casefold() not in {"", "ok"}]
    return build_employee_ai_context(
        source_mode="clickup",
        employee=employee,
        collection=collection,
        cards=cards,
        tasks=_clickup_tasks(tasks),
        analysis_filters=analysis_filters,
        collection_spaces=collection_spaces,
        period={
            "mode": "full_snapshot",
            "start": analysis_result.get("period_start") or None,
            "end": analysis_result.get("period_end") or analysis_result.get("cutoff") or None,
            "timezone": analysis_result.get("source_timezone") or "Asia/Damascus",
        },
        summaries={
            "status_distribution": _records(analysis_result.get("status_counts")),
            "due_state_distribution": _records(analysis_result.get("due_status_summary")),
            "weekly_flow": _records(analysis_result.get("weekly_flow")),
        },
        data_quality={
            "overall_state": "limited" if warnings or history_complete is False else "good",
            "history_complete": history_complete,
            "warnings": warnings,
        },
        question=question,
        subject_secret=subject_secret,
        max_evidence_rows=max_evidence_rows,
    )


def _combined_tasks(task_details: Any) -> list[dict[str, Any]]:
    mapping = {
        "source": "source_tool",
        "task_key": "task_id",
        "task_name": "task_name",
        "company": "company_name",
        "project_space": "source_space",
        "status": "final_status",
        "priority": "priority",
        "due_date": "due_date",
        "counted_in_kpis": "counted_in_kpis",
        "exclusion_reason": "exclusion_reason",
        "quality_flags": "data_quality_flags",
    }
    return [allowlisted_record(source, mapping) for source in _records(task_details)]


def build_combined_employee_context(
    *,
    employee: Mapping[str, Any],
    collection: Mapping[str, Any],
    analysis_result: Any,
    analysis_filters: Mapping[str, Any] | None = None,
    collection_spaces: Iterable[str] = (),
    question: str = "",
    subject_secret: str,
    max_evidence_rows: int = 25,
) -> EmployeeAIContext:
    """Adapt the source-neutral combined Jira + ClickUp Employee model."""

    result = _object_mapping(analysis_result)
    model = result.get("model") if isinstance(result.get("model"), Mapping) else _object_mapping(result.get("model"))
    if not model:
        raise ContextValidationError("Combined Employee analysis model is required")
    kpis = _object_mapping(model.get("kpis"))
    cards = []
    for source in _records(model.get("cards")):
        key = _metric_key(source.get("key"))
        components = {}
        if key == "completion_rate":
            components = {
                "numerator": kpis.get("completed_tasks"),
                "denominator": kpis.get("known_status_tasks"),
                "excluded": kpis.get("unknown_status_tasks"),
            }
        cards.append(_card(
            key,
            str(source.get("title") or source.get("key") or "Metric"),
            source.get("value"),
            unit="percent" if "rate" in key else "count",
            components=components,
        ))
    if not cards:
        cards = [
            _card("total_tasks", "Total Tasks", kpis.get("total_tasks")),
            _card(
                "completion_rate", "Completion Rate", kpis.get("completion_rate"), unit="percent",
                components={
                    "numerator": kpis.get("completed_tasks"),
                    "denominator": kpis.get("known_status_tasks"),
                    "excluded": kpis.get("unknown_status_tasks"),
                },
            ),
            _card("completed_tasks", "Completed Tasks", kpis.get("completed_tasks")),
            _card("on_time_rate", "On-Time Rate", kpis.get("on_time_completion_rate"), unit="percent"),
            _card("open_overdue", "Open Overdue", kpis.get("overdue_open_tasks")),
            _card("wip_tasks", "WIP Tasks", kpis.get("current_wip")),
        ]
    if kpis and not any(item["key"] == "completion_population" for item in cards):
        population = _card(
            "completion_population",
            "Tasks Included in Completion Rate",
            kpis.get("known_status_tasks"),
            components={"excluded": kpis.get("unknown_status_tasks")},
        )
        total_index = next(
            (index for index, item in enumerate(cards) if item["key"] == "total_tasks"),
            -1,
        )
        cards.insert(total_index + 1, population)
    coverage = _records(model.get("source_coverage"))
    quality = _records(model.get("data_quality"))
    history_limited = any("history" in str(row).casefold() for row in quality)
    period_start = model.get("period_start")
    period_end = model.get("period_end")
    return build_employee_ai_context(
        source_mode="combined",
        employee=employee,
        collection=collection,
        cards=cards,
        tasks=_combined_tasks(model.get("task_details")),
        analysis_filters=analysis_filters,
        collection_spaces=collection_spaces,
        period={"mode": "full_snapshot", "start": period_start, "end": period_end, "timezone": "Asia/Damascus"},
        summaries={
            "workflow_findings": _records(result.get("bottlenecks")),
        },
        data_quality={
            "overall_state": "limited" if quality else "good",
            "history_complete": False if history_limited else None,
            "source_coverage": coverage,
            "warnings": quality,
        },
        question=question,
        subject_secret=subject_secret,
        max_evidence_rows=max_evidence_rows,
    )
