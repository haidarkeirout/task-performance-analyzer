"""Approved source mappings and non-destructive task normalisation."""

from __future__ import annotations

from dataclasses import fields, replace
from datetime import date, datetime, timezone
from typing import Any, Iterable
from zoneinfo import ZoneInfo

from .models import AssigneeGroup, TaskRecord, UnifiedStatus


DAMASCUS = ZoneInfo("Asia/Damascus")


def _label(value: Any) -> str:
    return "" if value is None else " ".join(str(value).strip().casefold().split())


_STATUS_MAPPINGS: dict[str, dict[str, UnifiedStatus]] = {
    "clickup": {
        "to do": UnifiedStatus.NOT_STARTED,
        "planning": UnifiedStatus.NOT_STARTED,
        "in progress": UnifiedStatus.IN_EXECUTION,
        "at risk": UnifiedStatus.AT_RISK,
        "review": UnifiedStatus.IN_REVIEW,
        "on hold": UnifiedStatus.ON_HOLD,
        "complete": UnifiedStatus.COMPLETED,
        "cancelled": UnifiedStatus.CANCELLED,
    },
    "jira": {
        "idea": UnifiedStatus.NOT_STARTED,
        "in triage": UnifiedStatus.NOT_STARTED,
        "to do": UnifiedStatus.NOT_STARTED,
        "in progress": UnifiedStatus.IN_EXECUTION,
        "in review": UnifiedStatus.IN_REVIEW,
        "done": UnifiedStatus.COMPLETED,
        "rejected": UnifiedStatus.REJECTED,
    },
}


def normalize_source_tool(value: Any) -> str:
    label = _label(value)
    if label == "jira":
        return "Jira"
    if label == "clickup":
        return "ClickUp"
    return str(value).strip() if value is not None else "Unknown"


def normalize_status(source_tool: Any, raw_status: Any) -> UnifiedStatus:
    """Map only approved labels; any other label stays explicitly Unknown."""
    return _STATUS_MAPPINGS.get(_label(source_tool), {}).get(
        _label(raw_status), UnifiedStatus.UNKNOWN
    )


def normalize_priority(source_tool: Any, raw_priority: Any) -> str:
    """Return Critical, High, Medium, Low, or Unknown using the specification."""
    label = _label(raw_priority)
    source = _label(source_tool)
    if not label:
        return "Unknown"

    if source == "jira":
        mapping = {
            "highest": "Critical",
            "high": "High",
            "medium": "Medium",
            "normal": "Medium",
            "low": "Low",
        }
    elif source == "clickup":
        mapping = {
            "1": "Critical",
            "urgent": "Critical",
            "2": "High",
            "high": "High",
            "3": "Medium",
            "normal": "Medium",
            "4": "Low",
            "low": "Low",
        }
    else:
        mapping = {}
    return mapping.get(label, "Unknown")


def calendar_date(value: Any) -> date | None:
    """Read a date or timestamp as a Damascus calendar date without guessing."""
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        stamp = value
    elif isinstance(value, date):
        return value
    elif isinstance(value, str):
        try:
            stamp = datetime.fromisoformat(value.replace("Z", "+00:00"))
        except ValueError:
            return None
    else:
        return None
    if stamp.tzinfo is None:
        stamp = stamp.replace(tzinfo=DAMASCUS)
    return stamp.astimezone(DAMASCUS).date()


def normalize_assignees(values: Iterable[Any] | None) -> tuple[str, ...]:
    """Keep a task once and preserve its unique people for later review."""
    result: list[str] = []
    for value in values or ():
        text = str(value).strip() if value is not None else ""
        if text and text not in result:
            result.append(text)
    return tuple(result)


def assignee_group(task: TaskRecord) -> str:
    if not task.assignees:
        return AssigneeGroup.UNASSIGNED.value
    if len(task.assignees) > 1:
        return AssigneeGroup.MULTIPLE.value
    return task.assignees[0]


def apply_task_quality_flags(task: TaskRecord) -> TaskRecord:
    """Add non-fatal flags that must remain visible in Data Quality."""
    flags: set[str] = set()
    if not task.assignees:
        flags.add("Missing Assignee")
    elif len(task.assignees) > 1:
        flags.add("Multiple Assignees")
    if normalize_status(task.source_tool, task.raw_status) is UnifiedStatus.UNKNOWN:
        flags.add("Unmapped Status")
    if normalize_priority(task.source_tool, task.priority) == "Unknown":
        flags.add("Unmapped Priority")
    task.with_flags(flags)
    return task


def _blank(value: Any) -> bool:
    return value is None or value == "" or value == ()


def _latest_record(records: list[TaskRecord]) -> TaskRecord:
    minimum = datetime.min.replace(tzinfo=timezone.utc)

    def key(record: TaskRecord) -> datetime:
        stamp = record.collection_timestamp
        if stamp is None:
            return minimum
        return stamp if stamp.tzinfo else stamp.replace(tzinfo=timezone.utc)

    return max(records, key=key)


def deduplicate_tasks(records: Iterable[TaskRecord]) -> list[TaskRecord]:
    """Keep latest duplicate, backfill empty values, and preserve conflicts.

    Status history is not merged between conflicting source records: a merged
    synthetic timeline would be less trustworthy than the latest verified one.
    """
    groups: dict[tuple[str, str], list[TaskRecord]] = {}
    for record in records:
        groups.setdefault(record.unique_key, []).append(record)

    output: list[TaskRecord] = []
    protected = {"collection_timestamp", "workflow_history", "data_quality_flags"}
    for group in groups.values():
        selected = _latest_record(group)
        if len(group) == 1:
            output.append(apply_task_quality_flags(selected))
            continue

        merged = replace(selected, data_quality_flags=set(selected.data_quality_flags))
        flags = {"Duplicate Task Record"}
        for field in fields(TaskRecord):
            name = field.name
            if name in protected:
                continue
            selected_value = getattr(merged, name)
            other_values = [getattr(item, name) for item in group if item is not selected]
            if _blank(selected_value):
                for value in other_values:
                    if not _blank(value):
                        setattr(merged, name, value)
                        selected_value = value
                        break
            if not _blank(selected_value):
                for value in other_values:
                    if not _blank(value) and value != selected_value:
                        flags.add("Duplicate Conflicting Records")
                        break
        merged.with_flags(flags)
        output.append(apply_task_quality_flags(merged))
    return output
