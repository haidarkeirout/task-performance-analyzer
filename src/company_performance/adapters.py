"""Adapters from existing Jira and ClickUp collection payloads.

The Company Performance domain is intentionally additive: these helpers read
the payloads already produced by the source collectors and create neutral
``TaskRecord`` objects.  They do not call either API, mutate source payloads,
or change the existing source-specific reports.
"""

from __future__ import annotations

import copy
import json
from dataclasses import dataclass, field
from datetime import date, datetime, time, timezone
from types import MappingProxyType
from typing import Any, Iterable, Mapping

from .models import ParentClassification, StatusTransition, TaskRecord
from .normalization import DAMASCUS, calendar_date, normalize_assignees


@dataclass(frozen=True)
class SourceCoverage:
    """Explicit collection coverage for one mapped source.

    ``history_mode`` is deliberately separate from source availability: a
    ClickUp task snapshot can be usable for basic KPIs even though it has no
    chronological status history.
    """

    source_tool: str
    source_space: str | None
    unified_project: str | None
    source_available: bool
    task_count: int
    history_mode: str
    reason: str | None = None
    flags: tuple[str, ...] = ()


@dataclass
class AdaptedSource:
    """Neutral records plus read-only copies of their source payloads."""

    records: tuple[TaskRecord, ...]
    coverage: SourceCoverage
    raw_tasks: Mapping[tuple[str, str], Mapping[str, Any]] = field(default_factory=dict)
    raw_workflow: Mapping[tuple[str, str], Mapping[str, Any]] = field(default_factory=dict)
    raw_time_in_status: Mapping[tuple[str, str], Mapping[str, Any]] = field(default_factory=dict)


@dataclass(frozen=True)
class CompanyCollection:
    """The two isolated source adapters presented as one company collection."""

    records: tuple[TaskRecord, ...]
    coverages: tuple[SourceCoverage, ...]
    sources: tuple[AdaptedSource, ...]


def _freeze_mapping(value: Any) -> Mapping[str, Any]:
    if isinstance(value, Mapping):
        return MappingProxyType(copy.deepcopy(dict(value)))
    return MappingProxyType({"value": copy.deepcopy(value)})


def _text(value: Any) -> str | None:
    if value is None:
        return None
    value = str(value).strip()
    return value or None


def _named(value: Any) -> str | None:
    if isinstance(value, Mapping):
        for key in ("displayName", "name", "username", "email", "value", "id"):
            result = _text(value.get(key))
            if result:
                return result
        return None
    return _text(value)


def _clickup_status(value: Any) -> tuple[str | None, bool]:
    """Return a readable ClickUp status, using the internal id only as fallback."""
    if isinstance(value, Mapping):
        for key in ("status", "name", "status_name", "label"):
            result = _text(value.get(key))
            if result:
                return result, False
        fallback = _text(value.get("id"))
        return fallback, bool(fallback)
    return _text(value), False


def _clickup_priority(value: Any) -> str | None:
    """Read ClickUp's priority label before falling back to its identifier."""
    if isinstance(value, Mapping):
        for key in ("priority", "name", "label", "value", "id"):
            result = _text(value.get(key))
            if result:
                return result
        return None
    return _text(value)


def _collection_time(value: Any) -> datetime | None:
    if isinstance(value, datetime):
        return value if value.tzinfo else value.replace(tzinfo=timezone.utc)
    if isinstance(value, date):
        return datetime.combine(value, time.min, tzinfo=timezone.utc)
    if value in (None, ""):
        return None
    try:
        numeric = float(value)
        if abs(numeric) > 100_000_000_000:
            return datetime.fromtimestamp(numeric / 1000.0, tz=timezone.utc)
    except (TypeError, ValueError, OSError):
        pass
    if isinstance(value, str):
        try:
            parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
        except ValueError:
            return None
        return parsed if parsed.tzinfo else parsed.replace(tzinfo=timezone.utc)
    return None


def _source_date(value: Any) -> date | None:
    """Handle ISO Jira dates and ClickUp millisecond dates in Damascus time."""
    stamp = _collection_time(value)
    if stamp is not None:
        return stamp.astimezone(DAMASCUS).date()
    return calendar_date(value)


def _parent(value: Any) -> str | None:
    if isinstance(value, Mapping):
        for key in ("key", "id", "task_id"):
            result = _text(value.get(key))
            if result:
                return result
        return None
    if isinstance(value, str):
        try:
            parsed = json.loads(value)
        except json.JSONDecodeError:
            parsed = None
        if isinstance(parsed, Mapping):
            return _parent(parsed)
    return _text(value)


def _issue_type(value: Any) -> str | None:
    """Return Jira's work-item type without interpreting its business meaning."""
    return _named(value)


def _jira_type_key(value: str | None) -> str:
    return " ".join(str(value or "").strip().casefold().replace("_", " ").split())


def _jira_is_epic(issue_type: str | None) -> bool:
    return _jira_type_key(issue_type) in {"epic", "portfolio epic", "initiative"}


def _jira_is_subtask(issue_type: str | None) -> bool:
    return _jira_type_key(issue_type) in {"sub-task", "subtask", "sub task"}


def _parent_classification(parent_id: str | None) -> ParentClassification:
    return ParentClassification.SUBTASK if parent_id else ParentClassification.STANDALONE


def _classify_parent_relationships(records: list[TaskRecord]) -> None:
    """Mark parent tasks from the parent references returned by each source.

    A task with a parent reference is a Subtask.  A task referenced by one or
    more returned subtasks is an operational Parent Task.  Parent detection is
    source-local and uses task IDs, so Jira keys and ClickUp IDs remain intact.
    """
    referenced_parent_ids = {record.parent_id for record in records if record.parent_id}
    for record in records:
        # A Jira Epic is a container even when work items reference it as their
        # parent.  It must never be promoted to an operational parent task.
        if record.parent_classification is ParentClassification.CONTAINER:
            continue
        if record.parent_id:
            # Jira's ``parent`` field is also used for an Epic relationship in
            # newer Cloud APIs.  Only an explicitly typed Sub-task is a
            # completion/KPI child; Task/Story/Bug remain operational items.
            if not (
                record.source_tool.strip().casefold() == "jira"
                and record.issue_type
                and not _jira_is_subtask(record.issue_type)
            ):
                record.parent_classification = ParentClassification.SUBTASK
        elif record.task_id in referenced_parent_ids:
            record.parent_classification = ParentClassification.INDEPENDENT
        else:
            record.parent_classification = ParentClassification.STANDALONE


def _history_transitions(history: Mapping[str, Any] | None) -> tuple[StatusTransition, ...]:
    transitions: list[StatusTransition] = []
    for event in (history or {}).get("status_events") or ():
        if not isinstance(event, Mapping):
            continue
        changed_at = _collection_time(event.get("changed_at"))
        if changed_at is None:
            continue
        transitions.append(
            StatusTransition(
                changed_at=changed_at,
                from_status=_text(event.get("from_status")),
                to_status=_text(event.get("to_status")),
                performed_by=_text(event.get("author_name") or event.get("performed_by")),
            )
        )
    return tuple(sorted(transitions, key=lambda event: event.changed_at))


def _issue_assignees(value: Any) -> tuple[str, ...]:
    if isinstance(value, list):
        return normalize_assignees(_named(item) for item in value)
    return normalize_assignees((_named(value),))


def _clickup_assignees(values: Any) -> tuple[str, ...]:
    if not isinstance(values, list):
        return ()
    return normalize_assignees(_named(value) for value in values)


def _raw_mapping(values: Iterable[tuple[tuple[str, str], Any]]) -> Mapping[tuple[str, str], Mapping[str, Any]]:
    return MappingProxyType({key: _freeze_mapping(value) for key, value in values})


def adapt_jira_collection(
    issues: Iterable[Mapping[str, Any]] | None,
    histories: Mapping[str, Mapping[str, Any]] | None,
    *,
    unified_project: str | None,
    source_space: str | None = None,
    collection_timestamp: Any = None,
    planned_start_field: str | None = None,
    failure_reason: str | None = None,
) -> AdaptedSource:
    """Adapt Jira collector output without changing its collection contract.

    ``issues`` and ``histories`` are the values already held by the Jira
    collection job.  Jira's complete changelog becomes chronological neutral
    transitions; invalid/missing source history remains a visible task flag.
    """
    values = list(issues or ())
    histories = histories or {}
    collected_at = _collection_time(collection_timestamp)
    if failure_reason:
        return AdaptedSource(
            records=(),
            coverage=SourceCoverage(
                "Jira", source_space, unified_project, False, 0, "Unavailable",
                str(failure_reason), ("Jira Collection Failed",),
            ),
        )

    records: list[TaskRecord] = []
    raw_tasks: list[tuple[tuple[str, str], Any]] = []
    raw_workflow: list[tuple[tuple[str, str], Any]] = []
    complete_count = 0
    for issue in values:
        fields = issue.get("fields") if isinstance(issue.get("fields"), Mapping) else {}
        task_id = _text(issue.get("key") or issue.get("id"))
        if not task_id:
            continue
        project = fields.get("project") if isinstance(fields.get("project"), Mapping) else {}
        actual_space = source_space or _named(project.get("name")) or _named(project.get("key"))
        history = histories.get(task_id) or histories.get(str(issue.get("id", "")))
        if not isinstance(history, Mapping):
            history = {}
        workflow = _history_transitions(history)
        history_complete = history.get("history_complete") is True
        history_through = _collection_time(history.get("history_through"))
        if history_complete:
            complete_count += 1
        status = fields.get("status")
        priority = fields.get("priority")
        issue_type = _issue_type(fields.get("issuetype") or fields.get("issue_type"))
        parent_id = _parent(fields.get("parent")) or _parent(
            fields.get("epic_link") or fields.get("epicLink")
        )
        start_value = None
        if planned_start_field:
            start_value = fields.get(planned_start_field)
        if start_value is None:
            for key in ("start_date", "startDate", "planned_start_date", "customfield_start_date"):
                if fields.get(key) is not None:
                    start_value = fields.get(key)
                    break
        flags: set[str] = set()
        if not history_complete:
            flags.add("Missing Workflow History")
        if history and history.get("history_complete") is True and not history.get("history_through"):
            flags.add("Missing Workflow Coverage Timestamp")
        if history and len(workflow) < len(history.get("status_events") or ()):
            flags.add("Invalid Workflow Event")
        if _jira_is_epic(issue_type):
            parent_classification = ParentClassification.CONTAINER
        elif _jira_is_subtask(issue_type):
            parent_classification = _parent_classification(parent_id)
        else:
            # A standard Jira issue may have an Epic in ``fields.parent``;
            # that relationship is hierarchy metadata, not a Sub-task.
            parent_classification = ParentClassification.STANDALONE
        record = TaskRecord(
            source_tool="Jira",
            task_id=task_id,
            task_name=_text(fields.get("summary")),
            issue_type=issue_type,
            source_space=actual_space,
            unified_project=unified_project,
            department="Tech Development",
            raw_status=_named(status),
            initial_status=_text(history.get("initial_status")) or _named(status),
            priority=_named(priority),
            assignees=_issue_assignees(fields.get("assignee")),
            created_date=_source_date(fields.get("created")),
            planned_start_date=_source_date(start_value),
            due_date=_source_date(fields.get("duedate")),
            parent_id=parent_id,
            parent_classification=parent_classification,
            collection_timestamp=collected_at,
            history_through=history_through,
            workflow_history=workflow,
            history_complete=history_complete,
            data_quality_flags=flags,
        )
        records.append(record)
        raw_tasks.append((record.unique_key, issue))
        raw_workflow.append((record.unique_key, history))

    _classify_parent_relationships(records)

    flags: list[str] = []
    if not values:
        flags.append("Empty Source Result")
    if values and complete_count != len(records):
        flags.append("Partial Workflow History")
    mode = "Complete" if records and complete_count == len(records) else ("Partial" if complete_count else "Unavailable")
    return AdaptedSource(
        records=tuple(records),
        coverage=SourceCoverage(
            "Jira", source_space, unified_project, True, len(records), mode,
            None, tuple(flags),
        ),
        raw_tasks=_raw_mapping(raw_tasks),
        raw_workflow=_raw_mapping(raw_workflow),
    )


def adapt_clickup_collection(
    tasks: Iterable[Mapping[str, Any]] | None,
    time_in_status: Mapping[str, Mapping[str, Any]] | None = None,
    *,
    unified_project: str | None,
    source_space: str | None,
    collection_timestamp: Any = None,
    failure_reason: str | None = None,
) -> AdaptedSource:
    """Adapt ClickUp task snapshots and native Total time in Status payloads.

    ClickUp's current integration does not collect chronological transitions.
    The adapter therefore never manufactures them: records remain explicitly
    history-unavailable, while native status-duration payloads stay attached to
    the adapter result for later status-duration reporting.
    """
    values = list(tasks or ())
    status_payloads = time_in_status or {}
    collected_at = _collection_time(collection_timestamp)
    if failure_reason:
        return AdaptedSource(
            records=(),
            coverage=SourceCoverage(
                "ClickUp", source_space, unified_project, False, 0, "Unavailable",
                str(failure_reason), ("ClickUp Collection Failed",),
            ),
        )

    records: list[TaskRecord] = []
    raw_tasks: list[tuple[tuple[str, str], Any]] = []
    raw_time: list[tuple[tuple[str, str], Any]] = []
    timed_count = 0
    for task in values:
        task_id = _text(task.get("id"))
        if not task_id:
            continue
        raw_status, status_unresolved = _clickup_status(task.get("status"))
        priority = _clickup_priority(task.get("priority"))
        list_value = task.get("list") if isinstance(task.get("list"), Mapping) else {}
        department = _named(list_value.get("name"))
        parent_id = _parent(task.get("parent"))
        flags = {"ClickUp Chronological History Unavailable", "Missing Workflow History"}
        if status_unresolved:
            flags.add("ClickUp Status Unresolved")
        if not department:
            flags.add("Missing ClickUp List Name")
        payload = status_payloads.get(task_id)
        if isinstance(payload, Mapping):
            timed_count += 1
        else:
            flags.add("ClickUp Total Time in Status Unavailable")
        record = TaskRecord(
            source_tool="ClickUp",
            task_id=task_id,
            task_name=_text(task.get("name")),
            source_space=source_space,
            unified_project=unified_project,
            department=department,
            raw_status=raw_status,
            initial_status=None,
            priority=priority,
            assignees=_clickup_assignees(task.get("assignees")),
            created_date=_source_date(task.get("date_created")),
            planned_start_date=_source_date(task.get("start_date")),
            due_date=_source_date(task.get("due_date")),
            parent_id=parent_id,
            parent_classification=_parent_classification(parent_id),
            collection_timestamp=collected_at,
            history_through=collected_at,
            workflow_history=(),
            history_complete=False,
            data_quality_flags=flags,
        )
        records.append(record)
        raw_tasks.append((record.unique_key, task))
        if isinstance(payload, Mapping):
            raw_time.append((record.unique_key, payload))

    _classify_parent_relationships(records)

    flags: list[str] = []
    if not values:
        flags.append("Empty Source Result")
    if records and timed_count != len(records):
        flags.append("Partial Total Time in Status Coverage")
    if records and not timed_count:
        flags.append("Total Time in Status Unavailable")
    return AdaptedSource(
        records=tuple(records),
        coverage=SourceCoverage(
            "ClickUp", source_space, unified_project, True, len(records), "Unavailable",
            None, tuple(flags),
        ),
        raw_tasks=_raw_mapping(raw_tasks),
        raw_time_in_status=_raw_mapping(raw_time),
    )


def adapt_clickup_prepared(
    prepared_data: Any,
    *,
    unified_project: str | None,
    source_space: str | None = None,
) -> AdaptedSource:
    """Read the existing ``ClickUpPreparedData.history_json`` payload safely."""
    try:
        raw = getattr(prepared_data, "history_json")
        payload = json.loads(raw.decode("utf-8") if isinstance(raw, bytes) else raw)
    except (AttributeError, TypeError, UnicodeDecodeError, json.JSONDecodeError):
        return adapt_clickup_collection(
            (), unified_project=unified_project, source_space=source_space,
            failure_reason="ClickUp prepared payload is unreadable.",
        )
    return adapt_clickup_collection(
        payload.get("tasks") if isinstance(payload, Mapping) else (),
        payload.get("time_in_status") if isinstance(payload, Mapping) else None,
        unified_project=unified_project,
        source_space=source_space or _text(getattr(prepared_data, "space_name", None)),
        collection_timestamp=getattr(prepared_data, "collected_at", None),
    )


def combine_company_sources(*sources: AdaptedSource) -> CompanyCollection:
    """Combine adapters without cross-source de-duplication or source mutation."""
    return CompanyCollection(
        records=tuple(record for source in sources for record in source.records),
        coverages=tuple(source.coverage for source in sources),
        sources=tuple(sources),
    )
