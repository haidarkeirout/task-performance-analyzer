"""ClickUp-only client-side task filters.

The ClickUp Tasks API returns a snapshot of the selected Space.  These helpers
apply the same useful filter families shown by ClickUp Views to that snapshot
without reusing Jira query or parsing logic.
"""
from __future__ import annotations

import json
import re
from datetime import date
from typing import Any

import pandas as pd


TERMINAL_STATUSES = {"complete", "completed", "done", "closed"}
CANCELLED_STATUSES = {"cancelled", "canceled"}

# These labels deliberately mirror ClickUp's View/Dashboard filter wording.
MORE_FILTER_LABELS = [
    "Start date",
    "Date created",
    "Date updated",
    "Date closed",
    "Tags",
    "Location/List",
    "Task type",
    "Time estimates",
    "Time tracked",
    "Status is closed",
    "Recurring",
    "Milestone",
    "Dependencies",
    "Created by",
    "Archived",
    "Total time in Status",
    "Custom Fields",
]


def timestamp(value: Any):
    """Return a UTC timestamp for ClickUp's milliseconds or ISO values."""
    if value in (None, "", 0, "0"):
        return pd.NaT
    try:
        numeric = float(value)
        if numeric > 100_000_000_000:
            return pd.to_datetime(numeric, unit="ms", utc=True, errors="coerce")
    except (TypeError, ValueError):
        pass
    return pd.to_datetime(value, utc=True, errors="coerce")


def _string(value: Any, fallback: str = "") -> str:
    if value is None:
        return fallback
    if isinstance(value, dict):
        for key in ("name", "status", "priority", "username", "email", "label", "value", "id"):
            if value.get(key) not in (None, ""):
                return str(value[key])
        return fallback
    text = str(value).strip()
    return text if text and text.lower() not in {"none", "nan", "null"} else fallback


def status_name(task: dict) -> str:
    return _string(task.get("status"), "Unavailable")


def status_type(task: dict) -> str:
    status = task.get("status")
    return _string(status.get("type") if isinstance(status, dict) else "").casefold()


def is_closed(task: dict) -> bool:
    status = status_name(task).casefold()
    return status in TERMINAL_STATUSES or status_type(task) in {"done", "closed"}


def is_cancelled(task: dict) -> bool:
    return status_name(task).casefold() in CANCELLED_STATUSES


def assignees(task: dict) -> list[str]:
    values = []
    for person in task.get("assignees") or []:
        value = _string(person)
        if value:
            values.append(value)
    return values or ["Unassigned"]


def tags(task: dict) -> list[str]:
    values = []
    for tag in task.get("tags") or []:
        value = _string(tag)
        if value:
            values.append(value)
    return values


def location_values(task: dict) -> list[str]:
    values = []
    for label, key in (("Space", "space"), ("Folder", "folder"), ("List", "list")):
        value = _string(task.get(key))
        if value:
            values.append(f"{label}: {value}")
    return values


def task_type(task: dict) -> str:
    return _string(task.get("task_type") or task.get("type"), "Task")


def created_by(task: dict) -> str:
    return _string(task.get("creator"), "Unavailable")


def milliseconds_to_hours(value: Any):
    try:
        return float(value) / 3_600_000.0
    except (TypeError, ValueError):
        return None


def _custom_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (list, tuple)):
        return ", ".join(_custom_value(item) for item in value if _custom_value(item))
    if isinstance(value, dict):
        for key in ("name", "label", "value", "id"):
            if value.get(key) not in (None, ""):
                return _custom_value(value[key])
        return json.dumps(value, ensure_ascii=False, sort_keys=True)
    return _string(value)


def custom_fields(task: dict) -> dict[str, str]:
    result = {}
    for field in task.get("custom_fields") or []:
        if not isinstance(field, dict):
            continue
        name = _string(field.get("name"))
        if not name:
            continue
        value = _custom_value(field.get("value"))
        if value:
            result[name] = value
    return result


def filter_options(tasks: list[dict]) -> dict[str, Any]:
    """Return the available values for the selected Space snapshot."""
    options = {
        "status": set(),
        "assignee": set(),
        "priority": set(),
        "tag": set(),
        "location": set(),
        "task_type": set(),
        "created_by": set(),
        "time_status": set(),
        "custom_fields": {},
    }
    for task in tasks:
        options["status"].add(status_name(task))
        options["assignee"].update(assignees(task))
        priority = _string(task.get("priority"))
        if priority:
            options["priority"].add(priority)
        options["tag"].update(tags(task))
        options["location"].update(location_values(task))
        options["task_type"].add(task_type(task))
        creator = created_by(task)
        if creator != "Unavailable":
            options["created_by"].add(creator)
        for name, value in custom_fields(task).items():
            options["custom_fields"].setdefault(name, set()).add(value)
    return {
        "status": sorted(options["status"], key=str.casefold),
        "assignee": sorted(options["assignee"], key=str.casefold),
        "priority": sorted(options["priority"], key=str.casefold),
        "tag": sorted(options["tag"], key=str.casefold),
        "location": sorted(options["location"], key=str.casefold),
        "task_type": sorted(options["task_type"], key=str.casefold),
        "created_by": sorted(options["created_by"], key=str.casefold),
        "time_status": [],
        "custom_fields": {
            name: sorted(values, key=str.casefold)
            for name, values in sorted(options["custom_fields"].items(), key=lambda item: item[0].casefold())
        },
    }


def date_matches(value: Any, rule: dict | None) -> bool:
    """Apply a ClickUp-style date rule against a task timestamp."""
    if not rule or rule.get("mode", "Any time") == "Any time":
        return True
    value = timestamp(value)
    mode = rule.get("mode")
    if mode == "Is set":
        return pd.notna(value)
    if mode == "Is not set":
        return pd.isna(value)
    if pd.isna(value):
        return False
    current = value.date()
    start = rule.get("start")
    end = rule.get("end")
    if mode == "On":
        return isinstance(start, date) and current == start
    if mode == "Before":
        return isinstance(start, date) and current < start
    if mode == "After":
        return isinstance(start, date) and current > start
    if mode == "Between":
        return isinstance(start, date) and isinstance(end, date) and start <= current <= end
    return True


def number_matches(value: Any, rule: dict | None) -> bool:
    if not rule or rule.get("mode", "Any value") == "Any value":
        return True
    mode = rule.get("mode")
    try:
        number = float(value)
    except (TypeError, ValueError):
        number = None
    if mode == "Has value":
        return number is not None
    if mode == "No value":
        return number is None
    if number is None:
        return False
    minimum = rule.get("minimum")
    maximum = rule.get("maximum")
    if mode == "At least (hours)":
        return number >= float(minimum or 0)
    if mode == "At most (hours)":
        return number <= float(maximum if maximum is not None else minimum or 0)
    if mode == "Between (hours)":
        return float(minimum or 0) <= number <= float(maximum if maximum is not None else minimum or 0)
    return True


def _bool_matches(value: bool, rule: str | None) -> bool:
    if not rule or rule == "Any":
        return True
    return value if rule == "Is" else not value


def _tag_matches(task_tags: list[str], rule: dict | None) -> bool:
    if not rule or not rule.get("values"):
        return True
    selected = set(rule["values"])
    current = set(task_tags)
    mode = rule.get("mode", "Has any of")
    if mode == "Has all of":
        return selected.issubset(current)
    if mode == "Has none of":
        return not bool(selected & current)
    return bool(selected & current)


def _value_matches(values: list[str], selected: list[str] | tuple[str, ...] | None) -> bool:
    return not selected or bool(set(values) & set(selected))


def _status_minutes(payload: Any) -> dict[str, float]:
    values: dict[str, float] = {}
    if not isinstance(payload, dict):
        return values
    entries = list(payload.get("status_history") or [])
    current = payload.get("current_status")
    if isinstance(current, dict):
        entries.append(current)
    for entry in entries:
        if not isinstance(entry, dict) or not entry.get("status"):
            continue
        total_time = entry.get("total_time") or {}
        try:
            minutes = float(total_time.get("by_minute"))
        except (TypeError, ValueError, AttributeError):
            continue
        status = str(entry["status"])
        values[status] = max(values.get(status, 0.0), minutes)
    return values


def filter_tasks(
    tasks: list[dict],
    criteria: dict[str, Any],
    time_status_map: dict[str, Any] | None = None,
) -> list[dict]:
    """Filter ClickUp task snapshots without a Jira dependency.

    Every populated criterion becomes a predicate.  ``match`` follows the
    ClickUp View convention: all predicates (AND) by default, or any predicate
    when the user chooses OR.
    """
    time_status_map = time_status_map or {}
    result = []
    keyword = str(criteria.get("keyword") or "").strip().casefold()
    for task in tasks:
        predicates = []
        if keyword:
            searchable = " ".join([_string(task.get("name")), _string(task.get("description"))]).casefold()
            predicates.append(keyword in searchable)
        if criteria.get("status"):
            predicates.append(_value_matches([status_name(task)], criteria["status"]))
        if criteria.get("assignee"):
            predicates.append(_value_matches(assignees(task), criteria["assignee"]))
        if criteria.get("priority"):
            predicates.append(_value_matches([_string(task.get("priority"), "No priority")], criteria["priority"]))
        due_rule = criteria.get("due_date") or {}
        if due_rule.get("mode", "Any time") != "Any time":
            predicates.append(date_matches(task.get("due_date"), due_rule))

        extras = criteria.get("extras") or {}
        date_fields = {
            "Start date": "start_date",
            "Date created": "date_created",
            "Date updated": "date_updated",
            "Date closed": "date_closed",
        }
        for label, key in date_fields.items():
            rule = extras.get(label) or {}
            if rule.get("mode", "Any time") != "Any time":
                predicates.append(date_matches(task.get(key), rule))
        if "Tags" in extras:
            predicates.append(_tag_matches(tags(task), extras["Tags"]))
        if "Location/List" in extras:
            predicates.append(_value_matches(location_values(task), extras["Location/List"]))
        if "Task type" in extras:
            predicates.append(_value_matches([task_type(task)], extras["Task type"]))
        if "Created by" in extras:
            predicates.append(_value_matches([created_by(task)], extras["Created by"]))
        if "Time estimates" in extras:
            predicates.append(number_matches(milliseconds_to_hours(task.get("time_estimate")), extras["Time estimates"]))
        if "Time tracked" in extras:
            predicates.append(number_matches(milliseconds_to_hours(task.get("time_spent")), extras["Time tracked"]))
        if "Status is closed" in extras:
            predicates.append(_bool_matches(is_closed(task), extras["Status is closed"]))
        if "Recurring" in extras:
            predicates.append(_bool_matches(bool(task.get("recurring")), extras["Recurring"]))
        if "Milestone" in extras:
            predicates.append(_bool_matches(bool(task.get("is_milestone")), extras["Milestone"]))
        if "Dependencies" in extras:
            dependencies = bool(task.get("dependencies") or task.get("linked_tasks"))
            predicates.append(_bool_matches(dependencies, extras["Dependencies"]))
        if "Archived" in extras:
            predicates.append(_bool_matches(bool(task.get("archived")), extras["Archived"]))
        if "Total time in Status" in extras:
            rule = extras["Total time in Status"]
            minutes = _status_minutes(time_status_map.get(str(task.get("id", ""))))
            status = rule.get("status") if isinstance(rule, dict) else None
            value = (minutes.get(status) / 60.0) if status and minutes.get(status) is not None else None
            predicates.append(number_matches(value, rule))
        if "Custom Fields" in extras:
            for field_name, selected in (extras["Custom Fields"] or {}).items():
                predicates.append(_value_matches([custom_fields(task).get(field_name, "")], selected))

        if not predicates:
            matches = True
        elif criteria.get("match", "All (AND)") == "Any (OR)":
            matches = any(predicates)
        else:
            matches = all(predicates)
        if matches:
            result.append(task)
    return result


def criteria_summary(criteria: dict[str, Any]) -> str:
    """Compact, English-only scope text for the collection context and report."""
    pieces = []
    keyword = str(criteria.get("keyword") or "").strip()
    if keyword:
        pieces.append(f"Search: {keyword}")
    for key, label in (("status", "Status"), ("assignee", "Assignee"), ("priority", "Priority")):
        values = criteria.get(key) or []
        if values:
            pieces.append(f"{label}: {', '.join(map(str, values))}")
    due = criteria.get("due_date") or {}
    if due.get("mode") and due.get("mode") != "Any time":
        pieces.append(f"Due date: {due['mode']}")
    extras = criteria.get("extras") or {}
    for label in extras:
        pieces.append(label)
    if not pieces:
        return "All tasks in the selected ClickUp Space"
    joiner = " OR " if criteria.get("match") == "Any (OR)" else " AND "
    return joiner.join(pieces)


def widget_key(label: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", label.casefold()).strip("_")
