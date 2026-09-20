"""Department-first ClickUp collection.

This module is intentionally isolated from the existing Company, Project, and
Employee collection paths. A ClickUp List is treated as a department, and the
same List name may be present in more than one Space.
"""
from __future__ import annotations

import json
from datetime import datetime, time, timezone
from typing import Any
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

from clickup_export import collect_data
from clickup_filters import assignees, filter_options, filter_tasks, status_name, timestamp
from clickup_gateway import ClickUpCollectionError, ClickUpGateway
from jira_department_collection import render_jira_department_collection


def _gateway(settings, workspace_id: str = "") -> ClickUpGateway:
    return ClickUpGateway(
        settings.clickup_token,
        workspace_id=str(workspace_id or settings.clickup_workspace_id or ""),
    )


def _name(value: Any, fallback: str = "") -> str:
    if isinstance(value, dict):
        value = value.get("name") or value.get("status") or value.get("id")
    text = str(value or "").strip()
    return text or fallback


def _department_key(name: str) -> str:
    return " ".join(str(name or "").split()).casefold()


def _priority_name(task: dict) -> str:
    value = task.get("priority")
    if isinstance(value, dict):
        value = (
            value.get("priority")
            or value.get("name")
            or value.get("label")
            or value.get("id")
        )
    text = str(value or "").strip()
    return text or "Unavailable"


def _catalog_spaces(settings):
    gateway = _gateway(settings)
    try:
        workspaces = gateway.workspaces()
        workspace_id = str(settings.clickup_workspace_id or (workspaces[0]["id"] if workspaces else ""))
        if not workspace_id:
            raise ClickUpCollectionError("No ClickUp Workspace was found for this connection.")
        spaces = gateway.spaces(workspace_id)
        catalog = {}
        for space in spaces:
            space_id = str(space.get("id") or "")
            if not space_id:
                continue
            for current_list in gateway.all_lists_for_space(space_id):
                list_id = str(current_list.get("id") or "")
                list_name = _name(current_list.get("name"))
                if not list_id or not list_name:
                    continue
                key = _department_key(list_name)
                catalog.setdefault(key, {
                    "name": list_name,
                    "lists": [],
                })
                catalog[key]["lists"].append({
                    "list_id": list_id,
                    "list_name": list_name,
                    "space_id": space_id,
                    "space_name": _name(space.get("name"), space_id),
                    "folder_name": _name(current_list.get("folder_name")),
                })
        return workspace_id, catalog
    finally:
        gateway.close()


def _annotate_task(task: dict, source: dict) -> dict:
    item = dict(task)
    item["_department_space_id"] = source["space_id"]
    item["_department_space_name"] = source["space_name"]
    item["_department_list_id"] = source["list_id"]
    item["_department_list_name"] = source["list_name"]
    item["_department_folder_name"] = source.get("folder_name", "")
    item.setdefault("space_id", source["space_id"])
    item.setdefault("space", {"id": source["space_id"], "name": source["space_name"]})
    item.setdefault("list", {"id": source["list_id"], "name": source["list_name"]})
    return item


def _merge_sources(target: dict, source: dict) -> None:
    for field in ("_department_space_names", "_department_list_names"):
        current = list(target.get(field) or [])
        values = source.get(field) or []
        if isinstance(values, str):
            values = [values]
        for value in values:
            if value and value not in current:
                current.append(value)
        target[field] = current
    source_pairs = list(target.get("_department_sources") or [])
    pair = {
        "space_id": source.get("_department_space_id", ""),
        "space_name": source.get("_department_space_name", ""),
        "list_id": source.get("_department_list_id", ""),
        "list_name": source.get("_department_list_name", ""),
    }
    if pair not in source_pairs:
        source_pairs.append(pair)
    target["_department_sources"] = source_pairs


def _collect_department_tasks(settings, sources: list[dict], cache_key: str) -> list[dict]:
    cached_key = st.session_state.get("department_tasks_cache_key")
    if cached_key == cache_key:
        return list(st.session_state.get("department_tasks_cache", []))

    gateway = _gateway(settings)
    checkpoint_registry = dict(
        st.session_state.get("department_list_checkpoints") or {}
    )
    source_checkpoints = checkpoint_registry.setdefault(cache_key, {})
    try:
        collected = {}
        raw_count = 0
        overall = st.progress(0.0, text=f"Collecting department Lists: 0 / {len(sources)}")
        for index, source in enumerate(sources):
            list_id = source["list_id"]
            list_checkpoint = source_checkpoints.setdefault(list_id, {})

            def save_checkpoint(state, current_list_id=list_id):
                source_checkpoints[current_list_id] = state
                checkpoint_registry[cache_key] = source_checkpoints
                st.session_state["department_list_checkpoints"] = checkpoint_registry

            tasks = gateway.all_tasks_for_list(
                list_id,
                progress=lambda message, position=index: overall.progress(
                    position / len(sources) if sources else 1.0,
                    text=str(message),
                ),
                checkpoint=list_checkpoint,
                checkpoint_callback=save_checkpoint,
            )
            for raw_task in tasks:
                task_id = str(raw_task.get("id") or "")
                if task_id:
                    raw_count += 1
                if not task_id:
                    continue
                annotated = _annotate_task(raw_task, source)
                annotated["_department_space_names"] = [source["space_name"]]
                annotated["_department_list_names"] = [source["list_name"]]
                annotated["_department_sources"] = [{
                    "space_id": source["space_id"],
                    "space_name": source["space_name"],
                    "list_id": source["list_id"],
                    "list_name": source["list_name"],
                }]
                if task_id in collected:
                    _merge_sources(collected[task_id], annotated)
                else:
                    collected[task_id] = annotated
            overall.progress(
                (index + 1) / len(sources) if sources else 1.0,
                text=f"Collected department List {index + 1} / {len(sources)}",
            )
        overall.progress(1.0, text=f"Department collection complete: {len(collected)} tasks")
        tasks = list(collected.values())
    finally:
        gateway.close()

    st.session_state["department_tasks_cache_key"] = cache_key
    st.session_state["department_tasks_cache"] = tasks
    st.session_state["department_duplicate_count"] = max(0, raw_count - len(tasks))
    return tasks


def _date_value(value):
    if isinstance(value, tuple):
        return value[0] if value else None
    return value


def _period_bounds(start_date, end_date, source_timezone: str = "Asia/Damascus"):
    """Return local calendar-day boundaries converted to UTC.

    Department dates are business dates, not UTC dates.  Converting after the
    local boundary is built prevents tasks around midnight in Damascus from
    moving into the previous or following day.
    """
    zone = ZoneInfo(source_timezone)
    start = pd.Timestamp(
        datetime.combine(_date_value(start_date), time.min, tzinfo=zone)
    ).tz_convert("UTC")
    end = pd.Timestamp(
        datetime.combine(_date_value(end_date), time.max, tzinfo=zone)
    ).tz_convert("UTC")
    return start, end


def _analysis_cutoff(start_date, end_date, source_timezone: str) -> pd.Timestamp:
    """Return the selected local period end without extending into the future."""
    _, period_end = _period_bounds(start_date, end_date, source_timezone)
    return min(period_end, pd.Timestamp(datetime.now(timezone.utc)))


def _period_filter(
    tasks: list[dict],
    start_date,
    end_date,
    source_timezone: str = "Asia/Damascus",
):
    """Keep tasks whose active life can overlap the selected period.

    A department period is not a creation-date report.  Work created before the
    period remains in scope when it is still open, was completed during/after the
    period, or has insufficient close-date evidence to prove it was already
    closed.  Definitively closed tasks from before the period are excluded.
    """
    start, end = _period_bounds(start_date, end_date, source_timezone)
    filtered = []
    missing_created = 0
    for task in tasks:
        created = timestamp(task.get("date_created"))
        if pd.isna(created):
            missing_created += 1
            continue
        if created > end:
            continue

        closed = timestamp(task.get("date_closed") or task.get("date_done"))
        status = task.get("status") or {}
        status_type = str(status.get("type") or "").strip().casefold() if isinstance(status, dict) else ""
        status_label = status_name(task).strip().casefold()
        terminal = status_type in {"done", "closed"} or status_label in {
            "complete", "completed", "done", "closed", "cancelled", "canceled",
        }
        if terminal and pd.notna(closed) and closed < start:
            continue
        filtered.append(task)
    return filtered, missing_created


def _preview(tasks: list[dict]) -> pd.DataFrame:
    rows = []
    for task in tasks:
        rows.append({
            "Task ID": str(task.get("id", "")),
            "Task Name": task.get("name"),
            "Space": task.get("_department_space_name", ""),
            "Department": task.get("_department_list_name", ""),
            "Assignee": ", ".join(assignees(task)),
            "Priority": _priority_name(task),
            "Status": status_name(task),
            "Created": timestamp(task.get("date_created")),
            "Due Date": timestamp(task.get("due_date")),
            "Parent Task ID": str(task.get("parent") or ""),
            "Subtask": bool(task.get("parent")),
        })
    return pd.DataFrame(rows)


def render_department_collection(settings):
    """Render the approved Department-first ClickUp collection flow."""
    st.subheader("Department data collection")
    st.caption(
        "Choose a department, set the analysis period, and the system will "
        "automatically scan its connected source."
    )

    clickup_error = None
    try:
        workspace_id, catalog = _catalog_spaces(settings)
    except ClickUpCollectionError as exc:
        # Jira Tech remains usable if the ClickUp connector is unavailable.
        workspace_id, catalog = "", {}
        clickup_error = str(exc)

    jira_department_key = "jira-tech-development"
    catalog[jira_department_key] = {
        "name": "Tech (Jira)",
        "source": "Jira",
        "lists": [],
    }
    department_keys = sorted(catalog, key=lambda key: catalog[key]["name"].casefold())
    selected_key = st.selectbox(
        "Choose Department",
        [None, *department_keys],
        format_func=lambda key: "Choose a department" if key is None else catalog[key]["name"],
        key="department_selected_name",
    )
    if selected_key is None:
        return None, False

    selected = catalog[selected_key]
    if selected_key == jira_department_key:
        for key in (
            "clickup_prepared_data", "clickup_run_analysis", "clickup_error",
        ):
            st.session_state.pop(key, None)
        return render_jira_department_collection(settings)
    for key in (
        "jira_department_prepared_data", "jira_department_run_analysis",
    ):
        st.session_state.pop(key, None)
    st.session_state["data_source"] = "ClickUp"
    if clickup_error:
        st.error(clickup_error)
        return None, False
    sources = selected["lists"]
    source_spaces = sorted({item["space_name"] for item in sources}, key=str.casefold)
    source_caption = ", ".join(source_spaces)

    left, right = st.columns(2)
    with left:
        start_date = st.date_input("From date", value=None, key="department_period_start")
    with right:
        end_date = st.date_input("To date", value=None, key="department_period_end")

    start_date = _date_value(start_date)
    end_date = _date_value(end_date)
    if start_date is None or end_date is None:
        st.info("Select both From and To dates to load the department scope.")
        return None, False
    if start_date > end_date:
        st.error("The From date must be on or before the To date.")
        return None, False

    st.caption(
        f"Department: {selected['name']} · Matching Lists: {len(sources)} · "
        f"Spaces: {source_caption}"
    )

    cache_key = json.dumps(
        {"department": selected_key, "sources": sources},
        sort_keys=True,
        default=str,
    )
    try:
        all_tasks = _collect_department_tasks(settings, sources, cache_key)
    except ClickUpCollectionError as exc:
        st.error(str(exc))
        return None, False

    period_tasks, missing_created = _period_filter(
        all_tasks, start_date, end_date, settings.source_timezone
    )
    options = filter_options(period_tasks)
    st.subheader("Department filters")

    filter_left, filter_middle, filter_right = st.columns(3)
    with filter_left:
        keyword = st.text_input("Search tasks", key="department_search")
        selected_statuses = st.multiselect("Status", options["status"], key="department_status")
    with filter_middle:
        selected_assignees = st.multiselect("Assignee", options["assignee"], key="department_assignee")
        selected_priorities = st.multiselect("Priority", options["priority"], key="department_priority")
    with filter_right:
        match = st.selectbox("Match filters", ["All (AND)", "Any (OR)"], key="department_match")

    criteria = {
        "keyword": keyword,
        "status": selected_statuses,
        "assignee": selected_assignees,
        "priority": selected_priorities,
        "due_date": {"mode": "Any time"},
        "match": match,
        "extras": {},
    }
    filtered_tasks = filter_tasks(period_tasks, criteria)
    filter_summary = (
        f"Department: {selected['name']} · Period: {start_date} to {end_date} · "
        f"Source Spaces: {source_caption}"
    )

    st.subheader("Matching tasks")
    st.caption(
        f"{len(filtered_tasks)} of {len(period_tasks)} tasks in period · "
        f"{len(all_tasks)} total tasks collected · {missing_created} missing created date"
    )
    st.dataframe(_preview(filtered_tasks), hide_index=True, use_container_width=True)
    if not filtered_tasks:
        st.warning("No tasks match the selected department, period, and filters.")

    fingerprint = "department:" + json.dumps({
        "department": selected_key,
        "sources": sources,
        "from": str(start_date),
        "to": str(end_date),
        "criteria": criteria,
        "task_ids": [str(task.get("id", "")) for task in filtered_tasks],
    }, sort_keys=True, default=str)

    previous = st.session_state.get("clickup_prepared_data")
    if previous is not None and previous.fingerprint != fingerprint:
        # Keep the cached task preview/source available while invalidating only
        # the result. A later Run Analysis rebuilds the prepared source locally
        # for the newly selected period or filters.
        st.session_state.pop("department_analysis", None)

    if st.button("Done", type="primary", key="department_done", disabled=not filtered_tasks):
        st.session_state.pop("clickup_error", None)
        try:
            end_cutoff_value = _analysis_cutoff(
                start_date, end_date, settings.source_timezone
            )
            end_cutoff = end_cutoff_value.isoformat()
            with st.status("Preparing department tasks for analysis...", expanded=True) as progress:
                prepared = collect_data(
                    None,
                    filtered_tasks,
                    selected["name"],
                    fingerprint,
                    settings.source_timezone,
                    progress=lambda message: progress.update(label=message),
                    space_id="multiple",
                    time_status_data={},
                    time_status_error="",
                    filter_summary=filter_summary,
                    filter_criteria=criteria,
                    analysis_mode="department",
                    department_name=selected["name"],
                    department_id=selected_key,
                    cutoff=end_cutoff,
                )
                prepared.space_names = source_spaces
                prepared.period_start = str(start_date)
                prepared.period_end = str(end_date)
                prepared.duplicate_count = st.session_state.get("department_duplicate_count", 0)
                st.session_state["clickup_prepared_data"] = prepared
                progress.update(label="Department task collection completed.", state="complete", expanded=False)
        except Exception as exc:
            st.session_state["clickup_error"] = str(exc)
            st.error(f"Department collection could not be completed: {exc}")

    prepared = st.session_state.get("clickup_prepared_data")
    prepared_is_current = (
        prepared is not None
        and prepared.fingerprint == fingerprint
    )
    if prepared_is_current:
        st.success(
            f"Department collection completed: {prepared.count} tasks from "
            f"{len(source_spaces)} Space(s). The source is ready for analysis."
        )
        st.download_button(
            "Download Department source Excel",
            prepared.xlsx,
            prepared.filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            on_click="ignore",
        )

    run_clicked = st.button(
        "Run Analysis",
        type="primary",
        disabled=prepared is None or not filtered_tasks,
        key="department_run",
    )
    if run_clicked:
        if not prepared_is_current:
            try:
                with st.spinner("Updating the Department analysis source..."):
                    end_cutoff_value = _analysis_cutoff(
                        start_date, end_date, settings.source_timezone
                    )
                    end_cutoff = end_cutoff_value.isoformat()
                    prepared = collect_data(
                        None,
                        filtered_tasks,
                        selected["name"],
                        fingerprint,
                        settings.source_timezone,
                        space_id="multiple",
                        time_status_data={},
                        time_status_error="",
                        filter_summary=filter_summary,
                        filter_criteria=criteria,
                        analysis_mode="department",
                        department_name=selected["name"],
                        department_id=selected_key,
                        cutoff=end_cutoff,
                    )
                    prepared.space_names = source_spaces
                    prepared.period_start = str(start_date)
                    prepared.period_end = str(end_date)
                    prepared.duplicate_count = st.session_state.get(
                        "department_duplicate_count", 0
                    )
                    st.session_state["clickup_prepared_data"] = prepared
            except Exception as exc:
                st.error(f"Department source could not be updated: {exc}")
                return None, False
        st.session_state["department_run_requested"] = True
        st.rerun()

    run_requested = st.session_state.pop("department_run_requested", False)
    return prepared, run_requested
