"""Employee Performance collection with automatic Jira/ClickUp routing."""
from __future__ import annotations

import hashlib
import json
from datetime import datetime, timezone

import pandas as pd
import streamlit as st

from clickup_export import collect_data as collect_clickup_data
from clickup_gateway import ClickUpCollectionError, ClickUpGateway
from employee_directory import EmployeeRecord, load_employee_directory
from jira_export import PreparedData, _json, build_workbook
from jira_gateway import CollectionError, JiraGateway


def _clear_employee_state() -> None:
    for key in (
        "employee_snapshot", "employee_snapshot_key", "employee_selected_spaces",
        "employee_prepared", "employee_fingerprint", "prepared_data", "data_source",
        "clickup_prepared_data", "clickup_analysis", "task_metrics", "process_data",
        "validation_log", "department_analysis",
    ):
        st.session_state.pop(key, None)


def _employee_label(record: EmployeeRecord) -> str:
    suffix = f" · {record.department}" if record.department else ""
    return f"{record.name}{suffix}"


def _jira_space_name(issue: dict) -> str:
    project = (issue.get("fields") or {}).get("project") or {}
    return project.get("name") or project.get("key") or str(project.get("id") or "Unknown space")


def _jira_space_key(issue: dict) -> str:
    project = (issue.get("fields") or {}).get("project") or {}
    return str(project.get("id") or project.get("key") or "")


def _clickup_assigned(task: dict, user_id: str) -> bool:
    return any(str(person.get("id")) == str(user_id) for person in (task.get("assignees") or []))


def _load_jira(record: EmployeeRecord, settings) -> dict:
    gateway = JiraGateway(settings)
    try:
        query = f'assignee = "{record.jira_account_id}" ORDER BY created DESC'
        issues = gateway.all_issues(query, progress=lambda message: st.caption(message))
        spaces = {}
        for issue in issues:
            key = _jira_space_key(issue)
            if key:
                spaces[key] = _jira_space_name(issue)
        return {"source": "Jira", "query": query, "issues": issues, "spaces": spaces,
                "definitions": gateway.fields()}
    finally:
        gateway.close()


def _load_clickup(record: EmployeeRecord, settings) -> dict:
    gateway = ClickUpGateway(settings.clickup_token, workspace_id=str(settings.clickup_workspace_id or ""))
    try:
        workspaces = gateway.workspaces()
        workspace_id = str(settings.clickup_workspace_id or (workspaces[0].get("id") if workspaces else ""))
        if not workspace_id:
            raise ClickUpCollectionError("No ClickUp Workspace was found for this account.")
        spaces = gateway.spaces(workspace_id)
        tasks, space_names = [], {}
        for space in spaces:
            space_id = str(space.get("id") or "")
            if not space_id:
                continue
            current = gateway.all_tasks_for_space(space_id, progress=lambda message: st.caption(message))
            matching = []
            for task in current:
                if _clickup_assigned(task, record.clickup_user_id):
                    item = dict(task)
                    item["employee_space_id"] = space_id
                    item["employee_space_name"] = space.get("name") or space_id
                    item.setdefault("space_id", space_id)
                    matching.append(item)
            if matching:
                space_names[space_id] = space.get("name") or space_id
                tasks.extend(matching)
        return {"source": "ClickUp", "workspace_id": workspace_id, "tasks": tasks, "spaces": space_names}
    finally:
        gateway.close()


def _update_employee_progress(progress, message: str, total: int) -> None:
    """Update one compact collection bar instead of rendering one row per task."""
    text = str(message or "")
    prefix = "Preparing ClickUp task "
    if text.startswith(prefix):
        try:
            current = int(text[len(prefix):].split(" of ", 1)[0])
        except (TypeError, ValueError):
            current = None
        if current is not None:
            progress.progress(
                min(current / total, 1.0) if total else 1.0,
                text=f"Collecting ClickUp tasks: {current} / {total}",
            )
            return
    progress.progress(0.0, text=f"Collecting ClickUp tasks: 0 / {total}")


def _prepare_jira(record, settings, snapshot, selected_spaces, fingerprint):
    issues = [issue for issue in snapshot["issues"] if not selected_spaces or _jira_space_key(issue) in selected_spaces]
    if not issues:
        raise CollectionError("No tasks match the selected Spaces.")
    gateway = JiraGateway(settings)
    try:
        cutoff = datetime.now(timezone.utc).isoformat()
        histories, complete = {}, []
        total = len(issues)
        progress = st.progress(0.0, text=f"Collecting Jira tasks: 0 / {total}")
        for index, issue in enumerate(issues, 1):
            item, history = gateway.complete_issue(issue)
            if history.get("history_complete") is not True or not history.get("history_through"):
                raise CollectionError("A Jira task history is incomplete. No partial export was prepared.")
            complete.append(item)
            histories[item["key"]] = history
            progress.progress(index / total, text=f"Collecting Jira tasks: {index} / {total}")
        progress.progress(1.0, text=f"Collection complete: {total} Jira tasks")
        query = snapshot["query"]
        space_name = f"Employee: {record.name}"
        xlsx = build_workbook(
            complete, histories, snapshot["definitions"], cutoff=cutoff,
            collected_at=datetime.now(timezone.utc).isoformat(), query=query,
            space_name=space_name, source_timezone=settings.source_timezone,
            preferred_start=settings.start_date_field,
        )
        return PreparedData(
            xlsx, _json(histories).encode(), cutoff, datetime.now(timezone.utc).isoformat(),
            query, fingerprint, len(complete), f"Jira_Employee_{record.name.replace(' ', '_')}.xlsx",
            space_name, settings.source_timezone,
        )
    finally:
        gateway.close()


def _prepare_clickup(record, settings, snapshot, selected_spaces, fingerprint):
    tasks = [task for task in snapshot["tasks"] if not selected_spaces or task.get("employee_space_id") in selected_spaces]
    if not tasks:
        raise ClickUpCollectionError("No tasks match the selected Spaces.")
    names = sorted({task.get("employee_space_name") for task in tasks if task.get("employee_space_name")})
    total = len(tasks)
    progress = st.progress(0.0, text=f"Collecting ClickUp tasks: 0 / {total}")
    prepared = collect_clickup_data(
        None, tasks, f"Employee: {record.name}", fingerprint, settings.source_timezone,
        progress=lambda message: _update_employee_progress(progress, message, total),
        space_id=None,
        filter_summary=f"Assigned tasks for {record.name}" + (f" · Spaces: {', '.join(names)}" if names else ""),
        filter_criteria={"employee": record.name, "spaces": names},
        analysis_mode="employee",
    )
    progress.progress(1.0, text=f"Collection complete: {total} ClickUp tasks")
    return prepared


@st.fragment
def render_employee_collection(settings):
    try:
        records = load_employee_directory()
    except ValueError as exc:
        st.error(str(exc))
        return None, False
    if not records:
        st.info("No active employees are configured yet.")
        return None, False

    names = [record.name for record in records]
    selected_name = st.selectbox("Employee", [None, *names], format_func=lambda value: "Choose an employee" if value is None else value,
                                 key="employee_selected_name", on_change=_clear_employee_state)
    if selected_name is None:
        st.caption("Choose an employee to discover their source, Spaces, and assigned tasks automatically.")
        return None, False
    record = next(item for item in records if item.name == selected_name)
    st.caption(f"Department: {record.department or 'Not specified'} · Source detected automatically: {record.primary_source.title()}")

    snapshot_key = f"{record.name}:{record.primary_source}:{record.jira_account_id}:{record.clickup_user_id}"
    if st.button("Load Employee Tasks", type="primary", key="employee_load"):
        _clear_employee_state()
        try:
            with st.spinner(f"Finding {record.name}'s assigned tasks..."):
                snapshot = _load_jira(record, settings) if record.primary_source == "jira" else _load_clickup(record, settings)
            st.session_state["employee_snapshot"] = snapshot
            st.session_state["employee_snapshot_key"] = snapshot_key
            st.session_state["data_source"] = snapshot["source"]
            st.rerun()
        except (CollectionError, ClickUpCollectionError) as exc:
            st.error(str(exc))
            return None, False

    snapshot = st.session_state.get("employee_snapshot")
    if st.session_state.get("employee_snapshot_key") != snapshot_key:
        snapshot = None
    if snapshot is None:
        return None, False

    st.session_state["data_source"] = snapshot["source"]
    spaces = snapshot.get("spaces", {})
    selected_spaces = st.multiselect(
        "Filter by Spaces (optional)", list(spaces), format_func=lambda value: spaces[value],
        key="employee_selected_spaces",
    )
    if snapshot["source"] == "Jira":
        visible = [issue for issue in snapshot["issues"] if not selected_spaces or _jira_space_key(issue) in selected_spaces]
        rows = [{"Task": issue.get("key"), "Summary": (issue.get("fields") or {}).get("summary"), "Space": _jira_space_name(issue)} for issue in visible]
    else:
        visible = [task for task in snapshot["tasks"] if not selected_spaces or task.get("employee_space_id") in selected_spaces]
        rows = [{"Task": task.get("id"), "Summary": task.get("name"), "Space": task.get("employee_space_name")} for task in visible]
    st.subheader("Assigned tasks")
    st.caption(f"{len(visible)} task(s) found · {len(spaces)} Space(s). Leave the filter empty to include all Spaces.")
    st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)

    fingerprint = hashlib.sha256(json.dumps({"snapshot": snapshot_key, "spaces": selected_spaces}, sort_keys=True).encode()).hexdigest()
    prepared = st.session_state.get("employee_prepared")
    if prepared is not None and getattr(prepared, "fingerprint", "") != "employee:" + snapshot["source"].casefold() + ":" + fingerprint:
        st.session_state.pop("employee_prepared", None)
        prepared = None
    if st.button("Done", type="primary", disabled=not visible, key="employee_done"):
        try:
            with st.spinner("Preparing the employee analysis source..."):
                prepared = (
                    _prepare_jira(record, settings, snapshot, selected_spaces, "employee:jira:" + fingerprint)
                    if snapshot["source"] == "Jira"
                    else _prepare_clickup(record, settings, snapshot, selected_spaces, "employee:clickup:" + fingerprint)
                )
            st.session_state["employee_prepared"] = prepared
            st.session_state["prepared_data"] = prepared
            st.success("Employee task collection completed. The source is ready for analysis.")
        except (CollectionError, ClickUpCollectionError) as exc:
            st.error(str(exc))
            return None, False
    prepared = st.session_state.get("employee_prepared")
    run_clicked = st.button("Run Analysis", type="primary", disabled=prepared is None, key="employee_run")
    if prepared:
        st.download_button("Download Source Excel", prepared.xlsx, prepared.filename,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", on_click="ignore")
    if run_clicked:
        # The fragment keeps widget interactions in place, then asks the full
        # app rerun to consume the Run Analysis event in app.py.
        st.session_state["employee_run_requested"] = True
        st.rerun()
    run_requested = st.session_state.pop("employee_run_requested", False)
    return prepared, run_requested
