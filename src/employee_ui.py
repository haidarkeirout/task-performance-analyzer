"""Employee Performance collection with automatic Jira/ClickUp routing."""
from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass
from typing import Any

import pandas as pd
import streamlit as st

from clickup_export import collect_data as collect_clickup_data
from clickup_gateway import ClickUpCollectionError, ClickUpGateway
from employee_directory import EmployeeRecord, load_employee_directory_with_status
from jira_gateway import CollectionError, JiraGateway
from resumable_jira import collect_jira_query


@dataclass(frozen=True)
class EmployeePreparedBundle:
    """Prepared snapshots for one employee across all connected sources."""

    employee_name: str
    jira: Any | None
    clickup: Any | None
    fingerprint: str

    @property
    def source(self) -> str:
        if self.jira is not None and self.clickup is not None:
            return "Combined"
        return "Jira" if self.jira is not None else "ClickUp"

    @property
    def cutoff(self) -> str:
        values = [getattr(item, "cutoff", "") for item in (self.jira, self.clickup) if item is not None]
        return min(values) if values else ""

    @property
    def source_timezone(self) -> str:
        return getattr(self.clickup or self.jira, "source_timezone", "Asia/Damascus")

    @property
    def count(self) -> int:
        return sum(int(getattr(item, "count", 0) or 0) for item in (self.jira, self.clickup) if item is not None)

    @property
    def filename(self) -> str:
        return f"Employee_{self.employee_name.replace(' ', '_')}_sources.xlsx"


def _source_space_key(source: str, value: object) -> str:
    return f"{source}:{value}"


def _clear_employee_state() -> None:
    for key in (
        "employee_snapshot", "employee_snapshot_key", "employee_selected_spaces",
        "employee_prepared", "employee_fingerprint", "prepared_data", "data_source",
        "clickup_prepared_data", "clickup_analysis", "task_metrics", "process_data",
        "validation_log", "department_analysis", "employee_run_requested",
        "clickup_run_analysis", "cutoff_text", "prepared_data",
        "employee_company_analysis",
    ):
        st.session_state.pop(key, None)


def _employee_label(record: EmployeeRecord) -> str:
    departments = [value for value in (record.jira_department, record.clickup_department) if value]
    suffix = f" · {record.department}" if record.department else ""
    if len(set(departments)) > 1:
        suffix = f" · Jira: {record.jira_department or 'N/A'} · ClickUp: {record.clickup_department or 'N/A'}"
    return f"{record.name}{suffix}"


def _jira_space_name(issue: dict) -> str:
    project = (issue.get("fields") or {}).get("project") or {}
    return project.get("name") or project.get("key") or str(project.get("id") or "Unknown space")


def _jira_space_key(issue: dict) -> str:
    project = (issue.get("fields") or {}).get("project") or {}
    return str(project.get("id") or project.get("key") or "")


def _clickup_assigned(task: dict, user_id: str) -> bool:
    return any(str(person.get("id")) == str(user_id) for person in (task.get("assignees") or []))


def _scaled_progress(progress, fraction: float, text: str, start: float, end: float) -> None:
    if progress is None:
        return
    value = start + (end - start) * max(0.0, min(1.0, fraction))
    progress.progress(value, text=str(text))


def _load_jira(
    record: EmployeeRecord,
    settings,
    *,
    progress=None,
    progress_start: float = 0.0,
    progress_end: float = 1.0,
) -> dict:
    gateway = JiraGateway(settings)
    if progress is None:
        progress = st.progress(0.0, text="Collecting Jira tasks...")
    try:
        query = f'assignee = "{record.jira_account_id}" ORDER BY created DESC'
        issues = gateway.all_issues(
            query,
            progress=lambda message: _scaled_progress(
                progress, 0.0, str(message), progress_start, progress_end
            ),
        )
        spaces = {}
        for issue in issues:
            key = _jira_space_key(issue)
            if key:
                spaces[key] = _jira_space_name(issue)
        _scaled_progress(
            progress, 1.0, f"Jira collection complete: {len(issues)} tasks",
            progress_start, progress_end,
        )
        return {"source": "Jira", "query": query, "issues": issues, "spaces": spaces,
                "definitions": gateway.fields()}
    finally:
        gateway.close()


def _load_clickup(
    record: EmployeeRecord,
    settings,
    *,
    progress=None,
    progress_start: float = 0.0,
    progress_end: float = 1.0,
) -> dict:
    gateway = ClickUpGateway(settings.clickup_token, workspace_id=str(settings.clickup_workspace_id or ""))
    try:
        workspaces = gateway.workspaces()
        workspace_id = str(settings.clickup_workspace_id or (workspaces[0].get("id") if workspaces else ""))
        if not workspace_id:
            raise ClickUpCollectionError("No ClickUp Workspace was found for this account.")
        spaces = gateway.spaces(workspace_id)
        tasks, space_names = [], {}
        if progress is None:
            progress = st.progress(0.0, text=f"Collecting ClickUp Spaces: 0 / {len(spaces)}")
        overall = progress
        registry = dict(st.session_state.get("employee_clickup_checkpoints") or {})
        employee_key = f"{record.name}:{record.clickup_user_id}:{workspace_id}"
        employee_checkpoints = registry.setdefault(employee_key, {})
        for index, space in enumerate(spaces):
            space_id = str(space.get("id") or "")
            if not space_id:
                continue
            checkpoint = employee_checkpoints.setdefault(space_id, {})

            def save_checkpoint(state, current_space_id=space_id):
                employee_checkpoints[current_space_id] = state
                registry[employee_key] = employee_checkpoints
                st.session_state["employee_clickup_checkpoints"] = registry

            current = gateway.all_tasks_for_space(
                space_id,
                progress=lambda message, position=index: _scaled_progress(
                    overall,
                    position / len(spaces) if spaces else 1.0,
                    str(message),
                    progress_start,
                    progress_end,
                ),
                checkpoint=checkpoint,
                checkpoint_callback=save_checkpoint,
            )
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
            _scaled_progress(
                overall,
                (index + 1) / len(spaces) if spaces else 1.0,
                f"Collected ClickUp Space {index + 1} / {len(spaces)}",
                progress_start,
                progress_end,
            )
        _scaled_progress(
            overall, 1.0, f"ClickUp collection complete: {len(tasks)} assigned tasks",
            progress_start, progress_end,
        )
        return {"source": "ClickUp", "workspace_id": workspace_id, "tasks": tasks, "spaces": space_names}
    finally:
        gateway.close()


def _load_employee_sources(record: EmployeeRecord, settings) -> dict:
    """Collect every configured source for the selected employee.

    A failure in one connector is retained as a warning while the other source
    can still be inspected. If both connectors fail, the combined error is
    raised so the user does not mistake an empty result for no assigned work.
    """
    snapshots: dict[str, dict] = {}
    warnings: list[str] = []
    configured_sources = int(bool(record.jira_account_id)) + int(bool(record.clickup_user_id))
    overall = st.progress(0.0, text=f"Searching employee tasks: 0 / {configured_sources} sources")
    source_index = 0
    if record.jira_account_id:
        try:
            snapshots["Jira"] = _load_jira(
                record, settings, progress=overall,
                progress_start=source_index / configured_sources,
                progress_end=(source_index + 1) / configured_sources,
            )
        except CollectionError as exc:
            warnings.append(f"Jira: {exc}")
        source_index += 1
    if record.clickup_user_id:
        try:
            snapshots["ClickUp"] = _load_clickup(
                record, settings, progress=overall,
                progress_start=source_index / configured_sources,
                progress_end=(source_index + 1) / configured_sources,
            )
        except ClickUpCollectionError as exc:
            warnings.append(f"ClickUp: {exc}")
        source_index += 1
    if not snapshots:
        detail = " | ".join(warnings) or "No Jira account or ClickUp member ID is configured."
        raise CollectionError(f"Could not collect {record.name}'s tasks. {detail}")

    spaces: dict[str, str] = {}
    for source, snapshot in snapshots.items():
        for key, label in snapshot.get("spaces", {}).items():
            spaces[_source_space_key(source, key)] = f"{source} · {label}"
    overall.progress(1.0, text=f"Employee collection complete: {sum(len(item.get('issues', item.get('tasks', []))) for item in snapshots.values())} tasks")
    return {
        "source": "Combined" if len(snapshots) > 1 else next(iter(snapshots)),
        "sources": snapshots,
        "spaces": spaces,
        "warnings": warnings,
        "employee_record": record,
    }


def _snapshot_visible(snapshot: dict, selected_spaces: list[str]) -> tuple[list[dict], list[dict]]:
    """Return source-neutral task records and preview rows for a snapshot."""
    sources = snapshot.get("sources") or {snapshot.get("source"): snapshot}
    rows: list[dict] = []
    visible_records: list[dict] = []
    record = snapshot.get("employee_record")
    for source, payload in sources.items():
        if source == "Jira":
            for issue in payload.get("issues", []):
                space_id = _jira_space_key(issue)
                key = _source_space_key(source, space_id)
                if selected_spaces and key not in selected_spaces:
                    continue
                fields = issue.get("fields") or {}
                row = {
                    "Source": "Jira",
                    "Task": issue.get("key"),
                    "Summary": fields.get("summary"),
                    "Space": _jira_space_name(issue),
                    "Status": (fields.get("status") or {}).get("name") if isinstance(fields.get("status"), dict) else fields.get("status"),
                    "Priority": (fields.get("priority") or {}).get("name") if isinstance(fields.get("priority"), dict) else fields.get("priority"),
                    "Assignee": (fields.get("assignee") or {}).get("displayName") if isinstance(fields.get("assignee"), dict) else fields.get("assignee"),
                    "Created": fields.get("created"),
                    "Due Date": fields.get("duedate"),
                    "Department": getattr(record, "jira_department", "") or getattr(record, "department", "") if record else "",
                    "Position": getattr(record, "jira_position", "") if record else "",
                }
                rows.append(row)
                visible_records.append(issue)
        elif source == "ClickUp":
            for task in payload.get("tasks", []):
                space_id = str(task.get("employee_space_id") or task.get("space_id") or "")
                key = _source_space_key(source, space_id)
                if selected_spaces and key not in selected_spaces:
                    continue
                status = task.get("status") or {}
                priority = task.get("priority") or {}
                assignees = task.get("assignees") or []
                row = {
                    "Source": "ClickUp",
                    "Task": task.get("id"),
                    "Summary": task.get("name"),
                    "Space": task.get("employee_space_name") or task.get("space_name"),
                    "Status": status.get("status") if isinstance(status, dict) else status,
                    "Priority": priority.get("priority") if isinstance(priority, dict) else priority,
                    "Assignee": ", ".join(str(item.get("username") or item.get("email") or item.get("id")) for item in assignees),
                    "Created": task.get("date_created"),
                    "Due Date": task.get("due_date"),
                    "Department": getattr(record, "clickup_department", "") or getattr(record, "department", "") if record else "",
                    "Position": getattr(record, "clickup_position", "") if record else "",
                }
                rows.append(row)
                visible_records.append(task)
    return visible_records, rows


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
    total = len(issues)
    progress = st.progress(0.0, text=f"Collecting Jira tasks: 0 / {total}")
    query = snapshot["query"]
    space_name = f"Employee: {record.name}"
    # Include the discovered issue set and a version marker in the persistent
    # fingerprint. Employee snapshots can change while the selected name and
    # Spaces stay the same; an old checkpoint must not be reused for a new set.
    issue_signature = hashlib.sha256(
        json.dumps(
            sorted(str(issue.get("key")) for issue in issues),
            separators=(",", ":"),
        ).encode()
    ).hexdigest()
    durable_fingerprint = f"{fingerprint}:v2:{issue_signature}"
    result = collect_jira_query(
        settings,
        space={"id": "*", "key": f"employee-{record.name}", "name": space_name},
        query=query,
        fingerprint=durable_fingerprint,
        seed_issues=issues,
        progress=lambda fraction, message: progress.progress(
            fraction, text=message
        ),
    )
    progress.progress(1.0, text=f"Collection complete: {total} Jira tasks")
    result.filename = f"Jira_Employee_{record.name.replace(' ', '_')}.xlsx"
    return result


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
        records, directory_warning = load_employee_directory_with_status()
    except ValueError as exc:
        st.error(str(exc))
        return None, False
    if directory_warning:
        st.warning(directory_warning)
    if not records:
        st.info("No active employees are configured yet.")
        return None, False

    names = [record.name for record in records]
    selected_name = st.selectbox(
        "Employee", [None, *names],
        format_func=lambda value: "Choose an employee" if value is None else value,
        key="employee_selected_name", on_change=_clear_employee_state,
    )
    if selected_name is None:
        st.caption("Choose an employee to search Jira and ClickUp automatically.")
        return None, False
    record = next(item for item in records if item.name == selected_name)
    source_labels = " + ".join(source.title() for source in record.sources) or "Not configured"
    st.caption(
        f"Department: {record.department or 'Not specified'} · "
        f"Sources searched automatically: {source_labels}"
    )

    snapshot_key = f"{record.name}:{record.primary_source}:{record.jira_account_id}:{record.clickup_user_id}"
    if (
        st.session_state.get("employee_snapshot_key")
        and st.session_state.get("employee_snapshot_key") != snapshot_key
    ):
        _clear_employee_state()
        st.session_state["employee_selected_spaces"] = []
        st.session_state["employee_snapshot_key"] = snapshot_key
        st.rerun()

    if st.button("Load Employee Tasks", type="primary", key="employee_load"):
        _clear_employee_state()
        try:
            snapshot = _load_employee_sources(record, settings)
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
    for warning in snapshot.get("warnings", []):
        st.warning(f"Some tasks could not be loaded: {warning}")
    spaces = snapshot.get("spaces", {})
    selected_spaces = st.multiselect(
        "Filter by Spaces (optional)", list(spaces),
        format_func=lambda value: spaces[value],
        key="employee_selected_spaces",
    )
    visible, rows = _snapshot_visible(snapshot, selected_spaces)
    st.subheader("Assigned tasks")
    st.caption(
        f"{len(visible)} task(s) found across {len(spaces)} source Space(s). "
        "Leave the filter empty to include all."
    )
    st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)

    fingerprint = hashlib.sha256(
        json.dumps({"snapshot": snapshot_key, "spaces": selected_spaces}, sort_keys=True).encode()
    ).hexdigest()
    prepared = st.session_state.get("employee_prepared")
    prepared_fingerprint = str(getattr(prepared, "fingerprint", "")) if prepared is not None else ""
    same_scope = prepared_fingerprint.endswith(fingerprint) or f":{fingerprint}:v2:" in prepared_fingerprint
    if prepared is not None and not same_scope:
        st.session_state.pop("employee_prepared", None)
        prepared = None

    if st.button("Done", type="primary", disabled=not visible, key="employee_done"):
        try:
            with st.spinner("Preparing the employee analysis source..."):
                source_snapshots = snapshot.get("sources") or {snapshot["source"]: snapshot}
                jira_prepared = None
                clickup_prepared = None
                jira_spaces = [value.split(":", 1)[1] for value in selected_spaces if value.startswith("Jira:")]
                clickup_spaces = [value.split(":", 1)[1] for value in selected_spaces if value.startswith("ClickUp:")]
                if "Jira" in source_snapshots:
                    jira_prepared = _prepare_jira(
                        record, settings, source_snapshots["Jira"], jira_spaces,
                        "employee:jira:" + fingerprint,
                    )
                if "ClickUp" in source_snapshots:
                    clickup_prepared = _prepare_clickup(
                        record, settings, source_snapshots["ClickUp"], clickup_spaces,
                        "employee:clickup:" + fingerprint,
                    )
                if jira_prepared is not None and clickup_prepared is not None:
                    prepared = EmployeePreparedBundle(record.name, jira_prepared, clickup_prepared, "employee:combined:" + fingerprint)
                else:
                    prepared = jira_prepared or clickup_prepared
            st.session_state["employee_prepared"] = prepared
            st.session_state["prepared_data"] = prepared
            st.success("Employee task collection completed. The source is ready for analysis.")
        except (CollectionError, ClickUpCollectionError) as exc:
            st.error(str(exc))
            return None, False
    prepared = st.session_state.get("employee_prepared")
    run_clicked = st.button("Run Analysis", type="primary", disabled=prepared is None, key="employee_run")
    if prepared:
        if isinstance(prepared, EmployeePreparedBundle):
            downloads = st.columns(2)
            for column, item, label in (
                (downloads[0], prepared.jira, "Download Jira Source Excel"),
                (downloads[1], prepared.clickup, "Download ClickUp Source Excel"),
            ):
                if item is not None:
                    column.download_button(label, item.xlsx, item.filename,
                                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", on_click="ignore")
        else:
            st.download_button("Download Source Excel", prepared.xlsx, prepared.filename,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", on_click="ignore")
    if run_clicked:
        st.session_state["employee_run_requested"] = True
        st.rerun()
    run_requested = st.session_state.pop("employee_run_requested", False)
    return prepared, run_requested

