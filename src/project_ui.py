"""Project Performance: dual-source space selection and combined analysis."""
from __future__ import annotations

from collections import Counter, defaultdict
from datetime import date, datetime, timedelta, timezone
from io import BytesIO
from pathlib import Path
from typing import Any
from uuid import uuid4

import pandas as pd
import streamlit as st
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from excel_safety import write_excel_cell

from clickup_export import collect_data as collect_clickup_data
from clickup_gateway import ClickUpCollectionError, ClickUpGateway
from company_performance.application import CompanyAnalysisResult, build_company_analysis, build_company_preview
from company_performance.kpis import calculate_status_metrics
from kpi_transparency import population_from_snapshots, render_population_card
from jira_export import PreparedData, _json, build_workbook
from jira_gateway import CollectionError, JiraGateway
from resumable_jira import collect_jira_query
from standard_report_style import (
    add_report_table,
    add_report_title,
    configure_report_document,
)


_XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
_DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
_HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
_SECTION_FILL = PatternFill("solid", fgColor="D9EAF7")
_HEADER_FONT = Font(color="FFFFFF", bold=True)


def _clear_project_run() -> None:
    for key in (
        "project_selection_fingerprint",
        "project_jira_prepared",
        "project_clickup_prepared",
        "project_preview",
        "project_analysis",
        "project_scope_label",
        "project_scope_slug",
        "project_report_key",
        "project_collection_id",
        "project_collection_collected_at",
    ):
        st.session_state.pop(key, None)


def _clear_project_analysis() -> None:
    """Invalidate only the displayed result when the period changes."""
    for key in (
        "project_analysis",
        "project_report_key",
        "project_excel_bytes",
        "project_word_bytes",
    ):
        st.session_state.pop(key, None)


def _safe_filename(value: str) -> str:
    value = str(value or "").strip()
    value = value.replace("&", "and").replace("+", "and")
    value = value.replace("—", "_").replace("–", "_")
    value = "_".join(value.split())
    cleaned = "".join(char if char.isalnum() or char in "._-" else "_" for char in value)
    while "__" in cleaned:
        cleaned = cleaned.replace("__", "_")
    return cleaned.strip("_") or "Selected_Scope"


def _scope_label(jira_item: dict[str, Any] | None, clickup_item: dict[str, Any] | None) -> tuple[str, str]:
    parts = []
    names = []
    if jira_item:
        name = str(jira_item.get("name") or jira_item.get("key") or jira_item.get("id"))
        parts.append(f"Jira — {name}")
        names.append(name)
    if clickup_item:
        name = str(clickup_item.get("name") or clickup_item.get("id"))
        parts.append(f"ClickUp — {name}")
        names.append(name)
    return " + ".join(parts), "_and_".join(_safe_filename(name) for name in names)


def _load_catalogs(settings) -> tuple[dict[str, dict[str, Any]], dict[str, dict[str, Any]]]:
    revision = str(getattr(settings, "revision", ""))
    if st.session_state.get("project_catalog_revision") == revision:
        return (
            dict(st.session_state.get("project_jira_spaces", {})),
            dict(st.session_state.get("project_clickup_spaces", {})),
        )

    jira_spaces: dict[str, dict[str, Any]] = {}
    clickup_spaces: dict[str, dict[str, Any]] = {}
    st.session_state["project_jira_error"] = ""
    st.session_state["project_clickup_error"] = ""

    try:
        gateway = JiraGateway(settings)
        try:
            for item in gateway.spaces():
                identifier = str(item.get("id") or item.get("key") or "")
                if identifier:
                    jira_spaces[identifier] = dict(item)
        finally:
            gateway.close()
    except CollectionError as exc:
        st.session_state["project_jira_error"] = str(exc)

    try:
        gateway = ClickUpGateway(
            settings.clickup_token,
            workspace_id=str(settings.clickup_workspace_id or ""),
        )
        try:
            workspaces = gateway.workspaces()
            workspace_id = str(
                settings.clickup_workspace_id
                or (workspaces[0].get("id") if workspaces else "")
            )
            if workspace_id:
                for item in gateway.spaces(workspace_id):
                    identifier = str(item.get("id") or "")
                    if identifier:
                        clickup_spaces[identifier] = dict(item)
        finally:
            gateway.close()
    except ClickUpCollectionError as exc:
        st.session_state["project_clickup_error"] = str(exc)

    st.session_state["project_catalog_revision"] = revision
    st.session_state["project_jira_spaces"] = jira_spaces
    st.session_state["project_clickup_spaces"] = clickup_spaces
    return jira_spaces, clickup_spaces


def _collect_jira_space(
    settings,
    item: dict[str, Any],
    fingerprint: str,
    progress=None,
    *,
    collection_id: str | None = None,
    fresh: bool = False,
) -> PreparedData:
    project_key = str(item.get("key") or "")
    project_name = str(item.get("name") or project_key or item.get("id") or "Jira Project")
    if not project_key:
        raise CollectionError("The selected Jira Space has no usable project key.")

    query = f'project = "{project_key}" ORDER BY created DESC'
    collection_progress = None if progress is not None else st.progress(
        0, text="Finding tasks in Jira..."
    )

    def report_progress(fraction: float, message: str) -> None:
        if progress is not None:
            progress(fraction, message)
        elif collection_progress is not None:
            collection_progress.progress(fraction, text=message)

    report_progress(
        0,
        f"Finding tasks in Jira — {project_name}..."
        if progress is not None else "Finding tasks in Jira...",
    )
    return collect_jira_query(
        settings,
        space=item,
        query=query,
        fingerprint=fingerprint,
        collection_id=collection_id,
        fresh=fresh,
        progress=report_progress,
    )


def _collect_clickup_space(
    settings,
    item: dict[str, Any],
    fingerprint: str,
    progress=None,
    *,
    collection_id: str | None = None,
    fresh: bool = False,
):
    effective_fingerprint = (
        f"{fingerprint}:collection:{collection_id}"
        if collection_id else fingerprint
    )
    space_id = str(item.get("id") or "")
    space_name = str(item.get("name") or space_id or "ClickUp Space")
    if not space_id:
        raise ClickUpCollectionError("The selected ClickUp Space has no usable ID.")

    gateway = ClickUpGateway(
        settings.clickup_token,
        workspace_id=str(settings.clickup_workspace_id or ""),
    )
    collection_progress = None if progress is not None else st.progress(
        0,
        text=f"Loading task list from ClickUp — {space_name}...",
    )

    def report_progress(fraction: float, message: str) -> None:
        if progress is not None:
            progress(fraction, message)
        elif collection_progress is not None:
            collection_progress.progress(fraction, text=message)

    try:
        report_progress(0, f"Loading task list from ClickUp — {space_name}...")
        checkpoints = dict(st.session_state.get("project_clickup_checkpoints") or {})
        checkpoint = checkpoints.setdefault(fingerprint, {})

        def save_checkpoint(state):
            checkpoints[fingerprint] = state
            st.session_state["project_clickup_checkpoints"] = checkpoints

        tasks = gateway.all_tasks_for_space(
            space_id,
            progress=lambda message: report_progress(0.0, message),
            checkpoint=checkpoint,
            checkpoint_callback=save_checkpoint,
            fresh=fresh,
        )
    finally:
        gateway.close()

    prepared_tasks = []
    for task in tasks:
        copy_task = dict(task)
        copy_task["space_id"] = space_id
        copy_task["project_space_name"] = space_name
        prepared_tasks.append(copy_task)

    total = len(prepared_tasks)
    report_progress(0, f"Collecting tasks: 0 / {total}")
    collected = 0

    def update_progress(_message):
        nonlocal collected
        collected += 1
        report_progress(
            collected / total if total else 1.0,
            f"Collecting tasks: {collected} / {total}",
        )

    prepared = collect_clickup_data(
        None,
        prepared_tasks,
        space_name,
        effective_fingerprint,
        settings.source_timezone,
        progress=update_progress,
        space_id=space_id,
        filter_summary=f"Selected ClickUp Space: {space_name}",
        filter_criteria={"space": space_name, "space_id": space_id},
        analysis_mode="project",
    )
    report_progress(
        1.0,
        f"Collection complete: {total} / {total} tasks",
    )
    return prepared


def _preview_frame(preview) -> pd.DataFrame:
    return pd.DataFrame([
        {
            "Source": row.source_tool,
            "Space": row.space,
            "Task": row.task_name,
            "Original Status": row.original_status,
            "Final Status": row.final_status,
            "Assignee": row.assignee,
            "Priority": row.priority,
            "Due Date": row.due_date,
        }
        for row in preview
    ])


def _task_detail_frame(result: CompanyAnalysisResult) -> pd.DataFrame:
    rows = []
    for snapshot in result.snapshots:
        task = snapshot.task
        completed = snapshot.status_at_period_end.value == "Completed"
        late = bool(completed and task.due_date and snapshot.final_completion_date and snapshot.final_completion_date > task.due_date)
        overdue = bool(not completed and task.due_date and task.due_date < snapshot.period_end)
        rows.append({
            "Source": task.source_tool,
            "Space": task.source_space,
            "Task ID": task.task_id,
            "Task": task.task_name,
            "Assignee": task.assignee_group,
            "Priority": task.priority,
            "Task Type": "Subtask" if task.parent_id else "Task",
            "Original Status": task.raw_status,
            "Final Status": snapshot.status_at_period_end.value,
            "Created Date": task.created_date,
            "Start Date": snapshot.actual_start_date,
            "Due Date": task.due_date,
            "Completed Date": snapshot.final_completion_date,
            "Completed Late": late,
            "Open Overdue": overdue,
            "Workflow History Available": task.history_complete,
            "Data Quality Flags": "; ".join(sorted(snapshot.data_quality_flags)),
        })
    return pd.DataFrame(rows, columns=[
        "Source", "Space", "Task ID", "Task", "Assignee", "Priority", "Task Type",
        "Original Status", "Final Status", "Created Date", "Start Date", "Due Date",
        "Completed Date", "Completed Late", "Open Overdue", "Workflow History Available",
        "Data Quality Flags",
    ])


def _weekly_flow_frame(result: CompanyAnalysisResult) -> pd.DataFrame:
    values: dict[date, dict[str, int]] = defaultdict(lambda: {"Tasks Created": 0, "Tasks Completed": 0})
    for snapshot in result.snapshots:
        created = snapshot.task.created_date
        completed = snapshot.final_completion_date
        if created:
            week = created - timedelta(days=created.weekday())
            values[week]["Tasks Created"] += 1
        if completed:
            week = completed - timedelta(days=completed.weekday())
            values[week]["Tasks Completed"] += 1
    rows = [{"Week Starting": key, **value} for key, value in sorted(values.items())]
    return pd.DataFrame(rows, columns=["Week Starting", "Tasks Created", "Tasks Completed"])


def _source_space_frame(result: CompanyAnalysisResult) -> pd.DataFrame:
    rows = []
    grouped: dict[tuple[str, str], dict[str, int]] = defaultdict(lambda: {
        "Total Tasks": 0, "Completed": 0, "Open Overdue": 0,
    })
    for snapshot in result.snapshots:
        key = (snapshot.task.source_tool, snapshot.task.source_space or "Unknown")
        grouped[key]["Total Tasks"] += 1
        if snapshot.status_at_period_end.value == "Completed":
            grouped[key]["Completed"] += 1
        if (
            snapshot.status_at_period_end.is_open
            and snapshot.task.due_date
            and snapshot.task.due_date < snapshot.period_end
        ):
            grouped[key]["Open Overdue"] += 1
    for (source, space), values in sorted(grouped.items()):
        rows.append({"Source": source, "Space": space, **values})
    return pd.DataFrame(rows, columns=["Source", "Space", "Total Tasks", "Completed", "Open Overdue"])


def _assignee_frame(result: CompanyAnalysisResult) -> pd.DataFrame:
    grouped: dict[str, dict[str, int]] = defaultdict(lambda: {
        "Total Tasks": 0, "Completed": 0, "Open": 0, "Open Overdue": 0,
    })
    for snapshot in result.snapshots:
        assignee = snapshot.task.assignee_group
        values = grouped[assignee]
        values["Total Tasks"] += 1
        if snapshot.status_at_period_end.value == "Completed":
            values["Completed"] += 1
        if snapshot.status_at_period_end.is_open:
            values["Open"] += 1
        if (
            snapshot.status_at_period_end.is_open
            and snapshot.task.due_date
            and snapshot.task.due_date < snapshot.period_end
        ):
            values["Open Overdue"] += 1
    return pd.DataFrame([
        {"Assignee": assignee, **values}
        for assignee, values in sorted(grouped.items())
    ], columns=["Assignee", "Total Tasks", "Completed", "Open", "Open Overdue"])


def _write_excel_table(sheet, headers, rows, start_row=1):
    for row_index, row in enumerate([headers, *rows], start=start_row):
        for column_index, value in enumerate(row, 1):
            cell = write_excel_cell(sheet, row_index, column_index, _excel_value(value))
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if row_index == start_row:
                cell.fill = _HEADER_FILL
                cell.font = _HEADER_FONT
    return start_row + len(rows) + 1


def _excel_value(value):
    if value is None or value == "":
        return "N/A"
    if isinstance(value, (tuple, list, set)):
        return "; ".join(str(item) for item in value) if value else "N/A"
    if isinstance(value, date):
        return value.isoformat()
    return value


def _fit_sheet(sheet):
    for column_index in range(1, sheet.max_column + 1):
        width = 12
        for row_index in range(1, min(sheet.max_row, 250) + 1):
            width = max(width, min(42, len(str(sheet.cell(row_index, column_index).value or "")) + 2))
        sheet.column_dimensions[get_column_letter(column_index)].width = width
    sheet.freeze_panes = "A2"


def _kpi_detail_rows(model) -> list[tuple[str, Any, str]]:
    kpis = model.kpis
    rate = lambda value: "N/A" if value is None else f"{value:.1f}%"
    days = lambda value: "N/A" if value is None else f"{value:.1f}"
    return [
        ("Total Tasks", kpis.total_tasks, "All eligible tasks collected from the selected Jira/ClickUp Spaces."),
        ("Completed Tasks", kpis.completed_tasks, "Tasks in Completed status at the end of the selected period."),
        ("Open Tasks", kpis.open_tasks, "Tasks in an open workflow status at period end."),
        ("Current WIP", kpis.current_wip, "Tasks currently in In Execution or In Review."),
        ("Completion Rate", rate(kpis.completion_rate), "Completed tasks divided by tasks with a verified status at period end; Unknown statuses are excluded and shown in Data Quality."),
        ("Open Overdue Tasks", kpis.overdue_open_tasks, "Open tasks with a due date before the period end."),
        ("Open Overdue Rate", rate(kpis.overdue_open_rate), "Open overdue tasks divided by open tasks with a due date."),
        ("High-Priority Open Tasks", kpis.high_priority_open_tasks, "Open Critical or High priority tasks."),
        ("High-Priority Overdue Tasks", kpis.high_priority_overdue_tasks, "Critical or High priority tasks that are open and overdue."),
        ("Unassigned Open Tasks", kpis.unassigned_open_tasks, "Open tasks without an assigned owner."),
        ("On-Time Completion Rate", rate(kpis.on_time_completion_rate), "Completed tasks with a due date completed on or before that date."),
        ("Late Completion Rate", rate(kpis.late_completion_rate), "Completed tasks with a due date completed after that date."),
        ("Average Time to Start (days)", days(kpis.average_time_to_start_days), "Average time from creation to the first recorded start."),
        ("Average Execution Duration (days)", days(kpis.average_execution_duration_days), "Average time from start to completion."),
        ("Average Lead Time (days)", days(kpis.average_lead_time_days), "Average time from creation to completion."),
        ("Cancelled Tasks", kpis.cancelled_tasks, "Tasks cancelled at period end."),
        ("Rejected Tasks", kpis.rejected_tasks, "Tasks rejected at period end."),
    ]


def _status_analysis_frame(result: CompanyAnalysisResult) -> pd.DataFrame:
    rows = []
    for metric in calculate_status_metrics(result.snapshots):
        rows.append({
            "Status": metric.status.value,
            "Tasks Passed Through": metric.tasks_passed_through,
            "Average Days": metric.average_days,
            "Median Days": metric.median_days,
            "Total Days": metric.total_days,
            "Open Tasks at Period End": metric.open_tasks_now,
            "Open Overdue Tasks": metric.overdue_open_tasks,
            "Repeated Returns": metric.repeated_returns,
            "History Covered": metric.history_covered,
        })
    return pd.DataFrame(rows)


def _workflow_events_frame(result: CompanyAnalysisResult) -> pd.DataFrame:
    rows = []
    for snapshot in result.snapshots:
        task = snapshot.task
        for event in task.workflow_history:
            changed = event.changed_at.date()
            if snapshot.period_start <= changed <= snapshot.period_end:
                rows.append({
                    "Source": task.source_tool,
                    "Space": task.source_space,
                    "Task ID": task.task_id,
                    "Task": task.task_name,
                    "Changed Date": changed,
                    "From Status": event.from_status or "N/A",
                    "To Status": event.to_status or "N/A",
                    "Performed By": event.performed_by or "N/A",
                })
    return pd.DataFrame(rows, columns=[
        "Source", "Space", "Task ID", "Task", "Changed Date", "From Status",
        "To Status", "Performed By",
    ])


def _exception_frame(result: CompanyAnalysisResult) -> pd.DataFrame:
    rows = []
    for snapshot in result.snapshots:
        task = snapshot.task
        for exception in snapshot.exception_events:
            rows.append({
                "Source": task.source_tool,
                "Space": task.source_space,
                "Task ID": task.task_id,
                "Task": task.task_name,
                "Exception": exception,
            })
    return pd.DataFrame(rows, columns=["Source", "Space", "Task ID", "Task", "Exception"])


def _recommendation_rows(result: CompanyAnalysisResult) -> list[tuple[Any, ...]]:
    return [
        (item.severity, item.title, item.evidence, item.suggested_action)
        for item in result.recommendations
    ]



def _project_period_label(result: CompanyAnalysisResult) -> str:
    return f"{result.model.period_start.isoformat()} to {result.model.period_end.isoformat()}"


def _project_period_token(result: CompanyAnalysisResult) -> str:
    return f"{result.model.period_start.isoformat()}_to_{result.model.period_end.isoformat()}"


def _project_status_counts(snapshots) -> dict[str, int]:
    counts: dict[str, int] = Counter()
    for snapshot in snapshots:
        if snapshot.counted_in_kpis:
            counts[snapshot.status_at_period_end.value] += 1
    return dict(counts)


def _project_event_counts(snapshot) -> dict[str, int]:
    counts = {"rework": 0, "replanning": 0, "re_evaluation": 0, "reopen": 0}
    for event in snapshot.task.workflow_history:
        before = (event.from_status or "").casefold()
        after = (event.to_status or "").casefold()
        if "in review" in before and ("in progress" in after or "in execution" in after):
            counts["rework"] += 1
        elif "in review" in before and "to do" in after:
            counts["replanning"] += 1
        elif "in review" in before and "triage" in after:
            counts["re_evaluation"] += 1
        if "done" in before and after not in {"done", "completed"}:
            counts["reopen"] += 1
    return counts


def _project_average(values: list[float]) -> float | None:
    return round(sum(values) / len(values), 2) if values else None


def _project_median(values: list[float]) -> float | None:
    if not values:
        return None
    values = sorted(values)
    middle = len(values) // 2
    if len(values) % 2:
        return round(values[middle], 2)
    return round((values[middle - 1] + values[middle]) / 2, 2)


def _project_elapsed_hours(start: date | None, end: date | None) -> float | None:
    if not start or not end:
        return None
    return round(max(0, (end - start).days) * 24.0, 2)


def _project_last_status_date(snapshot) -> date | None:
    events = [event.changed_at.date() for event in snapshot.task.workflow_history if event.changed_at]
    return max(events) if events else snapshot.task.created_date


def _project_summary_values(snapshots) -> dict[str, Any]:
    items = [item for item in snapshots if item.counted_in_kpis]
    total = len(items)
    known = [item for item in items if item.status_at_period_end.value != "Unknown"]
    completed = [item for item in known if item.status_at_period_end.value == "Completed"]
    rejected = [item for item in known if item.status_at_period_end.value == "Rejected"]
    open_items = [item for item in known if item.status_at_period_end.is_open]
    wip = [item for item in items if item.status_at_period_end.value in {"In Execution", "In Review"}]
    completed_due = [
        item for item in completed
        if item.task.due_date and item.final_completion_date
    ]
    on_time = [
        item for item in completed_due
        if item.final_completion_date <= item.task.due_date
    ]
    overdue = [
        item for item in open_items
        if item.task.due_date and item.task.due_date < item.period_end
    ]
    open_with_due = [item for item in open_items if item.task.due_date]
    reviewed = [item for item in items if item.history_available]
    event_counts = [_project_event_counts(item) for item in items]
    execution_hours = [
        value for value in (
            _project_elapsed_hours(item.actual_start_date, item.final_completion_date)
            for item in completed
        ) if value is not None
    ]
    lead_hours = [
        value for value in (
            _project_elapsed_hours(item.task.created_date, item.final_completion_date)
            for item in completed
        ) if value is not None
    ]
    start_hours = [
        value for value in (
            _project_elapsed_hours(item.task.created_date, item.actual_start_date)
            for item in items
        ) if value is not None
    ]
    rework_events = sum(value["rework"] for value in event_counts)
    replanning_events = sum(value["replanning"] for value in event_counts)
    re_evaluation_events = sum(value["re_evaluation"] for value in event_counts)
    rework_tasks = sum(value["rework"] > 0 for value in event_counts)
    replanning_tasks = sum(value["replanning"] > 0 for value in event_counts)
    re_evaluation_tasks = sum(value["re_evaluation"] > 0 for value in event_counts)
    review_exception_tasks = sum(
        any(value[key] > 0 for key in ("rework", "replanning", "re_evaluation"))
        for value in event_counts
    )
    rate = lambda numerator, denominator: None if not denominator else round(numerator / denominator * 100.0, 1)
    return {
        "total_tasks": total,
        "completed_tasks": len(completed),
        "rejected_tasks": len(rejected),
        "open_tasks": len(open_items),
        "wip_tasks": len(wip),
        "completion_rate": rate(len(completed), len(known)),
        "rejection_rate": rate(len(rejected), total),
        "on_time_tasks": len(on_time),
        "on_time_valid_tasks": len(completed_due),
        "on_time_completion_rate": rate(len(on_time), len(completed_due)),
        "overdue_open_tasks": len(overdue),
        "overdue_valid_tasks": len(open_with_due),
        "open_overdue_rate": rate(len(overdue), len(open_with_due)),
        "tasks_with_rework": rework_tasks,
        "rework_valid_tasks": len(reviewed),
        "tasks_with_rework_rate": rate(rework_tasks, len(reviewed)),
        "mean_execution_business_hours": None,
        "median_execution_business_hours": None,
        "mean_lead_time_business_hours": None,
        "median_lead_time_business_hours": None,
        "mean_rework_count": round(rework_events / total, 2) if total else 0,
        "total_rework_count": rework_events,
        "history_complete_tasks": sum(item.task.history_complete for item in items),
        "unknown_status_tasks": sum(item.status_at_period_end.value == "Unknown" for item in items),
        "known_status_tasks": len(known),
        "history_excluded_tasks": sum(not item.history_available for item in items),
        "reviewed_valid_tasks": len(reviewed),
        "pending_before_execution": sum(item.status_at_period_end.value == "Not Started" for item in items),
        "rework_rate": rate(rework_tasks, len(reviewed)),
        "tasks_with_replanning": replanning_tasks,
        "total_replanning_count": replanning_events,
        "replanning_rate": rate(replanning_tasks, len(reviewed)),
        "tasks_with_re_evaluation": re_evaluation_tasks,
        "total_re_evaluation_count": re_evaluation_events,
        "re_evaluation_rate": rate(re_evaluation_tasks, len(reviewed)),
        "review_exception_tasks": review_exception_tasks,
        "review_exception_rate": rate(review_exception_tasks, len(reviewed)),
        "mean_execution_elapsed_hours": _project_average(execution_hours),
        "median_execution_elapsed_hours": _project_median(execution_hours),
        "execution_elapsed_hours_valid_tasks": len(execution_hours),
        "execution_business_hours_valid_tasks": 0,
        "mean_lead_time_elapsed_hours": _project_average(lead_hours),
        "median_lead_time_elapsed_hours": _project_median(lead_hours),
        "lead_time_elapsed_hours_valid_tasks": len(lead_hours),
        "lead_time_business_hours_valid_tasks": 0,
        "mean_time_to_start_elapsed_hours": _project_average(start_hours),
        "median_time_to_start_elapsed_hours": _project_median(start_hours),
        "time_to_start_elapsed_hours_valid_tasks": len(start_hours),
        "mean_time_to_start_business_hours": None,
        "median_time_to_start_business_hours": None,
        "time_to_start_business_hours_valid_tasks": 0,
    }


def _project_metric_row(item, summary_values: dict[str, Any]) -> list[Any]:
    return [summary_values.get(key) for key in (
        "total_tasks", "completed_tasks", "rejected_tasks", "open_tasks", "wip_tasks",
        "completion_rate", "rejection_rate", "on_time_tasks", "on_time_valid_tasks",
        "on_time_completion_rate", "overdue_open_tasks", "overdue_valid_tasks",
        "open_overdue_rate", "tasks_with_rework", "rework_valid_tasks",
        "tasks_with_rework_rate", "mean_execution_business_hours",
        "median_execution_business_hours", "mean_lead_time_business_hours",
        "median_lead_time_business_hours", "mean_rework_count", "total_rework_count",
        "history_complete_tasks", "unknown_status_tasks", "history_excluded_tasks",
        "reviewed_valid_tasks", "pending_before_execution", "rework_rate",
        "tasks_with_replanning", "total_replanning_count", "replanning_rate",
        "tasks_with_re_evaluation", "total_re_evaluation_count", "re_evaluation_rate",
        "review_exception_tasks", "review_exception_rate", "mean_execution_elapsed_hours",
        "median_execution_elapsed_hours", "execution_elapsed_hours_valid_tasks",
        "execution_business_hours_valid_tasks", "mean_lead_time_elapsed_hours",
        "median_lead_time_elapsed_hours", "lead_time_elapsed_hours_valid_tasks",
        "lead_time_business_hours_valid_tasks", "mean_time_to_start_elapsed_hours",
        "median_time_to_start_elapsed_hours", "time_to_start_elapsed_hours_valid_tasks",
        "mean_time_to_start_business_hours", "median_time_to_start_business_hours",
        "time_to_start_business_hours_valid_tasks",
    )]


def _project_task_metrics_rows(result: CompanyAnalysisResult) -> tuple[list[str], list[list[Any]]]:
    headers = [
        "issue_key", "task_name", "source_tool", "source_space", "issue_type", "labels",
        "assignee_id", "assignee_name", "priority", "status_at_cutoff",
        "source_snapshot_status", "status_known", "reached_review", "created_at",
        "actual_start_at", "completed_at", "due_date", "planned_start_date",
        "is_completed", "is_rejected", "is_open", "is_wip", "history_complete",
        "history_note", "execution_elapsed_hours", "execution_business_hours",
        "lead_time_elapsed_hours", "lead_time_business_hours", "time_to_start_elapsed_hours",
        "time_to_start_business_hours", "task_age_elapsed_hours", "task_age_business_hours",
        "current_status_age_elapsed_hours", "current_status_age_business_hours",
        "on_time_completion", "schedule_variance_days", "overdue_days",
        "start_schedule_variance_days", "rework_count", "replanning_count",
        "re_evaluation_count", "reopen_count", "time_in_status", "evaluation_period",
        "work_calendar_timezone", "work_calendar_days", "work_calendar_window",
        "data_quality_flags",
    ]
    rows: list[list[Any]] = []
    period_end = result.model.period_end
    for snapshot in result.snapshots:
        task = snapshot.task
        counts = _project_event_counts(snapshot)
        completed = snapshot.status_at_period_end.value == "Completed"
        rejected = snapshot.status_at_period_end.value == "Rejected"
        open_item = snapshot.status_at_period_end.is_open
        due = task.due_date
        late = completed and due and snapshot.final_completion_date and snapshot.final_completion_date > due
        overdue = open_item and due and due < period_end
        reached_review = any("review" in (interval.status.value or "").casefold() for interval in snapshot.status_intervals)
        intervals = "; ".join(
            f"{interval.status.value}: {interval.days * 24:.2f} elapsed hours"
            for interval in snapshot.status_intervals
        )
        start_variance = (
            (snapshot.actual_start_date - task.planned_start_date).days
            if snapshot.actual_start_date and task.planned_start_date else None
        )
        schedule_variance = (
            (snapshot.final_completion_date - due).days
            if snapshot.final_completion_date and due and completed else None
        )
        last_status_date = _project_last_status_date(snapshot)
        rows.append([
            task.task_id, task.task_name, task.source_tool, task.source_space or "N/A",
            "Subtask" if task.parent_id else "Task", "[]", "", task.assignee_group,
            task.priority or "N/A", snapshot.status_at_period_end.value,
            task.raw_status or "N/A", int(snapshot.status_at_period_end.value != "Unknown"),
            int(reached_review), task.created_date, snapshot.actual_start_date,
            snapshot.final_completion_date, due, task.planned_start_date, int(completed),
            int(rejected), int(open_item), int(snapshot.status_at_period_end.value in {"In Execution", "In Review"}),
            int(task.history_complete), "; ".join(sorted(snapshot.data_quality_flags)) or None,
            _project_elapsed_hours(snapshot.actual_start_date, snapshot.final_completion_date), None,
            _project_elapsed_hours(task.created_date, snapshot.final_completion_date), None,
            _project_elapsed_hours(task.created_date, snapshot.actual_start_date), None,
            _project_elapsed_hours(task.created_date, period_end), None,
            _project_elapsed_hours(last_status_date, period_end), None,
            None if late is None else int(late),
            schedule_variance, 0 if not overdue else max(0, (period_end - due).days),
            start_variance, counts["rework"], counts["replanning"], counts["re_evaluation"],
            counts["reopen"], intervals or None, _project_period_label(result),
            "Source configured timezone", "Unavailable", "Unavailable",
            "; ".join(sorted(snapshot.data_quality_flags)) or None,
        ])
    return headers, rows


def _project_detail_rows(result: CompanyAnalysisResult, snapshots=None) -> tuple[list[str], list[list[Any]]]:
    selected = list(result.snapshots if snapshots is None else snapshots)
    headers = [
        "issue_key", "task_name", "source_tool", "source_space", "assignee_name",
        "priority", "status_at_cutoff", "created_at", "planned_start_date",
        "actual_start_at", "start_schedule_variance_days", "due_date", "completed_at",
        "schedule_variance_days", "overdue_days", "task_age_elapsed_hours",
        "task_age_business_hours", "current_status_age_elapsed_hours",
        "current_status_age_business_hours",
    ]
    metrics_by_id = {row[0]: row for row in _project_task_metrics_rows(result)[1]}
    rows = []
    for snapshot in selected:
        task = snapshot.task
        metric = metrics_by_id.get(task.task_id, [])
        rows.append([
            task.task_id, task.task_name, task.source_tool, task.source_space or "N/A",
            task.assignee_group, task.priority or "N/A", snapshot.status_at_period_end.value,
            task.created_date, task.planned_start_date, snapshot.actual_start_date,
            metric[37] if len(metric) > 37 else None, task.due_date,
            snapshot.final_completion_date, metric[35] if len(metric) > 35 else None,
            0 if not (
                snapshot.status_at_period_end.is_open and task.due_date and task.due_date < snapshot.period_end
            ) else max(0, (snapshot.period_end - task.due_date).days),
            metric[30] if len(metric) > 30 else None, None,
            metric[32] if len(metric) > 32 else None, None,
        ])
    return headers, rows


def _project_events_rows(result: CompanyAnalysisResult) -> tuple[list[str], list[list[Any]]]:
    headers = [
        "issue_key", "task_name", "from_status", "to_status", "event_type",
        "changed_at", "author_name", "source", "source_space", "included_in_metrics",
    ]
    rows = []
    for snapshot in result.snapshots:
        for event in snapshot.task.workflow_history:
            if event.changed_at.date() > snapshot.period_end:
                continue
            before = (event.from_status or "").casefold()
            after = (event.to_status or "").casefold()
            event_type = "Normal"
            if "in review" in before and ("in progress" in after or "in execution" in after):
                event_type = "Rework"
            elif "in review" in before and "to do" in after:
                event_type = "Replanning"
            elif "in review" in before and "triage" in after:
                event_type = "Re-evaluation"
            elif "done" in after or "completed" in after:
                event_type = "Completion"
            elif "rejected" in after or "cancelled" in after:
                event_type = "Rejection"
            rows.append([
                snapshot.task.task_id, snapshot.task.task_name, event.from_status or "N/A",
                event.to_status or "N/A", event_type, event.changed_at, event.performed_by or "N/A",
                f"{snapshot.task.source_tool}_history", snapshot.task.source_space or "N/A",
                int(snapshot.history_available),
            ])
    return headers, rows


def _project_stage_rows(result: CompanyAnalysisResult) -> tuple[list[str], list[list[Any]]]:
    headers = [
        "status", "tasks_visited", "elapsed_total_hours", "elapsed_mean_hours",
        "elapsed_median_hours", "business_total_hours", "business_mean_hours",
        "business_median_hours", "open_tasks_currently_here",
    ]
    rows = []
    for metric in calculate_status_metrics(result.snapshots):
        intervals = [
            interval.days * 24.0
            for snapshot in result.snapshots
            if snapshot.counted_in_kpis and snapshot.history_available
            for interval in snapshot.status_intervals
            if interval.status is metric.status
        ]
        open_count = metric.open_tasks_now
        rows.append([
            metric.status.value, len(intervals), round(sum(intervals), 2),
            _project_average(intervals), _project_median(intervals),
            None, None, None, open_count,
        ])
    return headers, rows


def _project_deadline_rows(result: CompanyAnalysisResult) -> tuple[list[str], list[list[Any]]]:
    headers = ["due_status", "task_count", "share_of_open_known_tasks"]
    items = [item for item in result.snapshots if item.counted_in_kpis and item.status_at_period_end.is_open]
    with_due = [item for item in items if item.task.due_date]
    overdue = [item for item in with_due if item.task.due_date < item.period_end]
    within = [item for item in with_due if item.task.due_date >= item.period_end]
    without = [item for item in items if not item.task.due_date]
    denominator = len(with_due)
    share = lambda value: None if not denominator else round(value / denominator * 100.0, 1)
    return headers, [
        ["Open overdue", len(overdue), share(len(overdue))],
        ["Open within due date", len(within), share(len(within))],
        ["Open without due date", len(without), None if not items else round(len(without) / len(items) * 100.0, 1)],
    ]


def _project_weekly_rows(result: CompanyAnalysisResult) -> tuple[list[str], list[list[Any]]]:
    headers = ["week_start", "tasks_opened", "tasks_completed", "net_flow", "cumulative_net_flow"]
    values: dict[date, dict[str, int]] = defaultdict(lambda: {"tasks_opened": 0, "tasks_completed": 0})
    for snapshot in result.snapshots:
        if snapshot.task.created_date:
            week = snapshot.task.created_date - timedelta(days=snapshot.task.created_date.weekday())
            values[week]["tasks_opened"] += 1
        if snapshot.final_completion_date:
            week = snapshot.final_completion_date - timedelta(days=snapshot.final_completion_date.weekday())
            values[week]["tasks_completed"] += 1
    rows = []
    cumulative = 0
    for key, value in sorted(values.items()):
        net = value["tasks_opened"] - value["tasks_completed"]
        cumulative += net
        rows.append([key, value["tasks_opened"], value["tasks_completed"], net, cumulative])
    return headers, rows or [[result.model.period_start, 0, 0, 0, 0]]


def _project_findings_rows(result: CompanyAnalysisResult) -> tuple[list[str], list[list[Any]]]:
    headers = ["issue_key", "source_tool", "source_space", "observation", "follow_up"]
    rows = []
    for snapshot in result.snapshots:
        task = snapshot.task
        counts = _project_event_counts(snapshot)
        if task.due_date and snapshot.status_at_period_end.is_open and task.due_date < snapshot.period_end:
            days = (snapshot.period_end - task.due_date).days
            rows.append([
                task.task_id, task.source_tool, task.source_space or "N/A",
                f"Open and overdue by {days} calendar day(s) against the supplied due date.",
                "Confirm the due date, dependencies, and next action with the process owner.",
            ])
        for label, key in (
            ("rework", "rework"), ("replanning", "replanning"),
            ("re-evaluation", "re_evaluation"),
        ):
            if counts[key]:
                rows.append([
                    task.task_id, task.source_tool, task.source_space or "N/A",
                    f"{counts[key]} {label} return(s) before cutoff.",
                    "Review transition evidence with the process owner; the cause is not established by the count.",
                ])
    for metric in calculate_status_metrics(result.snapshots):
        if metric.tasks_passed_through:
            rows.append([
                f"Stage: {metric.status.value}", "Combined", "Selected Spaces",
                f"{metric.tasks_passed_through} task(s) visited; mean residence {metric.average_days or 0:.2f} days; "
                f"{metric.open_tasks_now} currently here.",
                "Inspect task-level stage times; residence is a review candidate, not proof of a bottleneck.",
            ])
    return headers, rows


def _project_summary_headers() -> list[str]:
    return [
        "total_tasks", "completed_tasks", "rejected_tasks", "open_tasks", "wip_tasks",
        "completion_rate", "rejection_rate", "on_time_tasks", "on_time_valid_tasks",
        "on_time_completion_rate", "overdue_open_tasks", "overdue_valid_tasks",
        "open_overdue_rate", "tasks_with_rework", "rework_valid_tasks",
        "tasks_with_rework_rate", "mean_execution_business_hours",
        "median_execution_business_hours", "mean_lead_time_business_hours",
        "median_lead_time_business_hours", "mean_rework_count", "total_rework_count",
        "history_complete_tasks", "unknown_status_tasks", "history_excluded_tasks",
        "reviewed_valid_tasks", "pending_before_execution", "rework_rate",
        "tasks_with_replanning", "total_replanning_count", "replanning_rate",
        "tasks_with_re_evaluation", "total_re_evaluation_count", "re_evaluation_rate",
        "review_exception_tasks", "review_exception_rate", "mean_execution_elapsed_hours",
        "median_execution_elapsed_hours", "execution_elapsed_hours_valid_tasks",
        "execution_business_hours_valid_tasks", "mean_lead_time_elapsed_hours",
        "median_lead_time_elapsed_hours", "lead_time_elapsed_hours_valid_tasks",
        "lead_time_business_hours_valid_tasks", "mean_time_to_start_elapsed_hours",
        "median_time_to_start_elapsed_hours", "time_to_start_elapsed_hours_valid_tasks",
        "mean_time_to_start_business_hours", "median_time_to_start_business_hours",
        "time_to_start_business_hours_valid_tasks",
    ]


def _project_assignee_rows(result: CompanyAnalysisResult) -> tuple[list[str], list[list[Any]]]:
    headers = ["assignee_name", *_project_summary_headers()]
    groups: dict[str, list[Any]] = defaultdict(list)
    for snapshot in result.snapshots:
        groups[snapshot.task.assignee_group].append(snapshot)
    return headers, [
        [name, *_project_metric_row(items, _project_summary_values(items))]
        for name, items in sorted(groups.items())
    ]


def _project_issue_type_rows(result: CompanyAnalysisResult) -> tuple[list[str], list[list[Any]]]:
    headers = ["issue_type", *_project_summary_headers()]
    groups: dict[str, list[Any]] = defaultdict(list)
    for snapshot in result.snapshots:
        groups["Subtask" if snapshot.task.parent_id else "Task"].append(snapshot)
    return headers, [
        [name, *_project_metric_row(items, _project_summary_values(items))]
        for name, items in sorted(groups.items())
    ]


def _dashboard_value(value: Any) -> Any:
    return "N/A" if value is None else value


def _dashboard_card(sheet, start_col: int, title: str, value: Any) -> None:
    end_col = start_col + 1
    sheet.merge_cells(start_row=3, start_column=start_col, end_row=3, end_column=end_col)
    sheet.merge_cells(start_row=4, start_column=start_col, end_row=5, end_column=end_col)
    header = sheet.cell(3, start_col, title)
    header.fill = _HEADER_FILL
    header.font = _HEADER_FONT
    header.alignment = Alignment(horizontal="center", vertical="center")
    value_cell = sheet.cell(4, start_col, _dashboard_value(value))
    value_cell.font = Font(size=16, bold=True, color="1F2937")
    value_cell.alignment = Alignment(horizontal="center", vertical="center")
    for row in range(3, 6):
        for col in range(start_col, end_col + 1):
            sheet.cell(row, col).border = sheet.cell(row, col).border.copy(
                left=sheet.cell(row, col).border.left,
                right=sheet.cell(row, col).border.right,
                top=sheet.cell(row, col).border.top,
                bottom=sheet.cell(row, col).border.bottom,
            )


def _add_project_dashboard_charts(sheet, weekly_rows, status_rows, due_rows, assignee_rows) -> None:
    weekly_start = 70
    _write_excel_table(sheet, ["Week Starting", "Tasks Opened", "Tasks Completed"], [
        [row[0], row[1], row[2]] for row in weekly_rows
    ], start_row=weekly_start)
    status_start = weekly_start + len(weekly_rows) + 3
    _write_excel_table(sheet, ["Status", "Tasks"], status_rows, start_row=status_start)
    due_start = status_start + len(status_rows) + 3
    _write_excel_table(sheet, ["Due Status", "Tasks"], due_rows, start_row=due_start)
    assignee_start = due_start + len(due_rows) + 3
    _write_excel_table(sheet, ["Assignee", "Completed", "Open", "Rejected"], assignee_rows, start_row=assignee_start)

    weekly_last = weekly_start + len(weekly_rows)
    line = LineChart()
    line.title = "Weekly Task Flow"
    line.style = 13
    line.height = 7.0
    line.width = 13.0
    line.add_data(Reference(sheet, min_col=2, max_col=3, min_row=weekly_start, max_row=weekly_last), titles_from_data=True)
    line.set_categories(Reference(sheet, min_col=1, min_row=weekly_start + 1, max_row=weekly_last))
    line.legend.position = "b"
    sheet.add_chart(line, "A13")

    status_last = status_start + len(status_rows)
    status_chart = BarChart()
    status_chart.type = "col"
    status_chart.style = 10
    status_chart.title = "Task Distribution by Status"
    status_chart.height = 7.0
    status_chart.width = 13.0
    status_chart.add_data(Reference(sheet, min_col=2, max_col=2, min_row=status_start, max_row=status_last), titles_from_data=True)
    status_chart.set_categories(Reference(sheet, min_col=1, min_row=status_start + 1, max_row=status_last))
    sheet.add_chart(status_chart, "I13")

    due_last = due_start + len(due_rows)
    due_chart = BarChart()
    due_chart.type = "col"
    due_chart.style = 10
    due_chart.title = "Open Tasks by Due Status"
    due_chart.height = 7.0
    due_chart.width = 13.0
    due_chart.add_data(Reference(sheet, min_col=2, max_col=2, min_row=due_start, max_row=due_last), titles_from_data=True)
    due_chart.set_categories(Reference(sheet, min_col=1, min_row=due_start + 1, max_row=due_last))
    sheet.add_chart(due_chart, "A29")

    assignee_last = assignee_start + len(assignee_rows)
    work_chart = BarChart()
    work_chart.type = "col"
    work_chart.grouping = "stacked"
    work_chart.overlap = 100
    work_chart.style = 10
    work_chart.title = "Work Distribution by Assignee"
    work_chart.height = 7.0
    work_chart.width = 13.0
    work_chart.add_data(Reference(sheet, min_col=2, max_col=4, min_row=assignee_start, max_row=assignee_last), titles_from_data=True)
    work_chart.set_categories(Reference(sheet, min_col=1, min_row=assignee_start + 1, max_row=assignee_last))
    work_chart.legend.position = "b"
    sheet.add_chart(work_chart, "I29")



def _project_management_averages(result: CompanyAnalysisResult) -> dict[str, float | None]:
    """Return management timing averages for the selected project scope."""
    items = [item for item in result.snapshots if item.counted_in_kpis]
    completed = [item for item in items if item.status_at_period_end.value == "Completed"]
    execution_hours = [
        value for value in (
            _project_elapsed_hours(item.actual_start_date, item.final_completion_date)
            for item in completed
        ) if value is not None
    ]
    lead_hours = [
        value for value in (
            _project_elapsed_hours(item.task.created_date, item.final_completion_date)
            for item in completed
        ) if value is not None
    ]
    start_hours = [
        value for value in (
            _project_elapsed_hours(item.task.created_date, item.actual_start_date)
            for item in items
        ) if value is not None
    ]
    late_completion_days = [
        float((item.final_completion_date - item.task.due_date).days)
        for item in completed
        if item.final_completion_date and item.task.due_date
        and item.final_completion_date > item.task.due_date
    ]
    open_overdue_days = [
        float((item.period_end - item.task.due_date).days)
        for item in items
        if item.status_at_period_end.is_open
        and item.task.due_date
        and item.task.due_date < item.period_end
    ]
    due_variance_days = [
        float((item.final_completion_date - item.task.due_date).days)
        for item in completed
        if item.final_completion_date and item.task.due_date
    ]
    return {
        "Avg Execution Time (Completed)": _project_average(execution_hours),
        "Avg Lead Time (Completed)": _project_average(lead_hours),
        "Avg Time to Start": _project_average(start_hours),
        "Avg Late Completion": _project_average(late_completion_days),
        "Avg Open Overdue": _project_average(open_overdue_days),
        "Avg Due Variance (Completed)": _project_average(due_variance_days),
    }


def _format_project_average(value: float | None, unit: str) -> str:
    return "Unavailable" if value is None else f"{value:.1f} {unit}"


def _project_excel_bytes(result: CompanyAnalysisResult, scope_label: str) -> bytes:
    """Build the exact Project Performance workbook requested by the user."""
    model = result.model
    values = _project_summary_values(result.snapshots)
    workbook = Workbook()
    workbook.remove(workbook.active)

    dashboard = workbook.create_sheet("Project_Executive_Dashboard")
    dashboard.sheet_view.showGridLines = False
    dashboard.merge_cells("A1:K1")
    dashboard["A1"] = "Project Performance"
    dashboard["A1"].fill = PatternFill("solid", fgColor="17324D")
    dashboard["A1"].font = Font(color="FFFFFF", bold=True, size=16)
    dashboard["A1"].alignment = Alignment(horizontal="center")
    dashboard.merge_cells("A2:K2")
    dashboard["A2"] = f"Project: {scope_label} | Analysis period: {_project_period_token(result)}"
    dashboard["A2"].font = Font(color="6B7280", italic=True)
    dashboard["A2"].alignment = Alignment(horizontal="center")
    for column in "ABCDEFGHIJK":
        dashboard.column_dimensions[column].width = 16

    cards = [
        ("Total Tasks", values["total_tasks"]),
        ("Completed", values["completed_tasks"]),
        ("Completion Rate", "N/A" if values["completion_rate"] is None else f'{values["completion_rate"]:.1f}%'),
        ("On-Time Rate", "N/A" if values["on_time_completion_rate"] is None else f'{values["on_time_completion_rate"]:.1f}%'),
        ("Open Overdue", values["overdue_open_tasks"]),
        ("WIP", values["wip_tasks"]),
        ("Average Lead Time", _format_project_average(_project_management_averages(result)["Avg Lead Time (Completed)"], "h")),
    ]
    for index, (title, value) in enumerate(cards):
        if index < 6:
            _dashboard_card(dashboard, 1 + index * 2, title, value)
        else:
            dashboard.merge_cells(start_row=7, start_column=1, end_row=7, end_column=2)
            dashboard.merge_cells(start_row=8, start_column=1, end_row=8, end_column=2)
            dashboard["A7"] = title
            dashboard["A7"].fill = _HEADER_FILL
            dashboard["A7"].font = _HEADER_FONT
            dashboard["A8"] = value

    weekly_headers, weekly_rows = _project_weekly_rows(result)
    status_counts = _project_status_counts(result.snapshots)
    status_rows = [[status, status_counts.get(status, 0)] for status in sorted(status_counts)] or [["No data", 0]]
    deadline_rows = _project_deadline_rows(result)[1]
    due_rows = [[row[0], row[1]] for row in deadline_rows] or [["No data", 0]]
    assignee_groups: dict[str, dict[str, int]] = defaultdict(lambda: {"Total Tasks": 0, "Completed": 0, "Open": 0})
    for item in result.snapshots:
        if not item.counted_in_kpis:
            continue
        group = assignee_groups[item.task.assignee_group]
        group["Total Tasks"] += 1
        if item.status_at_period_end.value == "Completed":
            group["Completed"] += 1
        elif item.status_at_period_end.is_open:
            group["Open"] += 1
    assignee_rows = [
        [name, data["Total Tasks"], data["Completed"], data["Open"]]
        for name, data in sorted(assignee_groups.items())
    ] or [["No data", 0, 0, 0]]

    weekly_start = 60
    _write_excel_table(dashboard, ["Week Starting", "Tasks Created", "Tasks Completed"], [
        [row[0], row[1], row[2]] for row in weekly_rows
    ] or [[model.period_start, 0, 0]], start_row=weekly_start)
    status_start = weekly_start + max(3, len(weekly_rows) + 3)
    _write_excel_table(dashboard, ["Status", "Tasks"], status_rows, start_row=status_start)
    due_start = status_start + len(status_rows) + 3
    _write_excel_table(dashboard, ["Due Status", "Tasks"], due_rows, start_row=due_start)
    assignee_start = due_start + len(due_rows) + 3
    _write_excel_table(dashboard, ["Assignee", "Total Tasks", "Completed", "Open"], assignee_rows, start_row=assignee_start)

    weekly_last = weekly_start + max(1, len(weekly_rows))
    line = LineChart()
    line.title = "Weekly Task Flow"
    line.style = 13
    line.height = 7
    line.width = 13
    line.add_data(Reference(dashboard, min_col=2, max_col=3, min_row=weekly_start, max_row=weekly_last), titles_from_data=True)
    line.set_categories(Reference(dashboard, min_col=1, min_row=weekly_start + 1, max_row=weekly_last))
    line.legend.position = "b"
    dashboard.add_chart(line, "A11")

    status_last = status_start + len(status_rows)
    status_chart = BarChart()
    status_chart.type = "col"
    status_chart.style = 10
    status_chart.title = "Task Distribution by Status"
    status_chart.height = 7
    status_chart.width = 13
    status_chart.add_data(Reference(dashboard, min_col=2, max_col=2, min_row=status_start, max_row=status_last), titles_from_data=True)
    status_chart.set_categories(Reference(dashboard, min_col=1, min_row=status_start + 1, max_row=status_last))
    dashboard.add_chart(status_chart, "I11")

    due_last = due_start + len(due_rows)
    due_chart = BarChart()
    due_chart.type = "col"
    due_chart.style = 10
    due_chart.title = "Open Tasks by Due Status"
    due_chart.height = 7
    due_chart.width = 13
    due_chart.add_data(Reference(dashboard, min_col=2, max_col=2, min_row=due_start, max_row=due_last), titles_from_data=True)
    due_chart.set_categories(Reference(dashboard, min_col=1, min_row=due_start + 1, max_row=due_last))
    dashboard.add_chart(due_chart, "A27")

    assignee_last = assignee_start + len(assignee_rows)
    work_chart = BarChart()
    work_chart.type = "col"
    work_chart.grouping = "stacked"
    work_chart.overlap = 100
    work_chart.style = 10
    work_chart.title = "Work Distribution by Assignee"
    work_chart.height = 7
    work_chart.width = 13
    work_chart.add_data(Reference(dashboard, min_col=2, max_col=4, min_row=assignee_start, max_row=assignee_last), titles_from_data=True)
    work_chart.set_categories(Reference(dashboard, min_col=1, min_row=assignee_start + 1, max_row=assignee_last))
    work_chart.legend.position = "b"
    dashboard.add_chart(work_chart, "I27")
    _fit_sheet(dashboard)

    source_spaces = " / ".join(
        f"{item.source_tool}: {item.source_space or 'N/A'}"
        for item in model.source_coverage
    ) or "N/A"
    summary = pd.DataFrame([{
        "Project": scope_label,
        "Source Spaces": source_spaces,
        "Analysis Period": _project_period_label(result),
        "Total Tasks": values["total_tasks"],
        "Completed": values["completed_tasks"],
        "Completion Rate": values["completion_rate"],
        "On-Time Rate": values["on_time_completion_rate"],
        "Open Overdue": values["overdue_open_tasks"],
        "WIP": values["wip_tasks"],
        "Average Lead Time (hours)": _project_management_averages(result)["Avg Lead Time (Completed)"],
    }])
    status_detail_headers, status_detail_rows = _project_detail_rows(result)
    task_headers, task_rows = _project_task_metrics_rows(result)
    assignee_headers, assignee_rows_detail = _project_assignee_rows(result)
    definitions = [
        ["Total Tasks", "Count of all tasks in the selected Project scope; subtasks remain visible."],
        ["Completed", "Tasks with normalized Completed status at period end."],
        ["Completion Rate", "Completed tasks divided by KPI-counted tasks with a verified status at period end; Unknown statuses are excluded and shown in Data Quality."],
        ["On-Time Rate", "Completed tasks on or before due date divided by completed tasks with a known due date."],
        ["Open Overdue", "Open tasks with a due date earlier than the analysis period end."],
        ["WIP", "Open tasks currently in In Execution or In Review at period end; On Hold and At Risk are not WIP."],
        ["Average Lead Time", "Average completion date minus creation date for completed KPI-counted tasks."],
        ["Due Variance", "Completion date minus due date; positive values indicate late completion."],
        ["Workflow Exceptions", "Recorded rework, replanning, re-evaluation, or reopen evidence."],
    ]
    stage_headers, stage_rows = _project_stage_rows(result)
    bottleneck_rows = [
        [item.status.value, item.strength, "; ".join(item.evidence),
         item.metrics.average_days, item.metrics.open_tasks_now, item.metrics.overdue_open_tasks]
        for item in result.bottlenecks
    ] or [["No candidate", "N/A", "No bottleneck candidate was identified.", None, 0, 0]]
    finding_headers, finding_rows = _project_findings_rows(result)
    overdue = [
        item for item in result.snapshots
        if item.counted_in_kpis and item.status_at_period_end.is_open
        and item.task.due_date and item.task.due_date < item.period_end
    ]
    late = [
        item for item in result.snapshots
        if item.counted_in_kpis and item.status_at_period_end.value == "Completed"
        and item.task.due_date and item.final_completion_date
        and item.final_completion_date > item.task.due_date
    ]
    completed_with_due = [
        item for item in result.snapshots
        if item.counted_in_kpis and item.status_at_period_end.value == "Completed"
        and item.task.due_date and item.final_completion_date
    ]
    due_variance_rows = [
        ["Completed On Time", sum(item.final_completion_date <= item.task.due_date for item in completed_with_due)],
        ["Late Completed", len(late)],
        ["Open Overdue", len(overdue)],
        ["Open Without Due Date", sum(
            item.counted_in_kpis and item.status_at_period_end.is_open and not item.task.due_date
            for item in result.snapshots
        )],
    ]
    exception_rows = [
        [item.task.task_id, item.task.task_name, item.task.source_tool, item.task.source_space or "N/A",
         exception]
        for item in result.snapshots for exception in item.exception_events
    ] or [["No workflow exception", "", "", "", ""]]
    quality_rows = [[item.flag, item.task_count] for item in model.data_quality] or [["No findings", 0]]
    context_rows = [
        ["Project", scope_label],
        ["Source Spaces", source_spaces],
        ["Analysis Period", _project_period_label(result)],
        ["Task Scope", "Selected Spaces unified into one Project scope; duplicate Source Tool + Task ID records removed."],
        ["Subtask Treatment", "Visible in task-level data; excluded from KPI rates."],
        ["Status Treatment", "Original source status is retained and normalized status is used for cross-source metrics."],
    ]
    sheet_specs = [
        ("Project Summary", list(summary.columns), summary.astype(object).where(pd.notna(summary), None).values.tolist()),
        ("Task Evaluation", task_headers, task_rows),
        ("Assignee Summary", assignee_headers, assignee_rows_detail),
        ("Status Summary", ["Status", "Task Count"], status_rows),
        ("Status Detail", status_detail_headers, status_detail_rows),
        ("Weekly Flow", weekly_headers, weekly_rows),
        ("Due Variance", ["Due Variance Category", "Task Count"], due_variance_rows),
        ("Overdue Tasks", status_detail_headers, _project_detail_rows(result, overdue)[1]),
        ("Late Completed Tasks", status_detail_headers, _project_detail_rows(result, late)[1]),
        ("Bottlenecks", ["Stage", "Assessment", "Evidence", "Average Days", "Open Tasks", "Open Overdue"], bottleneck_rows),
        ("Findings", finding_headers, finding_rows),
        ("Data Quality", ["Data Quality Flag", "Task Count"], quality_rows),
        ("Analysis Context", ["Field", "Value"], context_rows),
        ("Metric Definitions", ["Metric", "Definition"], definitions),
    ]
    for name, headers, rows in sheet_specs:
        sheet = workbook.create_sheet(name)
        _write_excel_table(sheet, headers, rows)
        _fit_sheet(sheet)

    for sheet in workbook.worksheets:
        sheet.sheet_view.showGridLines = False
    stream = BytesIO()
    workbook.save(stream)
    return stream.getvalue()

def _docx_value(value):
    if value is None or value == "":
        return "Unavailable"
    if isinstance(value, (tuple, list, set)):
        return "; ".join(str(item) for item in value) if value else "Unavailable"
    if isinstance(value, date):
        return value.isoformat()
    return str(value)



def _add_docx_table(document, headers, values):
    return add_report_table(document, headers, values)


def _project_word_task_pairs(snapshot) -> list[tuple[str, Any]]:
    task = snapshot.task
    counts = _project_event_counts(snapshot)
    completed = snapshot.status_at_period_end.value == "Completed"
    late = completed and task.due_date and snapshot.final_completion_date and snapshot.final_completion_date > task.due_date
    overdue_days = (
        max(0, (snapshot.period_end - task.due_date).days)
        if snapshot.status_at_period_end.is_open and task.due_date and task.due_date < snapshot.period_end
        else 0
    )
    execution_hours = _project_elapsed_hours(snapshot.actual_start_date, snapshot.final_completion_date)
    open_age_hours = _project_elapsed_hours(task.created_date, snapshot.period_end) if snapshot.status_at_period_end.is_open else None
    current_age_hours = _project_elapsed_hours(_project_last_status_date(snapshot), snapshot.period_end) if snapshot.status_at_period_end.is_open else None
    return [
        ("Source / Space", f"{task.source_tool} / {task.source_space or 'Unavailable'}"),
        ("Status at cutoff", snapshot.status_at_period_end.value),
        ("Actual start / completion", f"{snapshot.actual_start_date or 'Unavailable'} / {snapshot.final_completion_date or 'Unavailable'}"),
        ("Due date", task.due_date or "Unavailable"),
        ("On-time completion", "No" if late else ("Yes" if completed and task.due_date and snapshot.final_completion_date else "Unavailable")),
        ("Execution elapsed / business hours", f"{execution_hours:.2f} / Unavailable" if execution_hours is not None else "Unavailable / Unavailable"),
        ("Open task age elapsed / business hours", f"{open_age_hours:.2f} / Unavailable" if open_age_hours is not None else "Unavailable / Unavailable"),
        ("Current status age elapsed / business hours", f"{current_age_hours:.2f} / Unavailable" if current_age_hours is not None else "Unavailable / Unavailable"),
        ("Overdue days", overdue_days),
        ("Rework / replanning / re-evaluation events", f"{counts['rework']} / {counts['replanning']} / {counts['re_evaluation']}"),
        ("History complete", "Yes" if task.history_complete else "No"),
        ("Data quality flags", "; ".join(sorted(snapshot.data_quality_flags)) or "None"),
    ]


def _project_stage_word_rows(result: CompanyAnalysisResult) -> list[tuple[str, Any]]:
    rows = []
    for metric in calculate_status_metrics(result.snapshots):
        intervals = [
            interval.days * 24.0
            for snapshot in result.snapshots
            if snapshot.counted_in_kpis and snapshot.history_available
            for interval in snapshot.status_intervals
            if interval.status is metric.status
        ]
        rows.append((
            "Tasks visited", len(intervals),
        ))
        rows.append((
            "Mean elapsed / business hours",
            f"{_project_average(intervals) or 0:.2f} / Unavailable",
        ))
        rows.append((
            "Median elapsed / business hours",
            f"{_project_median(intervals) or 0:.2f} / Unavailable",
        ))
        rows.append(("Total elapsed / business hours", f"{sum(intervals):.2f} / Unavailable"))
        rows.append(("Open tasks currently here", metric.open_tasks_now))
    return rows


def _project_word_bytes(result: CompanyAnalysisResult, scope_label: str) -> bytes:
    """Build the exact Project Performance Word report requested by the user."""
    model = result.model
    values = _project_summary_values(result.snapshots)
    period = _project_period_label(result)
    document = Document()
    configure_report_document(
        document,
        header_label="PROJECT PERFORMANCE REPORT | PERFORMANCE EVALUATION",
    )
    add_report_title(
        document,
        "Project Performance Report",
        scope_label,
        metadata=(
            f"Source Spaces: {' / '.join(item.source_tool + ': ' + (item.source_space or 'N/A') for item in model.source_coverage) or 'N/A'}",
            f"Analysis period: {period}",
        ),
    )

    document.add_heading("Executive Summary", level=1)
    completion = "N/A" if values["completion_rate"] is None else f'{values["completion_rate"]:.1f}%'
    on_time = "N/A" if values["on_time_completion_rate"] is None else f'{values["on_time_completion_rate"]:.1f}%'
    document.add_paragraph(
        f"The selected Spaces were unified as one Project scope for {period}. "
        f"The scope contains {values['total_tasks']} task(s), including visible subtasks, "
        f"with {values['completed_tasks']} completed. Completion rate is {completion}; "
        f"on-time rate is {on_time}; {values['overdue_open_tasks']} open overdue task(s) "
        "require follow-up."
    )

    document.add_heading("Project Scope and Context", level=1)
    _add_docx_table(document, ["Field", "Value"], [
        ("Project", scope_label),
        ("Source", "Jira and/or ClickUp selected Spaces"),
        ("Source Spaces", " / ".join(
            f"{item.source_tool}: {item.source_space or 'N/A'}"
            for item in model.source_coverage
        ) or "N/A"),
        ("Analysis Period", period),
        ("Scope Rule", "Selected Spaces are unified into one Project scope and duplicate Source Tool + Task ID records are removed."),
        ("Subtask Rule", "Subtasks remain visible in task-level output but are excluded from KPI rates."),
    ])

    document.add_heading("Project Indicators", level=1)
    _add_docx_table(document, ["Indicator", "Value"], [
        ("Total Tasks", values["total_tasks"]),
        ("Completed", values["completed_tasks"]),
        ("Completion Rate", completion),
        ("On-Time Rate", on_time),
        ("Open Overdue", values["overdue_open_tasks"]),
        ("WIP", values["wip_tasks"]),
        ("Average Lead Time", _format_project_average(_project_management_averages(result)["Avg Lead Time (Completed)"], "hours")),
    ])

    document.add_heading("Task Distribution", level=1)
    status_counts = _project_status_counts(result.snapshots)
    _add_docx_table(document, ["Status", "Task Count"], [
        [status, status_counts.get(status, 0)] for status in sorted(status_counts)
    ] or [["N/A", 0]])
    priority_counts = Counter(
        (item.task.priority or "Unknown")
        for item in result.snapshots if item.counted_in_kpis
    )
    document.add_paragraph("Priority")
    _add_docx_table(document, ["Priority", "Task Count"], sorted(priority_counts.items()) or [["N/A", 0]])
    type_counts = Counter(
        "Subtask" if item.task.parent_id else "Task"
        for item in result.snapshots if item.in_scope
    )
    document.add_paragraph("Task Type")
    _add_docx_table(document, ["Task Type", "Task Count"], sorted(type_counts.items()) or [["N/A", 0]])

    document.add_heading("Weekly Progress", level=1)
    weekly_headers, weekly_rows = _project_weekly_rows(result)
    _add_docx_table(document, ["Week Starting", "Tasks Opened", "Tasks Completed", "Net Flow", "Cumulative Net Flow"], weekly_rows)

    document.add_heading("Bottlenecks by Project Stage", level=1)
    bottleneck_rows = [
        (item.status.value, item.strength, "; ".join(item.evidence),
         item.metrics.average_days, item.metrics.open_tasks_now, item.metrics.overdue_open_tasks)
        for item in result.bottlenecks
    ]
    _add_docx_table(document, ["Stage", "Assessment", "Evidence", "Average Days", "Open Tasks", "Open Overdue"], bottleneck_rows or [
        ("N/A", "No candidate", "No bottleneck candidate was identified.", "N/A", 0, 0)
    ])

    overdue = [
        item for item in result.snapshots
        if item.counted_in_kpis and item.status_at_period_end.is_open
        and item.task.due_date and item.task.due_date < item.period_end
    ]
    document.add_heading("Open Overdue Tasks", level=1)
    _add_docx_table(document, ["Task ID", "Task", "Source", "Space", "Assignee", "Due Date", "Overdue Days"], [
        (
            item.task.task_id, item.task.task_name or "N/A", item.task.source_tool,
            item.task.source_space or "N/A", item.task.assignee_group, item.task.due_date,
            (item.period_end - item.task.due_date).days,
        )
        for item in overdue
    ] or [("N/A", "No open overdue tasks", "", "", "", "", 0)])

    late = [
        item for item in result.snapshots
        if item.counted_in_kpis and item.status_at_period_end.value == "Completed"
        and item.task.due_date and item.final_completion_date
        and item.final_completion_date > item.task.due_date
    ]
    document.add_heading("Late Completed Tasks", level=1)
    _add_docx_table(document, ["Task ID", "Task", "Source", "Space", "Assignee", "Due Date", "Completed", "Variance (days)"], [
        (
            item.task.task_id, item.task.task_name or "N/A", item.task.source_tool,
            item.task.source_space or "N/A", item.task.assignee_group, item.task.due_date,
            item.final_completion_date, (item.final_completion_date - item.task.due_date).days,
        )
        for item in late
    ] or [("N/A", "No late completed tasks", "", "", "", "", "", 0)])

    document.add_heading("Workload by Assignee", level=1)
    assignee_headers, assignee_rows = _project_assignee_rows(result)
    simple_assignees = [
        (row[0], row[1], row[2], row[4], row[6])
        for row in assignee_rows
    ]
    _add_docx_table(document, ["Assignee", "Total Tasks", "Completed", "Open", "Completion Rate"], simple_assignees or [
        ("N/A", 0, 0, 0, "N/A")
    ])

    document.add_heading("Workflow Exceptions", level=1)
    exception_rows = [
        (item.task.task_id, item.task.task_name or "N/A", item.task.source_tool,
         item.task.source_space or "N/A", exception)
        for item in result.snapshots for exception in item.exception_events
    ]
    _add_docx_table(document, ["Task ID", "Task", "Source", "Space", "Exception"], exception_rows or [
        ("N/A", "No workflow exceptions", "", "", "")
    ])

    document.add_heading("Recommendations", level=1)
    _add_docx_table(document, ["Severity", "Recommendation", "Evidence", "Suggested Action"], [
        (item.severity, item.title, item.evidence, item.suggested_action)
        for item in result.recommendations
    ] or [("N/A", "No recommendations generated", "", "")])

    document.add_heading("Task-Level Evaluation", level=1)
    for snapshot in result.snapshots:
        document.add_heading(
            f"{snapshot.task.task_id} - {snapshot.task.task_name or 'Unnamed task'}",
            level=2,
        )
        _add_docx_table(document, ["Metric", "Value"], _project_word_task_pairs(snapshot))

    document.add_heading("Data Quality", level=1)
    _add_docx_table(document, ["Data Quality Flag", "Task Count"], [
        (item.flag, item.task_count) for item in model.data_quality
    ] or [("No findings", 0)])
    document.add_paragraph(
        "Unavailable values remain N/A. Source and Space identifiers are preserved, and "
        "cross-source totals use the same deduplicated Project dataset shown in Excel."
    )

    document.add_heading("Metric Definitions", level=1)
    _add_docx_table(document, ["Metric", "Definition"], [
        ("Total Tasks", "All tasks in the selected Project scope; subtasks remain visible."),
        ("Completed", "Tasks with normalized Completed status at period end."),
        ("Completion Rate", "Completed tasks divided by KPI-counted tasks."),
        ("On-Time Rate", "Completed on or before due date divided by completed tasks with a known due date."),
        ("Open Overdue", "Open tasks with a due date earlier than period end."),
        ("WIP", "Open tasks in In Execution or In Review at period end."),
        ("Average Lead Time", "Average completion date minus creation date for completed KPI-counted tasks."),
        ("Workflow Exceptions", "Recorded rework, replanning, re-evaluation, or reopen evidence."),
    ])

    stream = BytesIO()
    document.save(stream)
    return stream.getvalue()

def _render_project_result(st: Any, result: CompanyAnalysisResult, scope_label: str, scope_slug: str) -> None:
    if st.button("Start New Project Analysis", key="project_new_analysis"):
        _clear_project_run()
        st.rerun()

    st.divider()
    st.title("Project Performance Analysis")
    st.caption(f"{scope_label} · {result.model.period_start.isoformat()} to {result.model.period_end.isoformat()}")

    cards_by_key = {card.key: card for card in result.model.cards}
    population = population_from_snapshots(result.snapshots)
    headline = st.columns(3)
    headline[0].metric(
        cards_by_key["total-tasks"].title,
        cards_by_key["total-tasks"].value,
        help=cards_by_key["total-tasks"].supporting_text,
    )
    render_population_card(headline[1], population)
    headline[2].metric(
        cards_by_key["completion-rate"].title,
        cards_by_key["completion-rate"].value,
        help=cards_by_key["completion-rate"].supporting_text,
    )
    secondary = st.columns(4)
    for column, key in zip(
        secondary,
        ("total-projects", "current-wip", "overdue-open", "on-time-rate"),
    ):
        card = cards_by_key[key]
        column.metric(card.title, card.value, help=card.supporting_text)

    st.subheader("Management Averages")
    st.caption(
        "Execution, lead-time, and time-to-start averages are measured in elapsed hours. "
        "Due variance and overdue measures use calendar days."
    )
    management_averages = _project_management_averages(result)
    management_cards = [
        ("Avg Execution Time (Completed)", management_averages["Avg Execution Time (Completed)"], "h"),
        ("Avg Lead Time (Completed)", management_averages["Avg Lead Time (Completed)"], "h"),
        ("Avg Time to Start", management_averages["Avg Time to Start"], "h"),
        ("Avg Late Completion", management_averages["Avg Late Completion"], "days"),
        ("Avg Open Overdue", management_averages["Avg Open Overdue"], "days"),
        ("Avg Due Variance (Completed)", management_averages["Avg Due Variance (Completed)"], "days"),
    ]
    for start in range(0, len(management_cards), 3):
        columns = st.columns(3)
        for column, (title, value, unit) in zip(columns, management_cards[start:start + 3]):
            column.metric(title, _format_project_average(value, unit))

    dashboard, process, achievements, details, quality = st.tabs([
        "Executive Dashboard",
        "Process Analysis",
        "Individual Achievements",
        "Task Details",
        "Data Quality",
    ])

    with dashboard:
        for chart in result.model.executive_charts:
            st.subheader(chart.title)
            if chart.points:
                st.bar_chart({point.label: point.value for point in chart.points}, use_container_width=True)
            if chart.note:
                st.caption(chart.note)
            else:
                st.info("No eligible data is available for this chart.")

        st.subheader("Weekly Task Flow")
        weekly = _weekly_flow_frame(result)
        if weekly.empty:
            st.info("No created or completed dates are available.")
        else:
            st.line_chart(weekly.set_index("Week Starting")[["Tasks Created", "Tasks Completed"]], use_container_width=True)

        st.subheader("Source and Space Contribution")
        st.dataframe(_source_space_frame(result), hide_index=True, use_container_width=True)

    with process:
        st.subheader("Assignee Workload")
        st.dataframe(_assignee_frame(result), hide_index=True, use_container_width=True)
        st.subheader("Bottleneck Candidates")
        st.dataframe([
            {
                "Status": item.status.value,
                "Assessment": item.strength,
                "Evidence": item.evidence,
                "Average Days": item.metrics.average_days,
                "Open Tasks": item.metrics.open_tasks_now,
                "Open Overdue": item.metrics.overdue_open_tasks,
            }
            for item in result.bottlenecks
        ], hide_index=True, use_container_width=True)
        st.subheader("Recommendations")
        st.dataframe([
            {
                "Severity": item.severity,
                "Recommendation": item.title,
                "Evidence": item.evidence,
                "Suggested Action": item.suggested_action,
            }
            for item in result.recommendations
        ], hide_index=True, use_container_width=True)

    with achievements:
        st.dataframe(_assignee_frame(result), hide_index=True, use_container_width=True)

    with details:
        st.dataframe(_task_detail_frame(result), hide_index=True, use_container_width=True)

    with quality:
        st.subheader("Source Coverage")
        st.dataframe([
            {
                "Source": item.source_tool,
                "Space": item.source_space,
                "Tasks": item.task_count,
                "History Coverage": item.history_mode,
                "Notes": item.reason or "N/A",
            }
            for item in result.model.source_coverage
        ], hide_index=True, use_container_width=True)
        st.subheader("Data Quality")
        st.dataframe([
            {"Flag": item.flag, "Task Count": item.task_count}
            for item in result.model.data_quality
        ], hide_index=True, use_container_width=True)

    report_key = (
        f"{scope_slug}:{result.model.period_start.isoformat()}:{result.model.period_end.isoformat()}:"
        f"{len(result.snapshots)}"
    )
    if st.session_state.get("project_report_key") != report_key:
        st.session_state["project_excel_bytes"] = _project_excel_bytes(result, scope_label)
        st.session_state["project_word_bytes"] = _project_word_bytes(result, scope_label)
        st.session_state["project_report_key"] = report_key

    period_token = (
        f"{result.model.period_start.isoformat()}_to_{result.model.period_end.isoformat()}"
    )
    excel_name = f"Project_Performance_Analysis_Workbook_{scope_slug}_{period_token}.xlsx"
    word_name = f"Project_Performance_Analysis_Report_{scope_slug}_{period_token}.docx"
    st.subheader("Downloads")
    left, right = st.columns(2)
    left.download_button("Download Project Excel Workbook", st.session_state["project_excel_bytes"], excel_name, _XLSX_MIME, use_container_width=True)
    right.download_button("Download Project Word Report", st.session_state["project_word_bytes"], word_name, _DOCX_MIME, use_container_width=True)


@st.fragment
def _render_project_period_controls(settings, scope_label: str, scope_slug: str, preview) -> None:
    """Render period controls in an isolated fragment.

    Date changes rerun only this control block, so Streamlit does not rebuild
    the full page and move the user's viewport back to the top.
    """
    period_left, period_right = st.columns(2)
    period_start = period_left.date_input(
        "From Date",
        value=None,
        key="project_period_start",
        on_change=_clear_project_analysis,
    )
    period_end = period_right.date_input(
        "To Date",
        value=None,
        key="project_period_end",
        on_change=_clear_project_analysis,
    )
    valid_dates = isinstance(period_start, date) and isinstance(period_end, date)
    if valid_dates and period_end < period_start:
        st.error("To Date must be on or after From Date.")
        valid_dates = False

    run_clicked = st.button(
        "Run Project Analysis",
        type="primary",
        disabled=not valid_dates or not preview,
        key="run_project_analysis",
    )
    if run_clicked:
        try:
            with st.spinner("Calculating Project Performance Analysis..."):
                result = build_company_analysis(
                    period_start=period_start,
                    period_end=period_end,
                    jira_prepared=st.session_state.get("project_jira_prepared"),
                    clickup_prepared=st.session_state.get("project_clickup_prepared"),
                    unified_project=scope_label,
                )
            st.session_state["project_analysis"] = result
            st.session_state["project_scope_label"] = scope_label
            st.session_state["project_scope_slug"] = scope_slug
            st.session_state.pop("project_report_key", None)
            st.rerun()
        except ValueError as exc:
            st.error(str(exc))


def render_project_collection(settings):
    jira_spaces, clickup_spaces = _load_catalogs(settings)
    jira_error = st.session_state.get("project_jira_error", "")
    clickup_error = st.session_state.get("project_clickup_error", "")

    st.caption("Select one Jira Space, one ClickUp Space, or one from each to analyze them as a single project scope.")
    if jira_error:
        st.warning(f"Jira Spaces are unavailable: {jira_error}")
    if clickup_error:
        st.warning(f"ClickUp Spaces are unavailable: {clickup_error}")
    if not jira_spaces and not clickup_spaces:
        st.error("No Jira or ClickUp Spaces are available for this connection.")
        return None, False

    left, right = st.columns(2)
    selected_jira_id = left.selectbox(
        "Jira Space",
        [None, *jira_spaces],
        format_func=lambda value: "No Jira Space selected" if value is None else (
            f"{jira_spaces[value].get('name') or jira_spaces[value].get('key') or value} "
            f"({jira_spaces[value].get('key') or value})"
        ),
        key="project_selected_jira_space",
        on_change=_clear_project_run,
    )
    selected_clickup_id = right.selectbox(
        "ClickUp Space",
        [None, *clickup_spaces],
        format_func=lambda value: "No ClickUp Space selected" if value is None else (
            f"{clickup_spaces[value].get('name') or value}"
        ),
        key="project_selected_clickup_space",
        on_change=_clear_project_run,
    )

    jira_item = jira_spaces.get(selected_jira_id) if selected_jira_id else None
    clickup_item = clickup_spaces.get(selected_clickup_id) if selected_clickup_id else None
    if not jira_item and not clickup_item:
        st.info("Select at least one Jira or ClickUp Space to load project tasks.")
        return None, False

    selection_fingerprint = f"{selected_jira_id or '-'}:{selected_clickup_id or '-'}"
    selection_changed = (
        st.session_state.get("project_selection_fingerprint") != selection_fingerprint
    )
    if selection_changed:
        retrying = bool(st.session_state.pop("project_collection_retry", False))
        retry_id = st.session_state.get("project_collection_id") if retrying else None
        _clear_project_run()
        st.session_state["project_selection_fingerprint"] = selection_fingerprint
        st.session_state["project_collection_id"] = retry_id or uuid4().hex
        st.session_state["project_collection_fresh"] = not bool(retry_id)
        try:
            # Use one shared line even when the project combines Jira and
            # ClickUp.  Source callbacks update their assigned segment of the
            # same bar instead of leaving one line per source/page behind.
            source_count = int(bool(jira_item)) + int(bool(clickup_item))
            collection_progress = st.progress(
                0.0, text=f"Collecting project sources: 0 / {source_count}"
            )
            source_index = 0

            def source_progress(fraction: float, message: str, position: int) -> None:
                start = position / source_count if source_count else 0.0
                end = (position + 1) / source_count if source_count else 1.0
                collection_progress.progress(
                    start + (end - start) * max(0.0, min(1.0, fraction)),
                    text=str(message),
                )

            if jira_item:
                st.session_state["project_jira_prepared"] = _collect_jira_space(
                    settings,
                    jira_item,
                    f"project:jira:{selection_fingerprint}",
                    progress=lambda fraction, message, position=source_index: source_progress(
                        fraction, message, position
                    ),
                    collection_id=st.session_state["project_collection_id"],
                    fresh=st.session_state.get("project_collection_fresh", False),
                )
                source_index += 1
            if clickup_item:
                st.session_state["project_clickup_prepared"] = _collect_clickup_space(
                    settings,
                    clickup_item,
                    f"project:clickup:{selection_fingerprint}",
                    progress=lambda fraction, message, position=source_index: source_progress(
                        fraction, message, position
                    ),
                    collection_id=st.session_state["project_collection_id"],
                    fresh=st.session_state.get("project_collection_fresh", False),
                )
                source_index += 1
            collection_progress.progress(
                1.0, text=f"Project collection complete: {source_index} / {source_count} sources"
            )
            preview = build_company_preview(
                jira_prepared=st.session_state.get("project_jira_prepared"),
                clickup_prepared=st.session_state.get("project_clickup_prepared"),
            )
            st.session_state["project_preview"] = preview
            st.session_state["project_collection_fresh"] = False
            st.session_state["project_collection_collected_at"] = datetime.now(timezone.utc).isoformat()
        except (CollectionError, ClickUpCollectionError) as exc:
            st.session_state["project_collection_retry"] = True
            st.session_state.pop("project_selection_fingerprint", None)
            st.session_state.pop("project_jira_prepared", None)
            st.session_state.pop("project_clickup_prepared", None)
            st.error(str(exc))
            return None, False

    preview = st.session_state.get("project_preview", ())
    if st.session_state.get("project_collection_collected_at"):
        st.caption(f"Data collected at: {st.session_state['project_collection_collected_at']}")
    st.subheader("Selected Project Tasks")
    st.caption(f"{len(preview)} task(s) collected from the selected Space(s).")
    st.dataframe(_preview_frame(preview), hide_index=True, use_container_width=True)

    scope_label, scope_slug = _scope_label(jira_item, clickup_item)
    _render_project_period_controls(settings, scope_label, scope_slug, preview)

    if st.session_state.get("project_analysis") is not None:
        return None, False
    return None, False

def render_project_result(st_instance: Any, result: CompanyAnalysisResult) -> None:
    scope_label = st_instance.session_state.get("project_scope_label", "Selected Project Scope")
    scope_slug = st_instance.session_state.get("project_scope_slug", _safe_filename(scope_label))
    _render_project_result(st_instance, result, scope_label, scope_slug)
