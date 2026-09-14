"""Project Performance: dual-source space selection and combined analysis."""
from __future__ import annotations

from collections import Counter, defaultdict
from datetime import date, datetime, timedelta, timezone
from io import BytesIO
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from clickup_export import collect_data as collect_clickup_data
from clickup_gateway import ClickUpCollectionError, ClickUpGateway
from company_performance.application import CompanyAnalysisResult, build_company_analysis, build_company_preview
from company_performance.kpis import calculate_status_metrics
from jira_export import PreparedData, _json, build_workbook
from jira_gateway import CollectionError, JiraGateway


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


def _collect_jira_space(settings, item: dict[str, Any], fingerprint: str) -> PreparedData:
    project_key = str(item.get("key") or "")
    project_name = str(item.get("name") or project_key or item.get("id") or "Jira Project")
    if not project_key:
        raise CollectionError("The selected Jira Space has no usable project key.")

    query = f'project = "{project_key}" ORDER BY created DESC'
    gateway = JiraGateway(settings)
    try:
        issues = gateway.all_issues(query, progress=lambda message: st.caption(message))
        definitions = gateway.fields()
        histories: dict[str, dict[str, Any]] = {}
        complete: list[dict[str, Any]] = []
        for index, issue in enumerate(issues, 1):
            with st.status(
                f"Reading Jira task {index} of {len(issues)}...",
                expanded=False,
            ) as status:
                current, history = gateway.complete_issue(issue)
                if history.get("history_complete") is not True or not history.get("history_through"):
                    raise CollectionError(
                        "A Jira task history is incomplete. No partial project export was prepared."
                    )
                complete.append(current)
                histories[current["key"]] = history
                status.update(label=f"Read {current['key']}", state="complete")

        cutoff = datetime.now(timezone.utc).isoformat()
        collected_at = datetime.now(timezone.utc).isoformat()
        workbook = build_workbook(
            complete,
            histories,
            definitions,
            cutoff=cutoff,
            collected_at=collected_at,
            query=query,
            space_name=project_name,
            source_timezone=settings.source_timezone,
            preferred_start=settings.start_date_field,
        )
        return PreparedData(
            workbook,
            _json(histories).encode(),
            cutoff,
            collected_at,
            query,
            fingerprint,
            len(complete),
            f"Jira_{_safe_filename(project_name)}.xlsx",
            project_name,
            settings.source_timezone,
        )
    finally:
        gateway.close()


def _collect_clickup_space(settings, item: dict[str, Any], fingerprint: str):
    space_id = str(item.get("id") or "")
    space_name = str(item.get("name") or space_id or "ClickUp Space")
    if not space_id:
        raise ClickUpCollectionError("The selected ClickUp Space has no usable ID.")

    gateway = ClickUpGateway(
        settings.clickup_token,
        workspace_id=str(settings.clickup_workspace_id or ""),
    )
    try:
        with st.spinner(f"Loading tasks from ClickUp — {space_name}..."):
            tasks = gateway.all_tasks_for_space(
                space_id,
                progress=lambda message: st.caption(message),
            )
    finally:
        gateway.close()

    prepared_tasks = []
    for task in tasks:
        copy_task = dict(task)
        copy_task["space_id"] = space_id
        copy_task["project_space_name"] = space_name
        prepared_tasks.append(copy_task)

    return collect_clickup_data(
        None,
        prepared_tasks,
        space_name,
        fingerprint,
        settings.source_timezone,
        space_id=space_id,
        filter_summary=f"Selected ClickUp Space: {space_name}",
        filter_criteria={"space": space_name, "space_id": space_id},
        analysis_mode="project",
    )


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
    return pd.DataFrame(rows)


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
    ])


def _write_excel_table(sheet, headers, rows, start_row=1):
    for row_index, row in enumerate([headers, *rows], start=start_row):
        for column_index, value in enumerate(row, 1):
            cell = sheet.cell(row=row_index, column=column_index, value=_excel_value(value))
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
        ("Completion Rate", rate(kpis.completion_rate), "Completed tasks divided by eligible tasks."),
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
    return pd.DataFrame(rows)


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
    return pd.DataFrame(rows)


def _recommendation_rows(result: CompanyAnalysisResult) -> list[tuple[Any, ...]]:
    return [
        (item.severity, item.title, item.evidence, item.suggested_action)
        for item in result.recommendations
    ]


def _project_excel_bytes(result: CompanyAnalysisResult, scope_label: str) -> bytes:
    model = result.model
    period = f"{model.period_start.isoformat()} to {model.period_end.isoformat()}"
    workbook = Workbook()
    workbook.remove(workbook.active)

    summary = workbook.create_sheet("Project Summary")
    summary.append(["Project Performance Analysis Report"])
    summary.append(["Analysis Scope", scope_label])
    summary.append(["Analysis Period", period])
    summary.append(["Selected Source Spaces", " / ".join(
        f"{item.source_tool}: {item.source_space or 'N/A'}"
        for item in model.source_coverage
    ) or "N/A"])
    summary.append([])
    summary_row = _write_excel_table(
        summary,
        ["KPI", "Value", "Definition"],
        _kpi_detail_rows(model),
        start_row=6,
    )
    summary.cell(row=summary_row + 1, column=1, value="Executive Chart Data").fill = _SECTION_FILL
    chart_rows = []
    for chart in model.executive_charts:
        for point in chart.points:
            chart_rows.append((chart.title, point.label, point.value))
    _write_excel_table(summary, ["Chart", "Category", "Count"], chart_rows, start_row=summary_row + 2)
    _fit_sheet(summary)

    kpis = workbook.create_sheet("KPI Detail")
    _write_excel_table(kpis, ["KPI", "Value", "Definition"], _kpi_detail_rows(model), start_row=1)
    _fit_sheet(kpis)

    outcome = workbook.create_sheet("Delivery Outcome")
    outcome_rows = []
    for chart in model.executive_charts:
        for point in chart.points:
            outcome_rows.append((chart.title, point.label, point.value, chart.note or "N/A"))
    _write_excel_table(outcome, ["Chart", "Category", "Count", "Notes"], outcome_rows, start_row=1)
    _fit_sheet(outcome)

    status = workbook.create_sheet("Status Analysis")
    status_frame = _status_analysis_frame(result)
    _write_excel_table(
        status,
        list(status_frame.columns),
        status_frame.astype(object).values.tolist() if not status_frame.empty else [],
        start_row=1,
    )
    _fit_sheet(status)

    weekly = workbook.create_sheet("Weekly Flow")
    weekly_frame = _weekly_flow_frame(result)
    _write_excel_table(
        weekly,
        list(weekly_frame.columns),
        weekly_frame.astype(object).values.tolist() if not weekly_frame.empty else [],
        start_row=1,
    )
    _fit_sheet(weekly)

    source_space = workbook.create_sheet("Source-Space Breakdown")
    source_frame = _source_space_frame(result)
    _write_excel_table(
        source_space,
        list(source_frame.columns),
        source_frame.astype(object).values.tolist() if not source_frame.empty else [],
        start_row=1,
    )
    _fit_sheet(source_space)

    assignees = workbook.create_sheet("Assignee Analysis")
    assignee_frame = _assignee_frame(result)
    _write_excel_table(
        assignees,
        list(assignee_frame.columns),
        assignee_frame.astype(object).values.tolist() if not assignee_frame.empty else [],
        start_row=1,
    )
    _fit_sheet(assignees)

    bottlenecks = workbook.create_sheet("Bottlenecks")
    bottleneck_rows = [
        (
            item.status.value,
            item.strength,
            item.metrics.tasks_passed_through,
            item.metrics.average_days,
            item.metrics.open_tasks_now,
            item.metrics.overdue_open_tasks,
            "; ".join(item.evidence),
        )
        for item in result.bottlenecks
    ]
    _write_excel_table(
        bottlenecks,
        ["Status", "Assessment", "Tasks Passed Through", "Average Days", "Open Tasks",
         "Open Overdue", "Evidence"],
        bottleneck_rows,
        start_row=1,
    )
    _fit_sheet(bottlenecks)

    recommendations = workbook.create_sheet("Recommendations")
    _write_excel_table(
        recommendations,
        ["Severity", "Recommendation", "Evidence", "Suggested Action"],
        _recommendation_rows(result),
        start_row=1,
    )
    _fit_sheet(recommendations)

    details = workbook.create_sheet("Task Details")
    detail_frame = _task_detail_frame(result)
    _write_excel_table(
        details,
        list(detail_frame.columns),
        detail_frame.astype(object).values.tolist() if not detail_frame.empty else [],
        start_row=1,
    )
    _fit_sheet(details)

    workflow = workbook.create_sheet("Workflow Events")
    workflow_frame = _workflow_events_frame(result)
    _write_excel_table(
        workflow,
        list(workflow_frame.columns),
        workflow_frame.astype(object).values.tolist() if not workflow_frame.empty else [],
        start_row=1,
    )
    _fit_sheet(workflow)

    exceptions = workbook.create_sheet("Exceptions")
    exception_frame = _exception_frame(result)
    _write_excel_table(
        exceptions,
        list(exception_frame.columns),
        exception_frame.astype(object).values.tolist() if not exception_frame.empty else [],
        start_row=1,
    )
    _fit_sheet(exceptions)

    quality = workbook.create_sheet("Data Quality")
    coverage_rows = [
        (item.source_tool, item.source_space, item.unified_project, item.source_available,
         item.task_count, item.history_mode, item.reason or "N/A", "; ".join(item.flags) or "N/A")
        for item in model.source_coverage
    ]
    quality_row = _write_excel_table(
        quality,
        ["Source", "Space", "Unified Scope", "Source Available", "Tasks",
         "History Coverage", "Notes", "Coverage Flags"],
        coverage_rows,
        start_row=1,
    )
    quality.cell(row=quality_row + 1, column=1, value="Data Quality Flags").fill = _SECTION_FILL
    flag_rows = [(item.flag, item.task_count) for item in model.data_quality]
    _write_excel_table(quality, ["Flag", "Task Count"], flag_rows, start_row=quality_row + 2)
    _fit_sheet(quality)

    stream = BytesIO()
    workbook.save(stream)
    return stream.getvalue()


def _docx_value(value):
    if value is None or value == "":
        return "N/A"
    if isinstance(value, (tuple, list, set)):
        return "; ".join(str(item) for item in value) if value else "N/A"
    if isinstance(value, date):
        return value.isoformat()
    return str(value)


def _add_docx_table(document, headers, rows):
    values = list(rows)
    table = document.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    for cell, header in zip(table.rows[0].cells, headers):
        cell.text = _docx_value(header)
        for run in cell.paragraphs[0].runs:
            run.bold = True
    if not values:
        values = [("N/A",) + tuple("" for _ in headers[1:])]
    for row in values:
        cells = table.add_row().cells
        for cell, value in zip(cells, row):
            cell.text = _docx_value(value)


def _project_word_bytes(result: CompanyAnalysisResult, scope_label: str) -> bytes:
    model = result.model
    period = f"{model.period_start.isoformat()} to {model.period_end.isoformat()}"
    document = Document()
    document.styles["Normal"].font.name = "Arial"
    document.styles["Normal"].font.size = Pt(10)
    for section in document.sections:
        section.top_margin = Inches(0.65)
        section.bottom_margin = Inches(0.65)
        section.left_margin = Inches(0.7)
        section.right_margin = Inches(0.7)

    title = document.add_heading("Project Performance Analysis Report", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle = document.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.add_run(scope_label).bold = True
    period_line = document.add_paragraph(f"Analysis Period: {period}")
    period_line.alignment = WD_ALIGN_PARAGRAPH.CENTER

    document.add_heading("1. Report Scope and Source Coverage", level=1)
    document.add_paragraph(
        "This report evaluates the project scope represented by the selected source Spaces. "
        "Tasks are combined into one analysis while the original source and Space remain visible "
        "in every detailed view."
    )
    _add_docx_table(
        document,
        ["Source", "Selected Space", "Tasks", "History Coverage", "Notes"],
        [
            (
                item.source_tool,
                item.source_space,
                item.task_count,
                item.history_mode,
                item.reason or "N/A",
            )
            for item in model.source_coverage
        ],
    )

    document.add_heading("2. Analysis Period and Methodology", level=1)
    document.add_paragraph(
        f"Evaluation period: {model.period_start.isoformat()} to {model.period_end.isoformat()}. "
        "The analysis reconstructs each task's period-end status from the collected source data, "
        "then calculates delivery, workload, timeliness, overdue and workflow metrics. "
        "Jira workflow history is used when available; ClickUp history-dependent measures are "
        "marked as unavailable rather than estimated."
    )
    _add_docx_table(
        document,
        ["Field", "Value"],
        [
            ("Analysis Scope", scope_label),
            ("Analysis Period", period),
            ("Selected Spaces", " / ".join(
                f"{item.source_tool}: {item.source_space or 'N/A'}"
                for item in model.source_coverage
            ) or "N/A"),
            ("Task Records", model.kpis.total_tasks),
        ],
    )

    document.add_heading("3. Executive Summary", level=1)
    completion = "N/A" if model.kpis.completion_rate is None else f"{model.kpis.completion_rate:.1f}%"
    on_time = "N/A" if model.kpis.on_time_completion_rate is None else f"{model.kpis.on_time_completion_rate:.1f}%"
    document.add_paragraph(
        f"The selected Spaces contain {model.kpis.total_tasks} eligible task(s). "
        f"{model.kpis.completed_tasks} task(s) were completed by period end, resulting in a "
        f"{completion} completion rate. The analysis identifies {model.kpis.current_wip} current "
        f"WIP task(s), {model.kpis.overdue_open_tasks} open overdue task(s), and an on-time "
        f"completion rate of {on_time}."
    )

    document.add_heading("4. KPI Summary", level=1)
    _add_docx_table(document, ["KPI", "Value", "Definition"], _kpi_detail_rows(model))

    document.add_heading("5. Delivery Outcome", level=1)
    for chart in model.executive_charts:
        document.add_heading(chart.title, level=2)
        _add_docx_table(
            document,
            ["Category", "Count"],
            [(point.label, point.value) for point in chart.points],
        )
        if chart.note:
            document.add_paragraph(f"Note: {chart.note}")

    document.add_heading("6. Weekly Delivery Flow", level=1)
    weekly = _weekly_flow_frame(result)
    _add_docx_table(
        document,
        list(weekly.columns),
        weekly.astype(object).values.tolist() if not weekly.empty else [],
    )

    document.add_heading("7. Source and Space Contribution", level=1)
    source_space = _source_space_frame(result)
    _add_docx_table(
        document,
        list(source_space.columns),
        source_space.astype(object).values.tolist() if not source_space.empty else [],
    )

    document.add_heading("8. Workload and Individual Achievements", level=1)
    assignees = _assignee_frame(result)
    _add_docx_table(
        document,
        list(assignees.columns),
        assignees.astype(object).values.tolist() if not assignees.empty else [],
    )

    document.add_heading("9. Process and Status Analysis", level=1)
    status = _status_analysis_frame(result)
    _add_docx_table(
        document,
        list(status.columns),
        status.astype(object).values.tolist() if not status.empty else [],
    )

    document.add_heading("10. Bottlenecks and Recommendations", level=1)
    _add_docx_table(
        document,
        ["Status", "Assessment", "Tasks Passed Through", "Average Days", "Open Tasks",
         "Open Overdue", "Evidence"],
        [
            (
                item.status.value,
                item.strength,
                item.metrics.tasks_passed_through,
                item.metrics.average_days,
                item.metrics.open_tasks_now,
                item.metrics.overdue_open_tasks,
                "; ".join(item.evidence),
            )
            for item in result.bottlenecks
        ],
    )
    _add_docx_table(
        document,
        ["Severity", "Recommendation", "Evidence", "Suggested Action"],
        _recommendation_rows(result),
    )

    document.add_heading("11. Task-Level Detail", level=1)
    details = _task_detail_frame(result)
    _add_docx_table(
        document,
        list(details.columns),
        details.astype(object).values.tolist() if not details.empty else [],
    )

    document.add_heading("12. Workflow Events and Exceptions", level=1)
    workflow = _workflow_events_frame(result)
    _add_docx_table(
        document,
        list(workflow.columns),
        workflow.astype(object).values.tolist() if not workflow.empty else [],
    )
    exceptions = _exception_frame(result)
    _add_docx_table(
        document,
        list(exceptions.columns),
        exceptions.astype(object).values.tolist() if not exceptions.empty else [],
    )

    document.add_heading("13. Data Quality and Limitations", level=1)
    _add_docx_table(
        document,
        ["Source", "Space", "Unified Scope", "Source Available", "Tasks",
         "History Coverage", "Notes", "Coverage Flags"],
        [
            (
                item.source_tool,
                item.source_space,
                item.unified_project,
                item.source_available,
                item.task_count,
                item.history_mode,
                item.reason or "N/A",
                "; ".join(item.flags) or "N/A",
            )
            for item in model.source_coverage
        ],
    )
    _add_docx_table(
        document,
        ["Data Quality Flag", "Task Count"],
        [(item.flag, item.task_count) for item in model.data_quality],
    )
    document.add_paragraph(
        "Unavailable source fields are reported as N/A rather than being converted to zero. "
        "Cross-source totals are calculated only from the tasks in the selected Spaces and "
        "selected analysis period. Source-specific identifiers, statuses and Spaces remain "
        "available in Task-Level Detail for auditability."
    )

    document.add_heading("14. Conclusion", level=1)
    document.add_paragraph(
        f"This Project Performance Analysis Report is limited to {scope_label} and the period "
        f"{period}. It is intended to support project-level delivery review, workload follow-up "
        "and evidence-based action planning."
    )

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

    cards = st.columns(len(result.model.cards))
    for column, card in zip(cards, result.model.cards):
        column.metric(card.title, card.value, help=card.supporting_text)

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


def render_project_collection(settings):
    if st.session_state.get("project_analysis") is not None:
        return None, False
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
    if st.session_state.get("project_selection_fingerprint") != selection_fingerprint:
        _clear_project_run()
        st.session_state["project_selection_fingerprint"] = selection_fingerprint
        try:
            with st.spinner("Collecting tasks from the selected project Spaces..."):
                if jira_item:
                    st.session_state["project_jira_prepared"] = _collect_jira_space(
                        settings, jira_item, f"project:jira:{selection_fingerprint}"
                    )
                if clickup_item:
                    st.session_state["project_clickup_prepared"] = _collect_clickup_space(
                        settings, clickup_item, f"project:clickup:{selection_fingerprint}"
                    )
            preview = build_company_preview(
                jira_prepared=st.session_state.get("project_jira_prepared"),
                clickup_prepared=st.session_state.get("project_clickup_prepared"),
            )
            st.session_state["project_preview"] = preview
        except (CollectionError, ClickUpCollectionError) as exc:
            st.session_state.pop("project_selection_fingerprint", None)
            st.session_state.pop("project_jira_prepared", None)
            st.session_state.pop("project_clickup_prepared", None)
            st.error(str(exc))
            return None, False

    preview = st.session_state.get("project_preview", ())
    st.subheader("Selected Project Tasks")
    st.caption(f"{len(preview)} task(s) collected from the selected Space(s).")
    st.dataframe(_preview_frame(preview), hide_index=True, use_container_width=True)

    period_left, period_right = st.columns(2)
    period_start = period_left.date_input(
        "From Date",
        value=None,
        key="project_period_start",
    )
    period_end = period_right.date_input(
        "To Date",
        value=None,
        key="project_period_end",
    )
    valid_dates = isinstance(period_start, date) and isinstance(period_end, date)
    if valid_dates and period_end < period_start:
        st.error("To Date must be on or after From Date.")
        valid_dates = False

    scope_label, scope_slug = _scope_label(jira_item, clickup_item)
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

    if st.session_state.get("project_analysis") is not None:
        return None, False
    return None, False


def render_project_result(st_instance: Any, result: CompanyAnalysisResult) -> None:
    scope_label = st_instance.session_state.get("project_scope_label", "Selected Project Scope")
    scope_slug = st_instance.session_state.get("project_scope_slug", _safe_filename(scope_label))
    _render_project_result(st_instance, result, scope_label, scope_slug)
