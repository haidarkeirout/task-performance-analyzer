"""Excel and Word outputs for the isolated Company Performance analysis.

The module consumes the renderer-neutral dashboard model.  It deliberately
does not reuse or alter either legacy Jira or ClickUp output path, so a
company-level run can be added beside the existing source-specific reports.
"""

from __future__ import annotations

from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Iterable, Sequence

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .dashboard import CompanyDashboardModel
from .kpis import BottleneckCandidate, Recommendation
from .models import TaskPeriodSnapshot


SUMMARY_SHEET = "Executive Dashboard"
BREAKDOWN_SHEET = "Project Summary"
DETAILS_SHEET = "Task Details"
QUALITY_SHEET = "Data Quality"
COMPANY_SHEET_NAMES = (
    SUMMARY_SHEET,
    "Task Metrics",
    "Overall Summary",
    "Process Context",
    "Workflow Events",
    BREAKDOWN_SHEET,
    "Deadline Summary",
    "Weekly Flow",
    "Overdue Tasks",
    "Late Completed Tasks",
    "Open Tasks",
    QUALITY_SHEET,
    "Process Findings",
    "Metric Definitions",
    "By Assignee",
    "By Task Type",
    DETAILS_SHEET,
    "Source Coverage",
)

_HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
_HEADER_FONT = Font(color="FFFFFF", bold=True)
_CARD_FILL = PatternFill("solid", fgColor="EAF2F8")
_THIN_BORDER = Border(bottom=Side(style="thin", color="D9E2F3"))


def _display(value: Any) -> str:
    """Keep unavailable values explicit in executive deliverables."""
    if value is None or value == "":
        return "N/A"
    if isinstance(value, float):
        return f"{value:.1f}"
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, (tuple, list, set)):
        return "; ".join(str(item) for item in value) if value else "N/A"
    return str(value)


def _safe_text(value: Any) -> str:
    return _display(value).replace("\x00", "")


def _excel_value(value: Any) -> Any:
    """Keep dates and numbers typed in Excel while making collections readable."""
    if value is None or value == "":
        return None
    if isinstance(value, (tuple, list, set)):
        return "; ".join(str(item) for item in value) if value else None
    if isinstance(value, (date, datetime, int, float, bool)):
        return value
    return str(value).replace("\x00", "")


def _write_table(sheet: Any, rows: Iterable[Sequence[Any]], *, start_row: int = 1) -> int:
    """Write a simple header-first table and return the next free row."""
    for row_index, row in enumerate(rows, start_row):
        for column_index, value in enumerate(row, 1):
            cell = sheet.cell(row=row_index, column=column_index, value=_safe_text(value))
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if row_index == start_row:
                cell.fill = _HEADER_FILL
                cell.font = _HEADER_FONT
    return row_index + 1 if "row_index" in locals() else start_row


def _write_excel_table(sheet: Any, rows: Iterable[Sequence[Any]], *, start_row: int = 1) -> int:
    """Write a typed worksheet table with the reference workbook styling."""
    values = list(rows)
    if not values:
        return start_row
    for row_index, row in enumerate(values, start_row):
        for column_index, value in enumerate(row, 1):
            cell = sheet.cell(row=row_index, column=column_index, value=_excel_value(value))
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = _THIN_BORDER
            if row_index == start_row:
                cell.fill = _HEADER_FILL
                cell.font = _HEADER_FONT
            elif isinstance(value, (date, datetime)):
                cell.number_format = "yyyy-mm-dd"
    sheet.auto_filter.ref = f"A{start_row}:{get_column_letter(len(values[0]))}{start_row + len(values) - 1}"
    return start_row + len(values)


def _fit_columns(sheet: Any) -> None:
    for column_index in range(1, sheet.max_column + 1):
        width = 12
        for row_index in range(1, min(sheet.max_row, 200) + 1):
            value = sheet.cell(row_index, column_index).value
            width = max(width, min(42, len(str(value or "")) + 2))
        sheet.column_dimensions[get_column_letter(column_index)].width = width
    sheet.freeze_panes = "A2"


def _format_dashboard_sheet(sheet: Any) -> None:
    sheet.sheet_view.showGridLines = False
    sheet.freeze_panes = "A5"


def _chart(sheet: Any, chart_type: str, title: str, anchor: str, *, min_col: int, max_row: int) -> None:
    chart = BarChart() if chart_type == "bar" else LineChart()
    chart.title = title
    chart.style = 10
    chart.height = 7.2
    chart.width = 13.5
    chart.y_axis.title = "Tasks"
    chart.x_axis.title = ""
    data = Reference(sheet, min_col=min_col + 1, min_row=1, max_row=max_row)
    categories = Reference(sheet, min_col=min_col, min_row=2, max_row=max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)
    chart.legend = None
    if chart_type == "bar":
        chart.varyColors = True
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showVal = True
    sheet.add_chart(chart, anchor)


def _task_rows(model: CompanyDashboardModel, snapshots: Sequence[TaskPeriodSnapshot]) -> list[dict[str, Any]]:
    """Build the single Company task grain used by all exported sheets."""
    if snapshots:
        by_key = {(item.task.source_tool, item.task.task_id): item for item in snapshots}
    else:
        by_key = {}
    rows: list[dict[str, Any]] = []
    for detail in model.task_details:
        snapshot = by_key.get((detail.source_tool, detail.task_id))
        task = snapshot.task if snapshot else None
        status = detail.final_status
        completed = detail.final_completion_date
        created = detail.created_date
        actual_start = detail.actual_start_date
        due = detail.due_date
        rows.append({
            "Source Tool": detail.source_tool,
            "Source Space": detail.source_space,
            "Unified Project": detail.unified_project or "Unmapped Project",
            "Task ID": detail.task_id,
            "Task Name": detail.task_name or "Untitled task",
            "Original Status": detail.original_status,
            "Final Status": status,
            "Assignee": detail.assignee_group,
            "Assignees": detail.assignees,
            "Priority": detail.priority,
            "Created Date": created,
            "Planned Start Date": task.planned_start_date if task else None,
            "Actual Start Date": actual_start,
            "Due Date": due,
            "Completion Date": completed,
            "Lead Time (days)": (completed - created).days if completed and created else None,
            "Execution Time (days)": (completed - actual_start).days if completed and actual_start else None,
            "Due Variance (days)": (completed - due).days if completed and due else None,
            "Parent ID": detail.parent_id,
            "Parent Classification": detail.parent_classification,
            "Is Subtask": detail.is_subtask,
            "In Analysis Period": detail.in_analysis_period,
            "Counted in KPIs": detail.counted_in_kpis,
            "Workflow Events": detail.workflow_events,
            "Exception Events": detail.exception_events,
            "Data Quality Flags": detail.data_quality_flags,
            "Exclusion Reason": detail.exclusion_reason,
            "History Available": snapshot.history_available if snapshot else None,
            "Status Intervals": snapshot.status_intervals if snapshot else (),
        })
    return rows


def _workflow_rows(rows: Sequence[dict[str, Any]]) -> list[tuple[Any, ...]]:
    output: list[tuple[Any, ...]] = []
    for row in rows:
        for event in row["Workflow Events"] or ():
            raw = str(event)
            stamp, _, transition = raw.partition(": ")
            before, arrow, after = transition.partition(" → ")
            output.append((row["Source Tool"], row["Unified Project"], row["Task ID"], row["Task Name"], stamp,
                           before if arrow else None, after if arrow else transition, "Status Transition"))
    return output


def _project_rows(rows: Sequence[dict[str, Any]], period_end: date, *, key: str = "Unified Project") -> list[tuple[Any, ...]]:
    grouped: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in rows:
        grouped[str(row.get(key) or "Unknown")].append(row)
    output: list[tuple[Any, ...]] = []
    for label, items in sorted(grouped.items()):
        counted = [item for item in items if item["Counted in KPIs"]]
        completed = [item for item in counted if item["Final Status"] == "Completed"]
        due_completed = [item for item in completed if item["Due Variance (days)"] is not None]
        on_time = [item for item in due_completed if item["Due Variance (days)"] <= 0]
        overdue = [item for item in counted if item["Final Status"] not in {"Completed", "Cancelled", "Rejected"}
                   and item["Due Date"] is not None and item["Due Date"] < period_end]
        output.append((label, len(items), len(completed),
                       round(len(completed) / len(counted) * 100, 1) if counted else None,
                       round(len(on_time) / len(due_completed) * 100, 1) if due_completed else None,
                       len(overdue)))
    return output


def _weekly_rows(rows: Sequence[dict[str, Any]]) -> list[tuple[Any, ...]]:
    grouped: dict[date, list[int]] = defaultdict(lambda: [0, 0])
    for row in rows:
        if row["Created Date"]:
            week = row["Created Date"] - timedelta(days=row["Created Date"].weekday())
            grouped[week][0] += 1
        if row["Completion Date"]:
            week = row["Completion Date"] - timedelta(days=row["Completion Date"].weekday())
            grouped[week][1] += 1
    cumulative = 0
    output: list[tuple[Any, ...]] = []
    for week, (created, completed) in sorted(grouped.items()):
        net = created - completed
        cumulative += net
        output.append((week, created, completed, net, cumulative))
    return output


def write_company_excel(
    model: CompanyDashboardModel,
    output_path: str | Path,
    *,
    snapshots: Iterable[TaskPeriodSnapshot] = (),
    bottlenecks: Iterable[BottleneckCandidate] = (),
    recommendations: Iterable[Recommendation] = (),
) -> Path:
    """Write a reference-style Company Performance analytical workbook.

    This changes only the presentation/export layer.  The model, period rules,
    parent/subtask rules, and KPI calculations are supplied by the existing
    Company analysis pipeline and are not recalculated here.
    """
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    snapshots = tuple(snapshots)
    bottlenecks = tuple(bottlenecks)
    recommendations = tuple(recommendations)
    rows = _task_rows(model, snapshots)
    workbook = Workbook()
    workbook.remove(workbook.active)

    summary = workbook.create_sheet(SUMMARY_SHEET)
    summary.sheet_view.showGridLines = False
    summary.merge_cells("A1:F1")
    # Keep the existing Company export title stable while the body follows the
    # richer reference-report layout.
    summary["A1"] = "Company Performance — Company-Wide Scope"
    summary["A1"].font = Font(bold=True, size=18, color="1F4E78")
    summary["A2"] = "Company-wide Jira and ClickUp performance workbook"
    summary["A3"] = f"Analysis period: {model.period_start.isoformat()} to {model.period_end.isoformat()}"
    summary["A4"] = f"{model.in_period_task_count} task(s) collected across {len(model.source_coverage)} source space(s)"
    for cell in (summary["A2"], summary["A3"], summary["A4"]):
        cell.font = Font(italic=True, color="666666")
    _write_excel_table(summary, [("Metric", "Value", "Definition")] + [
        (card.title, card.value, card.supporting_text) for card in model.cards
    ], start_row=6)
    summary["A6"].fill = _HEADER_FILL
    for row_index in range(7, 6 + len(model.cards) + 1):
        summary.cell(row_index, 1).fill = _CARD_FILL
        summary.cell(row_index, 1).font = Font(bold=True)
    row = 8 + len(model.cards)
    _write_excel_table(summary, [
        ("Metric", "Value", "Definition"),
        ("Total Tasks", str(model.kpis.total_tasks), "All in-period tasks, including visible subtasks."),
        ("Completed Tasks", str(model.kpis.completed_tasks), "Parent/standalone tasks completed by period end."),
        ("Completion Rate", "N/A" if model.kpis.completion_rate is None else f"{model.kpis.completion_rate:.1f}%", "Completed tasks divided by KPI-counted tasks."),
        ("On-Time Rate", "N/A" if model.kpis.on_time_completion_rate is None else f"{model.kpis.on_time_completion_rate:.1f}%", "Completed tasks finished on or before due date."),
        ("Open Overdue", str(model.kpis.overdue_open_tasks), "Open KPI-counted tasks past due date."),
    ], start_row=row)
    # Chart source blocks stay visible for auditability and make the workbook
    # useful even when opened without the application dashboard.
    chart_start = row + 8
    project_counts = sorted(Counter(r["Unified Project"] for r in rows).items())
    _write_excel_table(summary, [("Project", "Tasks")] + project_counts or [("Project", "Tasks"), ("N/A", 0)], start_row=chart_start)
    status_start = chart_start + max(3, len(project_counts) + 3)
    status_counts = Counter(r["Final Status"] for r in rows)
    _write_excel_table(summary, [("Final Status", "Tasks")] + sorted(status_counts.items()) or [("Final Status", "Tasks"), ("N/A", 0)], start_row=status_start)
    weekly_start = status_start + max(3, len(status_counts) + 3)
    weekly_values = _weekly_rows(rows)
    _write_excel_table(summary, [("Week Starting", "Tasks Created", "Tasks Completed")] + weekly_values or [("Week Starting", "Tasks Created", "Tasks Completed")], start_row=weekly_start)
    _chart(summary, "bar", "Tasks by Project", "E6", min_col=1, max_row=chart_start + max(1, len(project_counts)))
    _chart(summary, "bar", "Final Status Distribution", "E21", min_col=1, max_row=status_start + max(1, len(status_counts)))
    _chart(summary, "line", "Weekly Task Flow", "E36", min_col=1, max_row=weekly_start + max(1, len(weekly_values)))
    _fit_columns(summary)
    summary.column_dimensions["A"].width = 28
    _format_dashboard_sheet(summary)

    task_headers = tuple(rows[0].keys()) if rows else ("Source Tool", "Task ID")
    task_metrics = workbook.create_sheet("Task Metrics")
    _write_excel_table(task_metrics, [task_headers] + [tuple(row.get(header) for header in task_headers) for row in rows])
    _fit_columns(task_metrics)

    overall = workbook.create_sheet("Overall Summary")
    _write_excel_table(overall, [("Metric", "Value", "Unit", "Interpretation Basis"),
        ("Total Tasks", model.kpis.total_tasks, "tasks", "All in-period tasks including subtasks"),
        ("Completed Tasks", model.kpis.completed_tasks, "tasks", "KPI-counted tasks at period end"),
        ("Completion Rate", model.kpis.completion_rate, "%", "Completed / KPI-counted tasks"),
        ("On-Time Rate", model.kpis.on_time_completion_rate, "%", "Completed with due date on time"),
        ("Current WIP", model.kpis.current_wip, "tasks", "In Execution or In Review at period end"),
        ("Open Overdue", model.kpis.overdue_open_tasks, "tasks", "Open task with due date before period end")])
    _fit_columns(overall)

    context = workbook.create_sheet("Process Context")
    _write_excel_table(context, [("Field", "Value"),
        ("Analysis Level", "Company"), ("Source Systems", "Jira; ClickUp"),
        ("Period From", model.period_start), ("Period To", model.period_end),
        ("Task Scope", "Active lifetime overlaps selected period"),
        ("Subtask Treatment", "Visible in total task count; excluded from parent KPI rates"),
        ("Status Treatment", "Original status preserved; final status normalized")])
    _fit_columns(context)

    events = workbook.create_sheet("Workflow Events")
    _write_excel_table(events, [("Source Tool", "Unified Project", "Task ID", "Task Name", "Event Date", "From Status", "To Status", "Event Type")]
                       + _workflow_rows(rows))
    _fit_columns(events)

    breakdown = workbook.create_sheet(BREAKDOWN_SHEET)
    _write_excel_table(breakdown, [("Unified Project", "Total Tasks", "Completed", "Completion Rate", "On-Time Rate", "Open Overdue")]
                       + _project_rows(rows, model.period_end))
    _fit_columns(breakdown)

    deadline = workbook.create_sheet("Deadline Summary")
    _write_excel_table(deadline, [("Deadline Status", "Task Count"),
        ("Completed On Time", sum(1 for r in rows if r["Counted in KPIs"] and r["Due Variance (days)"] is not None and r["Due Variance (days)"] <= 0)),
        ("Completed Late", sum(1 for r in rows if r["Counted in KPIs"] and r["Due Variance (days)"] is not None and r["Due Variance (days)"] > 0)),
        ("Open Overdue", sum(1 for r in rows if r["Counted in KPIs"] and r["Final Status"] not in {"Completed", "Cancelled", "Rejected"} and r["Due Date"] and r["Due Date"] < model.period_end)),
        ("No Due Date", sum(1 for r in rows if r["Due Date"] is None))])
    _fit_columns(deadline)

    weekly_sheet = workbook.create_sheet("Weekly Flow")
    _write_excel_table(weekly_sheet, [("Week Starting", "Tasks Created", "Tasks Completed", "Net Flow", "Cumulative Net")]
                       + _weekly_rows(rows))
    _fit_columns(weekly_sheet)

    overdue_rows = [r for r in rows if r["Counted in KPIs"] and r["Final Status"] not in {"Completed", "Cancelled", "Rejected"}
                    and r["Due Date"] and r["Due Date"] < model.period_end]
    overdue = workbook.create_sheet("Overdue Tasks")
    _write_excel_table(overdue, [("Unified Project", "Task ID", "Task Name", "Assignee", "Priority", "Due Date", "Final Status")]
                       + [tuple(r[key] for key in ("Unified Project", "Task ID", "Task Name", "Assignee", "Priority", "Due Date", "Final Status")) for r in overdue_rows])
    _fit_columns(overdue)

    late = workbook.create_sheet("Late Completed Tasks")
    late_rows = [r for r in rows if r["Counted in KPIs"] and r["Final Status"] == "Completed" and r["Due Variance (days)"] is not None and r["Due Variance (days)"] > 0]
    _write_excel_table(late, [("Unified Project", "Task ID", "Task Name", "Assignee", "Due Date", "Completion Date", "Due Variance (days)")]
                       + [tuple(r[key] for key in ("Unified Project", "Task ID", "Task Name", "Assignee", "Due Date", "Completion Date", "Due Variance (days)")) for r in late_rows])
    _fit_columns(late)

    open_tasks = workbook.create_sheet("Open Tasks")
    open_rows = [r for r in rows if r["Counted in KPIs"] and r["Final Status"] not in {"Completed", "Cancelled", "Rejected"}]
    _write_excel_table(open_tasks, [("Unified Project", "Task ID", "Task Name", "Assignee", "Priority", "Final Status", "Due Date")]
                       + [tuple(r[key] for key in ("Unified Project", "Task ID", "Task Name", "Assignee", "Priority", "Final Status", "Due Date")) for r in open_rows])
    _fit_columns(open_tasks)

    quality = workbook.create_sheet(QUALITY_SHEET)
    _write_excel_table(quality, [("Flag", "Task Count")] + [
        (item.flag, item.task_count) for item in model.data_quality
    ])
    row = quality.max_row + 2
    _write_excel_table(quality, [("Source Tool", "History Mode", "Limitation / Reason")] + [
        (item.source_tool, item.history_mode, item.reason or "N/A")
        for item in model.source_coverage
        if not item.source_available or item.history_mode.casefold() != "complete" or item.reason
    ], start_row=row)
    _fit_columns(quality)

    findings = workbook.create_sheet("Process Findings")
    _write_excel_table(findings, [("Type", "Severity", "Finding", "Evidence / Action")]
                       + [("Bottleneck", item.strength, item.status.value, item.evidence) for item in bottlenecks]
                       + [("Recommendation", item.severity, item.title, item.suggested_action) for item in recommendations])
    _fit_columns(findings)

    definitions = workbook.create_sheet("Metric Definitions")
    _write_excel_table(definitions, [("Metric", "Definition", "Scope / Exclusions"),
        ("Total Tasks", "Count of all in-period task rows.", "Subtasks remain visible."),
        ("Completion Rate", "Completed / KPI-counted tasks × 100.", "Subtasks and container parents excluded from rate."),
        ("On-Time Rate", "Completed tasks finished on or before due date / completed tasks with due date.", "Tasks without due dates excluded."),
        ("Open Overdue", "Open tasks with due date before period end.", "Cancelled, rejected, and completed excluded."),
        ("Current WIP", "Tasks in In Execution or In Review at period end.", "Unknown status excluded.")])
    _fit_columns(definitions)

    by_assignee = workbook.create_sheet("By Assignee")
    _write_excel_table(by_assignee, [("Assignee", "Total Tasks", "Completed", "Completion Rate", "Open Overdue")]
                       + _project_rows(rows, model.period_end, key="Assignee"))
    _fit_columns(by_assignee)

    by_type = workbook.create_sheet("By Task Type")
    _write_excel_table(by_type, [("Task Type", "Total Tasks", "Completed", "Completion Rate", "Open Overdue")]
                       + _project_rows(rows, model.period_end, key="Parent Classification"))
    _fit_columns(by_type)

    details = workbook.create_sheet(DETAILS_SHEET)
    detail_headers = ("Source Tool", "Source Space", "Task ID", "Task Name", "Unified Project", "Original Status", "Final Status",
                      "Assignee Group", "Assignees", "Priority", "Created Date", "Due Date", "Actual Start Date", "Final Completion Date",
                      "Workflow Events", "Exception Events", "Data Quality Flags", "Parent ID", "Parent Classification", "Is Subtask",
                      "In Analysis Period", "Counted in KPIs", "Exclusion Reason")
    _write_excel_table(details, [detail_headers] + [tuple(getattr(item, field) for field in (
        "source_tool", "source_space", "task_id", "task_name", "unified_project", "original_status", "final_status", "assignee_group",
        "assignees", "priority", "created_date", "due_date", "actual_start_date", "final_completion_date", "workflow_events",
        "exception_events", "data_quality_flags", "parent_id", "parent_classification", "is_subtask", "in_analysis_period",
        "counted_in_kpis", "exclusion_reason")) for item in model.task_details])
    _fit_columns(details)

    coverage = workbook.create_sheet("Source Coverage")
    _write_excel_table(coverage, [("Source Tool", "Source Space", "Unified Project", "Available", "Tasks in Analysis Period", "History Mode", "Reason", "Flags")] + [
        (item.source_tool, item.source_space, item.unified_project, item.source_available, item.task_count,
         item.history_mode, item.reason, item.flags) for item in model.source_coverage])
    _fit_columns(coverage)

    if tuple(workbook.sheetnames) != COMPANY_SHEET_NAMES:
        raise AssertionError("Company Performance workbook sheet order changed unexpectedly")
    workbook.save(output)
    return output


def write_company_raw_data(snapshots: Iterable[TaskPeriodSnapshot], output_path: str | Path) -> Path:
    """Write audit/raw task fields to a separate workbook, never the executive file."""
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Raw Collected Data"
    _write_table(sheet, [("Source Tool", "Task ID", "Task Name", "Source Space", "Raw Status", "Initial Status", "Assignees", "Created Date", "Due Date", "Collection Timestamp", "Workflow History", "History Complete", "Parent ID", "Parent Classification", "Is Subtask", "Flags")] + [
        (snapshot.task.source_tool, snapshot.task.task_id, snapshot.task.task_name,
         snapshot.task.source_space, snapshot.task.raw_status, snapshot.task.initial_status,
         snapshot.task.assignees, snapshot.task.created_date, snapshot.task.due_date,
         snapshot.task.collection_timestamp,
         tuple(f"{event.changed_at.isoformat()}: {event.from_status} → {event.to_status}" for event in snapshot.task.workflow_history),
         snapshot.task.history_complete, snapshot.task.parent_id,
         snapshot.task.parent_classification.value,
         snapshot.task.parent_classification.value == "Subtask",
         tuple(sorted(snapshot.task.data_quality_flags)))
        for snapshot in snapshots
    ])
    _fit_columns(sheet)
    workbook.save(output)
    return output


def _configure_document(document: Document) -> None:
    document.styles["Normal"].font.name = "Arial"
    document.styles["Normal"].font.size = Pt(10)
    for style_name in ("Title", "Heading 1", "Heading 2"):
        style = document.styles[style_name]
        style.font.name = "Arial"
        style.font.color.rgb = RGBColor(0, 0, 0)
    for section in document.sections:
        section.top_margin = Inches(0.65)
        section.bottom_margin = Inches(0.65)
        section.left_margin = Inches(0.7)
        section.right_margin = Inches(0.7)
        footer = section.footer.paragraphs[0]
        footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        footer.text = "Company Performance Analysis"
        for run in footer.runs:
            run.font.name = "Arial"
            run.font.size = Pt(8)
            run.font.color.rgb = RGBColor(102, 102, 102)


def _shade_cell(cell: Any, fill: str) -> None:
    properties = cell._tc.get_or_add_tcPr()
    shading = properties.find(qn("w:shd"))
    if shading is None:
        shading = OxmlElement("w:shd")
        properties.append(shading)
    shading.set(qn("w:fill"), fill)


def _set_cell_padding(cell: Any, padding: int = 90) -> None:
    properties = cell._tc.get_or_add_tcPr()
    margins = properties.find(qn("w:tcMar"))
    if margins is None:
        margins = OxmlElement("w:tcMar")
        properties.append(margins)
    for side in ("top", "start", "bottom", "end"):
        node = margins.find(qn(f"w:{side}"))
        if node is None:
            node = OxmlElement(f"w:{side}")
            margins.append(node)
        node.set(qn("w:w"), str(padding))
        node.set(qn("w:type"), "dxa")


def _add_table(document: Document, headers: Sequence[str], rows: Iterable[Sequence[Any]]) -> None:
    values = list(rows)
    table = document.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    for cell, header in zip(table.rows[0].cells, headers):
        cell.text = _safe_text(header)
        _shade_cell(cell, "1F4E78")
        _set_cell_padding(cell)
        for run in cell.paragraphs[0].runs:
            run.bold = True
            run.font.name = "Arial"
            run.font.size = Pt(9)
            run.font.color.rgb = RGBColor(255, 255, 255)
    if values:
        for row_index, values_row in enumerate(values):
            cells = table.add_row().cells
            for cell, value in zip(cells, values_row):
                cell.text = _safe_text(value)
                _set_cell_padding(cell)
                if row_index % 2 == 1:
                    _shade_cell(cell, "F2F6FA")
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = "Arial"
                        run.font.size = Pt(9)
    else:
        cells = table.add_row().cells
        cells[0].text = "N/A — no eligible data for this section."
        _set_cell_padding(cells[0])


def _compact_kpi_definition(title: str, fallback: str) -> str:
    definitions = {
        "Total Projects": "Distinct inferred project names in the selected Company scope.",
        "Total Tasks": "All in-period tasks; subtasks remain visible in this count.",
        "Completed Tasks": "KPI-counted tasks completed at period end.",
        "Completion Rate": "Completed tasks divided by KPI-counted tasks.",
        "On-Time Rate": "Completed tasks finished on or before their due date.",
        "On-Time Completion Rate": "Completed tasks finished on or before their due date.",
        "Open Overdue": "Open tasks with a due date before period end.",
        "Overdue Open Tasks": "Open tasks with a due date before period end.",
        "Current WIP": "Tasks in In Execution or In Review at period end.",
    }
    return definitions.get(title, _safe_text(fallback).replace("**", ""))


def write_company_word_report(
    model: CompanyDashboardModel,
    output_path: str | Path,
    *,
    snapshots: Iterable[TaskPeriodSnapshot] = (),
    bottlenecks: Iterable[BottleneckCandidate] = (),
    recommendations: Iterable[Recommendation] = (),
) -> Path:
    """Create the Company report in the same evidence-led shape as the reference."""
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    bottlenecks = tuple(bottlenecks)
    recommendations = tuple(recommendations)
    snapshots = tuple(snapshots)
    task_rows = _task_rows(model, snapshots)
    document = Document()
    _configure_document(document)
    title = document.add_heading("Company Performance Analysis", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle = document.add_paragraph("Company-wide analysis across all collected Jira projects and ClickUp spaces")
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in subtitle.runs:
        run.font.name = "Arial"
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(89, 89, 89)
    _add_table(document, ("Report Field", "Value"), [
        ("Analysis Level", "Company"),
        ("Source Systems", "Jira and ClickUp"),
        ("Observation Period", f"{model.period_start.isoformat()} to {model.period_end.isoformat()}"),
        ("Tasks in Scope", model.in_period_task_count),
        ("Tasks Counted in KPI Rates", model.kpis.total_tasks),
    ])

    document.add_heading("1. Executive Summary", level=1)
    document.add_paragraph(
        f"The analysis includes {_display(model.in_period_task_count)} task(s) within the selected analysis period, "
        f"of which {_display(model.kpis.total_tasks)} are included in headline KPI calculations. "
        f"Completion rate: {_display(model.kpis.completion_rate)}%; current WIP: {_display(model.kpis.current_wip)}; "
        f"open overdue work: {_display(model.kpis.overdue_open_tasks)}. The report combines all accessible "
        f"projects and source spaces under one Company view; it does not evaluate a space as a separate analysis level."
    )

    document.add_heading("2. Scope and Analysis Period", level=1)
    document.add_paragraph(
        "Tasks are included when their active period overlaps the selected analysis period. "
        "Headline KPIs use the final task status at the end of that period. The selected From/To dates are "
        "preserved in the report so the same period can be reproduced in Dashboard, Excel, and Word."
    )

    document.add_heading("3. Data Sources and Coverage", level=1)
    _add_table(document, ("Source", "Space", "Unified Project", "Tasks in Analysis Period", "History Coverage", "Notes"), [
        (item.source_tool, item.source_space, item.unified_project, item.task_count,
         item.history_mode, item.reason) for item in model.source_coverage
    ])

    document.add_heading("4. Methodology and Assignment Rules", level=1)
    _add_table(document, ("Step", "Company-level rule"), [
        ("Collection", "Read all accessible Jira projects and ClickUp spaces."),
        ("Grouping", "Keep the original source space and group records under the unified Company view."),
        ("Period", "Keep only tasks whose active lifetime overlaps the selected From/To period."),
        ("Status", "Preserve original status and expose a normalized final status."),
        ("Hierarchy", "Keep parent/subtask relationships; show subtasks in totals but exclude them from KPI rates."),
    ])
    document.add_paragraph(
        "Statuses are normalised across Jira and ClickUp. Actual Start is the first entry into In Progress; "
        "a direct completion without that event remains N/A and is flagged. Multiple assignees are reported "
        "as a separate group, while the underlying names remain available in Task Details. "
        "Parent tasks and subtasks are identified from the source parent relationship; subtasks remain visible "
        "in the task count but are excluded from parent-level performance KPIs."
    )

    document.add_heading("5. Headline KPIs", level=1)
    _add_table(document, ("KPI", "Value", "Definition"), [
        (card.title, card.value, _compact_kpi_definition(card.title, card.supporting_text)) for card in model.cards
    ])

    document.add_heading("6. Delivery Outcome", level=1)
    delivery = next(chart for chart in model.executive_charts if chart.key == "delivery-outcome")
    _add_table(document, ("Final Status", "Task Count"), [(point.label, point.value) for point in delivery.points])
    if delivery.note:
        document.add_paragraph(delivery.note)

    document.add_heading("7. Workload and Overdue Work", level=1)
    workload = next(chart for chart in model.executive_charts if chart.key == "workload-by-project")
    overdue = next(chart for chart in model.executive_charts if chart.key == "overdue-by-priority")
    _add_table(document, ("Unified Project", "Task Count"), [(point.label, point.value) for point in workload.points])
    _add_table(document, ("Overdue Priority", "Open Task Count"), [(point.label, point.value) for point in overdue.points])
    _add_table(document, ("Unified Project", "Total Tasks", "Completed", "Completion Rate", "On-Time Rate", "Open Overdue"),
               _project_rows(task_rows, model.period_end))

    document.add_heading("8. Workflow Efficiency", level=1)
    _add_table(document, ("Metric", "Value"), [
        ("Average Time to Start (days)", model.kpis.average_time_to_start_days),
        ("Average Execution Duration (days)", model.kpis.average_execution_duration_days),
        ("Average Lead Time (days)", model.kpis.average_lead_time_days),
        ("On-Time Completion Rate", None if model.kpis.on_time_completion_rate is None else f"{model.kpis.on_time_completion_rate:.1f}%"),
    ])

    document.add_heading("9. Bottleneck Candidates", level=1)
    document.add_paragraph("Candidates indicate evidence requiring follow-up; they do not confirm a root cause.")
    _add_table(document, ("Status", "Assessment", "Evidence"), [
        (item.status.value, item.strength, item.evidence) for item in bottlenecks
    ])

    document.add_heading("10. Data Quality and Limitations", level=1)
    _add_table(document, ("Data Quality Flag", "Task Count"), [
        (item.flag, item.task_count) for item in model.data_quality
    ])
    document.add_paragraph(
        "Historical status metrics are excluded when history is unavailable for a past period end. "
        "A current-status fallback is permitted only for a collection-date snapshot and is flagged."
    )

    document.add_heading("11. Recommendations and Next Steps", level=1)
    _add_table(document, ("Severity", "Recommendation", "Evidence", "Suggested Action"), [
        (item.severity, item.title, item.evidence, item.suggested_action) for item in recommendations
    ])

    document.add_heading("12. Workbook Architecture and Data Roles", level=1)
    document.add_paragraph(
        "The Excel output is organized for both executive reading and auditability. Executive Dashboard and "
        "Overall Summary contain the headline view; Project Summary, By Assignee, and By Task Type provide "
        "Company-level breakdowns; Task Metrics, Task Details, Workflow Events, and Source Coverage preserve "
        "the evidence used to explain the results; Deadline Summary, Weekly Flow, Overdue Tasks, and Late "
        "Completed Tasks support operational follow-up."
    )
    document.add_heading("13. Interpretation Rules and Handoff Objective", level=1)
    document.add_paragraph(
        "Use the Company dashboard and exported files to answer: What is the overall state of the company-wide "
        "workload across Jira and ClickUp during the selected period? Counts in the dashboard, Excel, and Word "
        "must match because each output is generated from the same Company model and the same selected task set. "
        "Subtasks may increase the visible Total Tasks count, but they do not independently increase completion "
        "or completion-rate numerators and denominators."
    )

    for paragraph in document.paragraphs:
        if paragraph.style.name.startswith("Heading"):
            paragraph.paragraph_format.keep_with_next = True
        if paragraph.text.startswith("6. Delivery Outcome"):
            paragraph.paragraph_format.page_break_before = True

    document.save(output)
    return output
