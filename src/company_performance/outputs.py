"""Excel and Word outputs for the isolated Company Performance analysis.

The module consumes the renderer-neutral dashboard model.  It deliberately
does not reuse or alter either legacy Jira or ClickUp output path, so a
company-level run can be added beside the existing source-specific reports.
"""

from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Any, Iterable, Sequence

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .dashboard import CompanyDashboardModel
from .kpis import BottleneckCandidate, Recommendation
from .models import TaskPeriodSnapshot


SUMMARY_SHEET = "Company Summary"
BREAKDOWN_SHEET = "Performance Breakdown"
DETAILS_SHEET = "Task Details"
QUALITY_SHEET = "Exceptions & Data Quality"
COMPANY_SHEET_NAMES = (SUMMARY_SHEET, BREAKDOWN_SHEET, DETAILS_SHEET, QUALITY_SHEET)

_HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
_HEADER_FONT = Font(color="FFFFFF", bold=True)
_SECTION_FILL = PatternFill("solid", fgColor="D9EAF7")


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


def _fit_columns(sheet: Any) -> None:
    for column_index in range(1, sheet.max_column + 1):
        width = 12
        for row_index in range(1, min(sheet.max_row, 200) + 1):
            value = sheet.cell(row_index, column_index).value
            width = max(width, min(42, len(str(value or "")) + 2))
        sheet.column_dimensions[get_column_letter(column_index)].width = width
    sheet.freeze_panes = "A2"


def _append_section(sheet: Any, row: int, title: str) -> int:
    cell = sheet.cell(row=row, column=1, value=title)
    cell.fill = _SECTION_FILL
    cell.font = Font(bold=True, color="1F1F1F")
    return row + 1


def write_company_excel(
    model: CompanyDashboardModel,
    output_path: str | Path,
    *,
    bottlenecks: Iterable[BottleneckCandidate] = (),
    recommendations: Iterable[Recommendation] = (),
) -> Path:
    """Write the approved four-sheet Company Performance workbook."""
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    workbook.remove(workbook.active)

    summary = workbook.create_sheet(SUMMARY_SHEET)
    summary.append(["Company Performance — Selected Projects"])
    summary["A1"].font = Font(bold=True, size=14, color="1F4E78")
    summary.append(["Analysis Period", f"{model.period_start.isoformat()} to {model.period_end.isoformat()}"])
    _append_section(summary, 4, "Executive KPI Cards")
    _write_table(summary, [("KPI", "Value", "Definition")] + [
        (card.title, card.value, card.supporting_text) for card in model.cards
    ], start_row=5)
    row = summary.max_row + 2
    row = _append_section(summary, row, "Executive Charts (data)")
    chart_rows: list[tuple[str, str, int, str]] = [("Chart", "Label", "Count", "Note")]
    for chart in model.executive_charts:
        if chart.points:
            chart_rows.extend((chart.title, point.label, point.value, chart.note or "") for point in chart.points)
        else:
            chart_rows.append((chart.title, "N/A", 0, chart.note or "No eligible data"))
    _write_table(summary, chart_rows, start_row=row)
    _fit_columns(summary)

    breakdown = workbook.create_sheet(BREAKDOWN_SHEET)
    breakdown.append(["Performance Breakdown"])
    breakdown["A1"].font = Font(bold=True, size=14, color="1F4E78")
    _append_section(breakdown, 3, "Source Coverage")
    _write_table(breakdown, [("Source Tool", "Source Space", "Unified Project", "Available", "Task Count", "History Mode", "Reason", "Flags")] + [
        (item.source_tool, item.source_space, item.unified_project, item.source_available,
         item.task_count, item.history_mode, item.reason, item.flags)
        for item in model.source_coverage
    ] or [("Source Tool", "Source Space", "Unified Project", "Available", "Task Count", "History Mode", "Reason", "Flags")], start_row=4)
    row = breakdown.max_row + 2
    _append_section(breakdown, row, "Bottleneck Candidates")
    row += 1
    _write_table(breakdown, [("Status", "Assessment", "Evidence", "Average Days", "Open Tasks", "Overdue Open Tasks", "Repeated Returns")] + [
        (candidate.status.value, candidate.strength, candidate.evidence,
         candidate.metrics.average_days, candidate.metrics.open_tasks_now,
         candidate.metrics.overdue_open_tasks, candidate.metrics.repeated_returns)
        for candidate in bottlenecks
    ], start_row=row)
    row = breakdown.max_row + 2
    _append_section(breakdown, row, "Recommendations")
    _write_table(breakdown, [("Severity", "Title", "Evidence", "Suggested Action")] + [
        (item.severity, item.title, item.evidence, item.suggested_action)
        for item in recommendations
    ], start_row=row + 1)
    _fit_columns(breakdown)

    details = workbook.create_sheet(DETAILS_SHEET)
    _write_table(details, [(
        "Source Tool", "Source Space", "Task ID", "Task Name", "Unified Project",
        "Original Status", "Final Status", "Assignee Group", "Assignees", "Priority",
        "Created Date", "Due Date", "Actual Start Date", "Final Completion Date",
        "Workflow Events", "Exception Events", "Data Quality Flags",
    )] + [
        (item.source_tool, item.source_space, item.task_id, item.task_name, item.unified_project,
         item.original_status, item.final_status, item.assignee_group, item.assignees, item.priority,
         item.created_date, item.due_date, item.actual_start_date, item.final_completion_date,
         item.workflow_events, item.exception_events, item.data_quality_flags)
        for item in model.task_details
    ])
    _fit_columns(details)

    quality = workbook.create_sheet(QUALITY_SHEET)
    _append_section(quality, 1, "Data Quality Flags")
    _write_table(quality, [("Flag", "Task Count")] + [
        (item.flag, item.task_count) for item in model.data_quality
    ], start_row=2)
    row = quality.max_row + 2
    _append_section(quality, row, "Coverage Limitations")
    _write_table(quality, [("Source Tool", "History Mode", "Limitation / Reason")] + [
        (item.source_tool, item.history_mode, item.reason or "N/A")
        for item in model.source_coverage
        if not item.source_available or item.history_mode.casefold() != "complete" or item.reason
    ], start_row=row + 1)
    _fit_columns(quality)

    if tuple(workbook.sheetnames) != COMPANY_SHEET_NAMES:
        raise AssertionError("Company Performance workbook must contain exactly four approved sheets")
    workbook.save(output)
    return output


def write_company_raw_data(snapshots: Iterable[TaskPeriodSnapshot], output_path: str | Path) -> Path:
    """Write audit/raw task fields to a separate workbook, never the executive file."""
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Raw Collected Data"
    _write_table(sheet, [("Source Tool", "Task ID", "Task Name", "Source Space", "Raw Status", "Initial Status", "Assignees", "Created Date", "Due Date", "Collection Timestamp", "Workflow History", "History Complete", "Flags")] + [
        (snapshot.task.source_tool, snapshot.task.task_id, snapshot.task.task_name,
         snapshot.task.source_space, snapshot.task.raw_status, snapshot.task.initial_status,
         snapshot.task.assignees, snapshot.task.created_date, snapshot.task.due_date,
         snapshot.task.collection_timestamp,
         tuple(f"{event.changed_at.isoformat()}: {event.from_status} → {event.to_status}" for event in snapshot.task.workflow_history),
         snapshot.task.history_complete, tuple(sorted(snapshot.task.data_quality_flags)))
        for snapshot in snapshots
    ])
    _fit_columns(sheet)
    workbook.save(output)
    return output


def _configure_document(document: Document) -> None:
    document.styles["Normal"].font.name = "Arial"
    document.styles["Normal"].font.size = Pt(10)
    for section in document.sections:
        section.top_margin = Inches(0.65)
        section.bottom_margin = Inches(0.65)
        section.left_margin = Inches(0.7)
        section.right_margin = Inches(0.7)


def _add_table(document: Document, headers: Sequence[str], rows: Iterable[Sequence[Any]]) -> None:
    values = list(rows)
    table = document.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    for cell, header in zip(table.rows[0].cells, headers):
        cell.text = _safe_text(header)
        for run in cell.paragraphs[0].runs:
            run.bold = True
    if values:
        for values_row in values:
            cells = table.add_row().cells
            for cell, value in zip(cells, values_row):
                cell.text = _safe_text(value)
    else:
        cells = table.add_row().cells
        cells[0].text = "N/A — no eligible data for this section."


def write_company_word_report(
    model: CompanyDashboardModel,
    output_path: str | Path,
    *,
    bottlenecks: Iterable[BottleneckCandidate] = (),
    recommendations: Iterable[Recommendation] = (),
) -> Path:
    """Create the approved English, evidence-led 11-section Word report."""
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    bottlenecks = tuple(bottlenecks)
    recommendations = tuple(recommendations)
    document = Document()
    _configure_document(document)
    title = document.add_heading("Company Performance Analysis", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    document.add_paragraph("Selected Projects")
    document.add_paragraph(f"Analysis period: {model.period_start.isoformat()} to {model.period_end.isoformat()}")

    document.add_heading("1. Executive Summary", level=1)
    document.add_paragraph(
        f"The analysis covers {_display(model.kpis.total_tasks)} eligible task(s). "
        f"Completion rate: {_display(model.kpis.completion_rate)}%; current WIP: {_display(model.kpis.current_wip)}; "
        f"open overdue work: {_display(model.kpis.overdue_open_tasks)}."
    )

    document.add_heading("2. Scope and Analysis Period", level=1)
    document.add_paragraph(
        "Tasks are included when their active period overlaps the selected analysis period. "
        "Headline KPIs use the final task status at the end of that period."
    )

    document.add_heading("3. Data Sources and Coverage", level=1)
    _add_table(document, ("Source", "Space", "Unified Project", "Tasks", "History Coverage", "Notes"), [
        (item.source_tool, item.source_space, item.unified_project, item.task_count,
         item.history_mode, item.reason) for item in model.source_coverage
    ])

    document.add_heading("4. Methodology and Assignment Rules", level=1)
    document.add_paragraph(
        "Statuses are normalised across Jira and ClickUp. Actual Start is the first entry into In Progress; "
        "a direct completion without that event remains N/A and is flagged. Multiple assignees are reported "
        "as a separate group, while the underlying names remain available in Task Details."
    )

    document.add_heading("5. Headline KPIs", level=1)
    _add_table(document, ("KPI", "Value", "Definition"), [
        (card.title, card.value, card.supporting_text) for card in model.cards
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

    document.save(output)
    return output
