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
from excel_safety import write_excel_cell

from .dashboard import CompanyDashboardModel
from .kpis import BottleneckCandidate, Recommendation
from .models import TaskPeriodSnapshot


SUMMARY_SHEET = "Company_Executive_Dashboard"
BREAKDOWN_SHEET = "Project Summary"
DETAILS_SHEET = "Task Details"
QUALITY_SHEET = "Data Quality"
COMPANY_SHEET_NAMES = (
    SUMMARY_SHEET,
    BREAKDOWN_SHEET,
    DETAILS_SHEET,
    "Overdue Tasks",
    "Late Completed Tasks",
    "Bottlenecks",
    "Workflow Exceptions",
    "Weekly Flow",
    QUALITY_SHEET,
    "Analysis Context",
    "Metric Definitions",
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
            cell = write_excel_cell(sheet, row_index, column_index, _safe_text(value))
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
            cell = write_excel_cell(sheet, row_index, column_index, _excel_value(value))
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
    """Write the exact Company Performance workbook requested by the user."""
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    snapshots = tuple(snapshots)
    bottlenecks = tuple(bottlenecks)
    recommendations = tuple(recommendations)
    rows = _task_rows(model, snapshots)
    workbook = Workbook()
    workbook.remove(workbook.active)

    dashboard = workbook.create_sheet("Company_Executive_Dashboard")
    dashboard.sheet_view.showGridLines = False
    dashboard.merge_cells("A1:K1")
    dashboard["A1"] = "Company Performance"
    dashboard["A1"].fill = PatternFill("solid", fgColor="17324D")
    dashboard["A1"].font = Font(color="FFFFFF", bold=True, size=16)
    dashboard["A1"].alignment = Alignment(horizontal="center")
    dashboard.merge_cells("A2:K2")
    dashboard["A2"] = f"Jira and ClickUp | Analysis period: {model.period_start.isoformat()} to {model.period_end.isoformat()}"
    dashboard["A2"].font = Font(color="6B7280", italic=True)
    dashboard["A2"].alignment = Alignment(horizontal="center")
    for column in "ABCDEFGHIJK":
        dashboard.column_dimensions[column].width = 16

    for index, card in enumerate(model.cards):
        _dashboard_start = 1 + (index % 6) * 2
        _dashboard_end = _dashboard_start + 1
        dashboard.merge_cells(start_row=3, start_column=_dashboard_start, end_row=3, end_column=_dashboard_end)
        dashboard.merge_cells(start_row=4, start_column=_dashboard_start, end_row=5, end_column=_dashboard_end)
        dashboard.cell(3, _dashboard_start, card.title).fill = _HEADER_FILL
        dashboard.cell(3, _dashboard_start).font = _HEADER_FONT
        dashboard.cell(3, _dashboard_start).alignment = Alignment(horizontal="center")
        dashboard.cell(4, _dashboard_start, card.value)
        dashboard.cell(4, _dashboard_start).font = Font(size=16, bold=True, color="1F2937")
        dashboard.cell(4, _dashboard_start).alignment = Alignment(horizontal="center", vertical="center")
    dashboard["A7"] = f"Tasks in scope: {model.in_period_task_count}"
    dashboard["D7"] = f"Projects: {len({row['Unified Project'] for row in rows})}"
    dashboard["G7"] = f"Source spaces: {len(model.source_coverage)}"
    for cell in ("A7", "D7", "G7"):
        dashboard[cell].font = Font(bold=True)

    counted = [row for row in rows if row["Counted in KPIs"]]
    project_groups: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in counted:
        project_groups[str(row["Unified Project"] or "Unknown")].append(row)
    project_counts = sorted((name, len(items)) for name, items in project_groups.items())
    project_comparison = []
    for name, items in sorted(project_groups.items()):
        completed = [item for item in items if item["Final Status"] == "Completed"]
        with_due = [item for item in completed if item["Due Variance (days)"] is not None]
        on_time = [item for item in with_due if item["Due Variance (days)"] <= 0]
        overdue = [
            item for item in items
            if item["Final Status"] not in {"Completed", "Cancelled", "Rejected"}
            and item["Due Date"] is not None and item["Due Date"] < model.period_end
        ]
        project_comparison.append([
            name, len(items), len(completed),
            None if not items else round(len(completed) / len(items) * 100.0, 1),
            None if not with_due else round(len(on_time) / len(with_due) * 100.0, 1),
            len(overdue),
        ])
    project_counts = project_counts or [("No data", 0)]
    project_comparison = project_comparison or [["No data", 0, 0, None, None, 0]]

    weekly = defaultdict(lambda: [0, 0])
    for row in counted:
        if row["Created Date"]:
            week = row["Created Date"] - timedelta(days=row["Created Date"].weekday())
            weekly[week][0] += 1
        if row["Completion Date"]:
            week = row["Completion Date"] - timedelta(days=row["Completion Date"].weekday())
            weekly[week][1] += 1
    weekly_rows = [[week, values[0], values[1]] for week, values in sorted(weekly.items())] or [[model.period_start, 0, 0]]
    status_counts = sorted(Counter(row["Final Status"] for row in counted).items()) or [("No data", 0)]

    weekly_start = 55
    _write_excel_table(dashboard, [["Week Starting", "Tasks Created", "Tasks Completed"], *weekly_rows], start_row=weekly_start)
    status_start = weekly_start + len(weekly_rows) + 3
    _write_excel_table(dashboard, [["Status", "Tasks"], *status_counts], start_row=status_start)
    project_start = status_start + len(status_counts) + 3
    _write_excel_table(dashboard, [["Project", "Tasks"], *project_counts], start_row=project_start)
    comparison_start = project_start + len(project_counts) + 3
    _write_excel_table(dashboard, [["Project", "Completion Rate", "On-Time Rate"], *[
        [row[0], row[3], row[4]] for row in project_comparison
    ]], start_row=comparison_start)

    weekly_last = weekly_start + len(weekly_rows)
    line = LineChart()
    line.title = "Weekly Task Flow"
    line.style = 13
    line.height = 7
    line.width = 13
    line.add_data(Reference(dashboard, min_col=2, max_col=3, min_row=weekly_start, max_row=weekly_last), titles_from_data=True)
    line.set_categories(Reference(dashboard, min_col=1, min_row=weekly_start + 1, max_row=weekly_last))
    line.legend.position = "b"
    dashboard.add_chart(line, "A10")

    status_last = status_start + len(status_counts)
    status_chart = BarChart()
    status_chart.type = "col"
    status_chart.style = 10
    status_chart.title = "Task Distribution by Status"
    status_chart.height = 7
    status_chart.width = 13
    status_chart.add_data(Reference(dashboard, min_col=2, max_col=2, min_row=status_start, max_row=status_last), titles_from_data=True)
    status_chart.set_categories(Reference(dashboard, min_col=1, min_row=status_start + 1, max_row=status_last))
    dashboard.add_chart(status_chart, "I10")

    project_last = project_start + len(project_counts)
    project_chart = BarChart()
    project_chart.type = "col"
    project_chart.style = 10
    project_chart.title = "Work Distribution by Project"
    project_chart.height = 7
    project_chart.width = 13
    project_chart.add_data(Reference(dashboard, min_col=2, max_col=2, min_row=project_start, max_row=project_last), titles_from_data=True)
    project_chart.set_categories(Reference(dashboard, min_col=1, min_row=project_start + 1, max_row=project_last))
    dashboard.add_chart(project_chart, "A26")

    comparison_last = comparison_start + len(project_comparison)
    comparison_chart = BarChart()
    comparison_chart.type = "col"
    comparison_chart.style = 10
    comparison_chart.title = "Project Comparison"
    comparison_chart.height = 7
    comparison_chart.width = 13
    comparison_chart.add_data(Reference(dashboard, min_col=2, max_col=3, min_row=comparison_start, max_row=comparison_last), titles_from_data=True)
    comparison_chart.set_categories(Reference(dashboard, min_col=1, min_row=comparison_start + 1, max_row=comparison_last))
    comparison_chart.legend.position = "b"
    dashboard.add_chart(comparison_chart, "I26")
    _fit_columns(dashboard)

    project_sheet = workbook.create_sheet("Project Summary")
    _write_excel_table(project_sheet, [["Project", "Total Tasks", "Completed", "Completion Rate", "On-Time Rate", "Open Overdue"], *project_comparison])
    _fit_columns(project_sheet)

    detail_headers = tuple(rows[0].keys()) if rows else ("Source Tool", "Task ID")
    details = workbook.create_sheet("Task Details")
    _write_excel_table(details, [detail_headers] + [tuple(row.get(header) for header in detail_headers) for row in rows])
    _fit_columns(details)

    overdue_rows = [
        row for row in rows
        if row["Counted in KPIs"] and row["Final Status"] not in {"Completed", "Cancelled", "Rejected"}
        and row["Due Date"] is not None and row["Due Date"] < model.period_end
    ]
    late_rows = [
        row for row in rows
        if row["Counted in KPIs"] and row["Final Status"] == "Completed"
        and row["Due Variance (days)"] is not None and row["Due Variance (days)"] > 0
    ]
    exception_rows = [
        [row["Unified Project"], row["Task ID"], row["Task Name"], row["Source Tool"], row["Source Space"], event]
        for row in rows for event in row["Exception Events"] or ()
    ] or [["No workflow exception", "", "", "", "", ""]]
    bottleneck_rows = [
        [item.status.value, item.strength, "; ".join(item.evidence),
         item.metrics.average_days, item.metrics.open_tasks_now, item.metrics.overdue_open_tasks]
        for item in bottlenecks
    ] or [["No candidate", "N/A", "No bottleneck candidate was identified.", None, 0, 0]]

    _write_excel_table(
        workbook.create_sheet("Overdue Tasks"),
        [["Project", "Task ID", "Task Name", "Source", "Space", "Assignee", "Priority", "Due Date", "Final Status"]] + [
            [row["Unified Project"], row["Task ID"], row["Task Name"], row["Source Tool"], row["Source Space"],
             row["Assignee"], row["Priority"], row["Due Date"], row["Final Status"]]
            for row in overdue_rows
        ],
    )
    _write_excel_table(
        workbook.create_sheet("Late Completed Tasks"),
        [["Project", "Task ID", "Task Name", "Source", "Space", "Assignee", "Due Date", "Completion Date", "Due Variance (days)"]] + [
            [row["Unified Project"], row["Task ID"], row["Task Name"], row["Source Tool"], row["Source Space"],
             row["Assignee"], row["Due Date"], row["Completion Date"], row["Due Variance (days)"]]
            for row in late_rows
        ],
    )
    _write_excel_table(workbook.create_sheet("Bottlenecks"), [["Stage", "Assessment", "Evidence", "Average Days", "Open Tasks", "Open Overdue"], *bottleneck_rows])
    _write_excel_table(workbook.create_sheet("Workflow Exceptions"), [["Project", "Task ID", "Task Name", "Source", "Space", "Exception"], *exception_rows])
    _write_excel_table(
        workbook.create_sheet("Weekly Flow"),
        [["Week Starting", "Tasks Created", "Tasks Completed"]] + weekly_rows,
    )

    quality = workbook.create_sheet("Data Quality")
    _write_excel_table(quality, [["Data Quality Flag", "Task Count"]] + [
        [item.flag, item.task_count] for item in model.data_quality
    ] or [["Data Quality Flag", "Task Count"], ["No findings", 0]])
    coverage_start = quality.max_row + 2
    _write_excel_table(quality, [["Source", "Space", "Project", "Tasks", "History Coverage", "Notes"]] + [
        [item.source_tool, item.source_space, item.unified_project, item.task_count, item.history_mode, item.reason or "N/A"]
        for item in model.source_coverage
    ], start_row=coverage_start)
    _fit_columns(quality)

    context = workbook.create_sheet("Analysis Context")
    _write_excel_table(context, [["Field", "Value"],
        ["Analysis Level", "Company"],
        ["Source Systems", "Jira and ClickUp"],
        ["Analysis Period", f"{model.period_start.isoformat()} to {model.period_end.isoformat()}"],
        ["Task Scope", "All collected source Spaces unified into Company Performance."],
        ["Deduplication", "Distinct Source Tool + Task ID."],
        ["Subtasks", "Visible in Task Details; excluded from KPI rates."],
    ])
    _fit_columns(context)

    definitions = workbook.create_sheet("Metric Definitions")
    _write_excel_table(definitions, [["Metric", "Definition"],
        ["Total Tasks", "All tasks in the selected Company scope; subtasks remain visible."],
        ["Completed", "Tasks with normalized Completed status at period end."],
        ["Completion Rate", "Completed tasks divided by KPI-counted tasks with a verified status at period end; Unknown statuses are excluded and shown in Data Quality."],
        ["On-Time Rate", "Completed on or before due date divided by completed tasks with a known due date."],
        ["Open Overdue", "Open tasks with a due date earlier than the period end."],
        ["WIP", "Open tasks currently in In Execution or In Review at period end; On Hold and At Risk are not WIP."],
        ["Average Lead Time", "Average completion date minus creation date for completed KPI-counted tasks."],
        ["Average Execution Time", "Average completion date minus actual start date for completed tasks where both dates are known."],
        ["Workflow Exceptions", "Recorded rework, replanning, re-evaluation, or reopen evidence."],
    ])
    _fit_columns(definitions)

    for sheet in workbook.worksheets:
        sheet.sheet_view.showGridLines = False
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
        "Completion Rate": "Completed tasks divided by KPI-counted tasks with a verified status at period end; Unknown statuses are excluded and shown in Data Quality.",
        "On-Time Rate": "Completed tasks finished on or before their due date.",
        "On-Time Completion Rate": "Completed tasks finished on or before their due date.",
        "Open Overdue": "Open tasks with a due date before period end.",
        "Overdue Open Tasks": "Open tasks with a due date before period end.",
        "Current WIP": "Tasks in In Execution or In Review at period end; On Hold and At Risk are not WIP.",
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
    """Create the exact Company Performance Word report requested by the user."""
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    snapshots = tuple(snapshots)
    bottlenecks = tuple(bottlenecks)
    recommendations = tuple(recommendations)
    rows = _task_rows(model, snapshots)
    document = Document()
    _configure_document(document)
    title = document.add_heading("Company Performance Report", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle = document.add_paragraph("Company-wide Jira and ClickUp analysis")
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_table(document, ("Report Field", "Value"), [
        ("Analysis Level", "Company"),
        ("Data Sources", "Jira and ClickUp"),
        ("Analysis Period", f"{model.period_start.isoformat()} to {model.period_end.isoformat()}"),
        ("Tasks in Scope", model.in_period_task_count),
    ])

    document.add_heading("Company Executive Summary", level=1)
    document.add_paragraph(
        f"The analysis covers {model.in_period_task_count} task(s) from the selected Jira projects and ClickUp Spaces. "
        f"{model.kpis.completed_tasks} task(s) are completed, completion rate is {_display(model.kpis.completion_rate)}, "
        f"on-time rate is {_display(model.kpis.on_time_completion_rate)}, WIP is {model.kpis.current_wip}, "
        f"and {model.kpis.overdue_open_tasks} open overdue task(s) need attention."
    )

    document.add_heading("Data Sources and Analysis Scope", level=1)
    _add_table(document, ("Source", "Space", "Project", "Tasks", "History Coverage", "Notes"), [
        (item.source_tool, item.source_space, item.unified_project, item.task_count,
         item.history_mode, item.reason or "N/A")
        for item in model.source_coverage
    ])
    document.add_paragraph(
        "The Company view combines Jira and ClickUp records, removes duplicate Source Tool + Task ID pairs, "
        "and keeps original source and Space identifiers available in Excel Task Details."
    )

    document.add_heading("Company KPI Summary", level=1)
    _add_table(document, ("KPI", "Value"), [
        ("Total Tasks", model.kpis.total_tasks),
        ("Completed", model.kpis.completed_tasks),
        ("Completion Rate", _display(model.kpis.completion_rate)),
        ("On-Time Rate", _display(model.kpis.on_time_completion_rate)),
        ("Open Overdue", model.kpis.overdue_open_tasks),
        ("WIP", model.kpis.current_wip),
        ("Average Lead Time", _display(model.kpis.average_lead_time_days)),
        ("Average Execution Time", _display(model.kpis.average_execution_duration_days)),
    ])

    counted = [item for item in snapshots if item.counted_in_kpis]
    projects: dict[str, list[TaskPeriodSnapshot]] = defaultdict(list)
    departments: dict[str, list[TaskPeriodSnapshot]] = defaultdict(list)
    for item in counted:
        projects[item.task.unified_project or "Unknown"].append(item)
        departments[item.task.department or "Unknown"].append(item)

    document.add_heading("Department Performance Comparison", level=1)
    department_rows = []
    for name, items in sorted(departments.items()):
        completed = [item for item in items if item.status_at_period_end.value == "Completed"]
        with_due = [item for item in completed if item.task.due_date and item.final_completion_date]
        on_time = [item for item in with_due if item.final_completion_date <= item.task.due_date]
        overdue = [
            item for item in items
            if item.status_at_period_end.is_open and item.task.due_date and item.task.due_date < item.period_end
        ]
        department_rows.append((
            name, len(items), len(completed),
            f"{len(completed) / len(items) * 100.0:.1f}%" if items else "N/A",
            f"{len(on_time) / len(with_due) * 100.0:.1f}%" if with_due else "N/A",
            len(overdue),
        ))
    _add_table(document, ("Department", "Total Tasks", "Completed", "Completion Rate", "On-Time Rate", "Open Overdue"), department_rows or [
        ("Unknown", 0, 0, "N/A", "N/A", 0)
    ])

    document.add_heading("Company-wide Bottlenecks", level=1)
    _add_table(document, ("Stage", "Assessment", "Evidence", "Average Days", "Open Tasks", "Open Overdue"), [
        (item.status.value, item.strength, "; ".join(item.evidence),
         item.metrics.average_days, item.metrics.open_tasks_now, item.metrics.overdue_open_tasks)
        for item in bottlenecks
    ] or [("N/A", "No candidate", "No bottleneck candidate was identified.", "N/A", 0, 0)])

    overdue_rows = [
        (row["Unified Project"], row["Task ID"], row["Task Name"], row["Source Tool"],
         row["Source Space"], row["Assignee"], row["Due Date"])
        for row in rows
        if row["Counted in KPIs"] and row["Final Status"] not in {"Completed", "Cancelled", "Rejected"}
        and row["Due Date"] is not None and row["Due Date"] < model.period_end
    ]
    document.add_heading("Key Risks and Overdue Tasks", level=1)
    _add_table(document, ("Project", "Task ID", "Task", "Source", "Space", "Assignee", "Due Date"), overdue_rows or [
        ("N/A", "N/A", "No open overdue tasks", "", "", "", "")
    ])

    document.add_heading("Workflow Exceptions", level=1)
    exception_rows = [
        (row["Unified Project"], row["Task ID"], row["Task Name"], row["Source Tool"], event)
        for row in rows for event in row["Exception Events"] or ()
    ]
    _add_table(document, ("Project", "Task ID", "Task", "Source", "Exception"), exception_rows or [
        ("N/A", "N/A", "No workflow exceptions", "", "")
    ])

    document.add_heading("Top Departments Requiring Attention", level=1)
    attention_rows = [
        row for row in department_rows
        if row[5] > 0 or row[3] == "N/A" or float(row[3].rstrip("%")) < 70.0
    ]
    _add_table(document, ("Department", "Total Tasks", "Completed", "Completion Rate", "On-Time Rate", "Open Overdue"), attention_rows or [
        ("N/A", 0, 0, "N/A", "N/A", 0)
    ])

    document.add_heading("Executive Recommendations", level=1)
    _add_table(document, ("Severity", "Recommendation", "Evidence", "Suggested Action"), [
        (item.severity, item.title, item.evidence, item.suggested_action)
        for item in recommendations
    ] or [("N/A", "No recommendations generated", "", "")])

    document.add_heading("Data Quality and Coverage Limitations", level=1)
    _add_table(document, ("Data Quality Flag", "Task Count"), [
        (item.flag, item.task_count) for item in model.data_quality
    ] or [("No findings", 0)])
    document.add_paragraph(
        "History-dependent workflow metrics remain unavailable where the source does not provide complete history. "
        "Department labels are shown from the source record when available; records without a reliable department "
        "label remain under Unknown. Unavailable values are not converted to zero."
    )

    document.add_heading("Metric Definitions", level=1)
    _add_table(document, ("Metric", "Definition"), [
        ("Total Tasks", "All tasks in the selected Company scope; subtasks remain visible."),
        ("Completed", "Tasks with normalized Completed status at period end."),
        ("Completion Rate", "Completed tasks divided by KPI-counted tasks with a verified status at period end; Unknown statuses are excluded and shown in Data Quality."),
        ("On-Time Rate", "Completed on or before due date divided by completed tasks with a known due date."),
        ("Open Overdue", "Open tasks with a due date earlier than the analysis period end."),
        ("WIP", "Open tasks currently in In Execution or In Review at period end; On Hold and At Risk are not WIP."),
        ("Average Lead Time", "Average completion date minus creation date for completed KPI-counted tasks."),
        ("Average Execution Time", "Average completion date minus actual start date for completed tasks with both dates known."),
        ("Workflow Exceptions", "Recorded rework, replanning, re-evaluation, or reopen evidence."),
    ])

    document.save(output)
    return output
