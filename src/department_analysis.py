"""Department-specific presentation and exports built on normalized ClickUp results."""
from __future__ import annotations

from io import BytesIO
from typing import Any

import pandas as pd
from docx import Document
from docx.shared import Inches
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Font, PatternFill


def _metric(result: dict, name: str, default=None):
    matches = result["overall"].loc[result["overall"]["Metric"].eq(name), "Value"]
    return default if matches.empty else matches.iloc[0]


def _rate(numerator: int, denominator: int):
    return None if not denominator else numerator / denominator * 100.0


def build_department_result(clickup_result: dict, prepared_data) -> dict:
    """Create the isolated department view without changing legacy metrics."""
    tasks = clickup_result["tasks"].copy()
    total = len(tasks)
    is_subtask = tasks.get("Is Subtask", pd.Series(False, index=tasks.index)).fillna(False).astype(bool)
    parent_tasks = tasks.loc[~is_subtask].copy()
    cancelled = int(parent_tasks["Cancelled?"].sum()) if not parent_tasks.empty else 0
    eligible = len(parent_tasks) - cancelled
    completed = int(parent_tasks["Completed?"].sum()) if not parent_tasks.empty else 0
    open_tasks = int(tasks["Open?"].sum()) if total else 0
    overdue = int((tasks["Open?"] & tasks["Overdue Days"].notna()).sum()) if total else 0
    due_known = (
        parent_tasks["Completed?"] & parent_tasks["Due Date"].notna() & parent_tasks["Completed"].notna()
        if not parent_tasks.empty else pd.Series(dtype=bool)
    )
    on_time = int((due_known & parent_tasks["On Time?"].eq(True)).sum()) if not parent_tasks.empty else 0

    kpis = pd.DataFrame([
        ("Total Tasks", total, total, total),
        ("Task Completion Rate (%)", _rate(completed, eligible), completed, eligible),
        ("On-Time Completion Rate (%)", _rate(on_time, int(due_known.sum())), on_time, int(due_known.sum())),
        ("Open Overdue Tasks", overdue, overdue, open_tasks),
        ("Open Overdue Rate (%)", _rate(overdue, open_tasks), overdue, open_tasks),
        ("Average Lead Time (hours)", pd.to_numeric(tasks.loc[tasks["Completed?"], "Lead Time Hours"], errors="coerce").mean(), completed, completed),
        ("Workflow Exception Rate (%)", None, 0, 0),
        ("Cancelled/Rejected Tasks", cancelled, cancelled, total),
        ("Cancellation/Rejected Rate (%)", _rate(cancelled, total), cancelled, total),
    ], columns=["KPI", "Value", "Numerator", "Denominator"])

    employees = clickup_result["assignee_summary"].copy()
    if not employees.empty:
        employees = employees.rename(columns={"Total Tasks": "Total Assigned"})
        overdue_by_employee = (tasks.assign(_overdue=tasks["Open?"] & tasks["Overdue Days"].notna())
                               .groupby("Assignee", dropna=False)["_overdue"].sum())
        employees["Overdue"] = employees["Assignee"].map(overdue_by_employee).fillna(0).astype(int)
        employees["Workflow Exceptions"] = None

    attention_mask = pd.Series(False, index=tasks.index)
    if total:
        priority = tasks["Priority"].fillna("").astype(str).str.casefold()
        attention_mask = (
            (tasks["Open?"] & tasks["Overdue Days"].notna())
            | tasks["Assignee"].fillna("Unassigned").eq("Unassigned")
            | tasks["Due Date"].isna()
            | (tasks["Open?"] & priority.isin({"urgent", "high"}) & tasks["Overdue Days"].notna())
        )
    attention = tasks.loc[attention_mask].copy()
    if not attention.empty:
        def issue(row):
            issues = []
            if row["Open?"] and pd.notna(row["Overdue Days"]): issues.append("Open overdue")
            if str(row.get("Priority", "")).casefold() in {"urgent", "high"} and row["Open?"] and pd.notna(row["Overdue Days"]): issues.append("High-priority overdue")
            if row.get("Assignee") in (None, "", "Unassigned"): issues.append("Unassigned")
            if pd.isna(row.get("Due Date")): issues.append("Missing due date")
            return "; ".join(issues)
        attention["Issue"] = attention.apply(issue, axis=1)

    bottlenecks = clickup_result["status_summary"].copy()
    if not bottlenecks.empty:
        bottlenecks = bottlenecks.sort_values(["Total Hours", "Tasks"], ascending=False).head(5)
        bottlenecks["Interpretation"] = "Bottleneck candidate; review workload and blockers before concluding cause."

    quality_rows = []
    for row in clickup_result["quality"].to_dict("records"):
        if row.get("Status") != "OK":
            quality_rows.append({"Issue Type": row["Check"], "Count": row["Value"], "Analysis Impact": row["Status"]})
    quality = pd.DataFrame(quality_rows, columns=["Issue Type", "Count", "Analysis Impact"])
    return {
        **clickup_result,
        "analysis_type": "Department Performance",
        "department_name": prepared_data.department_name,
        "department_id": prepared_data.department_id,
        "kpis": kpis,
        "employee_breakdown": employees,
        "attention": attention,
        "bottlenecks": bottlenecks,
        "department_quality": quality,
        "space_names": getattr(prepared_data, "space_names", [prepared_data.space_name]),
        "data_source": "ClickUp",
    }


def build_jira_department_result(task_metrics: pd.DataFrame, process_data: dict, space_name: str, cutoff) -> dict:
    """Adapt verified Jira metrics to the same department presentation contract."""
    source = task_metrics.copy()
    tasks = pd.DataFrame({
        "Task ID": source.get("issue_key"), "Task Name": source.get("task_name"),
        "Assignee": source.get("assignee_name", pd.Series("Unassigned", index=source.index)).fillna("Unassigned"),
        "Priority": source.get("priority"), "Task Type": source.get("issue_type"),
        "Current Status": source.get("status_at_cutoff"), "Created": source.get("created_at"),
        "Start Date": source.get("actual_start_at"), "Due Date": source.get("due_date"),
        "Completed": source.get("completed_at"), "Completed?": source.get("is_completed", False),
        "Cancelled?": source.get("is_rejected", False), "Open?": source.get("is_open", False),
        "WIP?": source.get("is_wip", False), "Lead Time Hours": source.get("lead_time_business_hours"),
        "Execution Hours": source.get("execution_business_hours"),
        "On Time?": source.get("on_time_completion"), "Overdue Days": source.get("overdue_days"),
    })
    total = len(tasks); cancelled = int(tasks["Cancelled?"].eq(True).sum()); eligible = total - cancelled
    completed = int(tasks["Completed?"].eq(True).sum()); open_count = int(tasks["Open?"].eq(True).sum())
    overdue = int((tasks["Open?"].eq(True) & pd.to_numeric(tasks["Overdue Days"], errors="coerce").gt(0)).sum())
    due_known = tasks["Completed?"].eq(True) & tasks["Due Date"].notna() & tasks["Completed"].notna()
    on_time = int((due_known & tasks["On Time?"].eq(True)).sum())
    history_valid = source.get("history_complete", pd.Series(False, index=source.index)).eq(True)
    exception = (
        source.get("rework_count", pd.Series(0, index=source.index)).fillna(0).gt(0)
        | source.get("replanning_count", pd.Series(0, index=source.index)).fillna(0).gt(0)
        | source.get("re_evaluation_count", pd.Series(0, index=source.index)).fillna(0).gt(0)
    )
    exception_count = int((history_valid & exception).sum()); exception_denominator = int(history_valid.sum())
    kpis = pd.DataFrame([
        ("Total Tasks", total, total, total),
        ("Task Completion Rate (%)", _rate(completed, eligible), completed, eligible),
        ("On-Time Completion Rate (%)", _rate(on_time, int(due_known.sum())), on_time, int(due_known.sum())),
        ("Open Overdue Tasks", overdue, overdue, open_count),
        ("Open Overdue Rate (%)", _rate(overdue, open_count), overdue, open_count),
        ("Average Lead Time (hours)", pd.to_numeric(tasks.loc[tasks["Completed?"].eq(True), "Lead Time Hours"], errors="coerce").mean(), completed, completed),
        ("Workflow Exception Rate (%)", _rate(exception_count, exception_denominator), exception_count, exception_denominator),
        ("Cancelled/Rejected Tasks", cancelled, cancelled, total),
        ("Cancellation/Rejected Rate (%)", _rate(cancelled, total), cancelled, total),
    ], columns=["KPI", "Value", "Numerator", "Denominator"])
    employees = (tasks.groupby("Assignee", dropna=False)
                 .agg(**{"Total Assigned": ("Task ID", "count"), "Completed": ("Completed?", "sum"),
                         "Open": ("Open?", "sum"), "WIP": ("WIP?", "sum")}).reset_index())
    overdue_map = tasks.assign(_overdue=tasks["Open?"].eq(True) & pd.to_numeric(tasks["Overdue Days"], errors="coerce").gt(0)).groupby("Assignee")["_overdue"].sum()
    employees["Overdue"] = employees["Assignee"].map(overdue_map).fillna(0).astype(int)
    employees["Workflow Exceptions"] = source.assign(_assignee=tasks["Assignee"], _exception=exception).groupby("_assignee")["_exception"].sum().reindex(employees["Assignee"]).to_numpy()
    attention_mask = tasks["Open?"].eq(True) & pd.to_numeric(tasks["Overdue Days"], errors="coerce").gt(0)
    attention_mask |= tasks["Assignee"].eq("Unassigned") | tasks["Due Date"].isna() | exception
    attention = tasks.loc[attention_mask].copy()
    if not attention.empty:
        attention["Issue"] = ["; ".join(filter(None, [
            "Open overdue" if bool(row["Open?"]) and pd.notna(row["Overdue Days"]) and row["Overdue Days"] > 0 else "",
            "Unassigned" if row["Assignee"] == "Unassigned" else "",
            "Missing due date" if pd.isna(row["Due Date"]) else "",
            "Workflow exception" if bool(exception.loc[index]) else "",
        ])) for index, row in attention.iterrows()]
    stage = process_data.get("stage_summary", pd.DataFrame()).copy()
    bottlenecks = stage.sort_values("total_elapsed_hours", ascending=False).head(5) if "total_elapsed_hours" in stage else stage.head(5)
    if not bottlenecks.empty: bottlenecks["Interpretation"] = "Bottleneck candidate; validate cause with the department."
    status_counts = tasks["Current Status"].fillna("Unavailable").value_counts().rename_axis("Status").reset_index(name="Tasks")
    delivery = pd.DataFrame({"Due Variance Category": ["Completed on time", "Completed late", "Open overdue"],
                             "Tasks": [on_time, int((due_known & tasks["On Time?"].eq(False)).sum()), overdue]})
    quality = process_data.get("data_quality", pd.DataFrame()).copy()
    if not quality.empty:
        quality = quality.rename(
            columns={"issue_key": "Issue Type", "finding": "Analysis Impact"}
        )
        quality["Count"] = 1
        quality = quality[["Issue Type", "Count", "Analysis Impact"]]
    else:
        quality = pd.DataFrame(
            columns=["Issue Type", "Count", "Analysis Impact"]
        )
    return {
        "tasks": tasks, "kpis": kpis, "employee_breakdown": employees, "attention": attention,
        "bottlenecks": bottlenecks, "department_quality": quality, "status_counts": status_counts,
        "due_variance_summary": delivery, "weekly_flow": process_data.get("weekly_flow", pd.DataFrame()).rename(
            columns={"week_start":"Week Starting", "tasks_opened":"Tasks Created", "tasks_completed":"Tasks Completed"}),
        "department_name": "Tech Development", "department_id": "jira-tech-development",
        "space_name": space_name, "cutoff": cutoff, "data_source": "Jira",
    }


def _write_frame(ws, frame: pd.DataFrame, start_row=1, title=None):
    row = start_row
    if title:
        ws.cell(row, 1, title).font = Font(bold=True, size=14, color="17324D")
        row += 2
    for column, value in enumerate(frame.columns, 1):
        cell = ws.cell(row, column, value)
        cell.fill = PatternFill("solid", fgColor="17324D")
        cell.font = Font(color="FFFFFF", bold=True)
    for values in frame.itertuples(index=False, name=None):
        row += 1
        for column, value in enumerate(values, 1):
            if isinstance(value, pd.Timestamp): value = value.isoformat()
            if pd.isna(value) if not isinstance(value, (list, dict)) else False: value = None
            ws.cell(row, column, value)
    return row


def department_excel(result: dict) -> bytes:
    """Return the approved four-sheet department workbook."""
    wb = Workbook()
    wb.remove(wb.active)
    summary = wb.create_sheet("Department Summary")
    info = pd.DataFrame([
        ("Department", result["department_name"]),
        ("Source Spaces", ", ".join(map(str, result.get("space_names", [result["space_name"]])))),
        ("Data Source", result.get("data_source", "ClickUp")), ("Analysis Period End", result["cutoff"]),
        ("Analyzed Tasks", len(result["tasks"])),
    ], columns=["Field", "Value"])
    end = _write_frame(summary, info, title="Department Performance Report")
    _write_frame(summary, result["kpis"], start_row=end + 3, title="KPI Summary")
    _write_frame(wb.create_sheet("Employee Breakdown"), result["employee_breakdown"])
    _write_frame(wb.create_sheet("Task Details"), result["tasks"])
    exceptions = wb.create_sheet("Exceptions & Data Quality")
    attention_columns = [c for c in ["Task ID", "Task Name", "Assignee", "Current Status", "Priority", "Due Date", "Issue"] if c in result["attention"]]
    end = _write_frame(exceptions, result["attention"][attention_columns], title="Workflow Exceptions and Tasks Requiring Attention")
    end = _write_frame(exceptions, result["bottlenecks"], start_row=end + 3, title="Bottleneck Candidates")
    _write_frame(exceptions, result["department_quality"], start_row=end + 3, title="Data Quality Issues")
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        for column_cells in ws.columns:
            ws.column_dimensions[column_cells[0].column_letter].width = min(max(len(str(c.value or "")) for c in column_cells) + 2, 45)
    stream = BytesIO(); wb.save(stream); return stream.getvalue()


def _display(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)): return "Unavailable"
    if isinstance(value, float): return f"{value:.1f}"
    return str(value)


def department_word(result: dict) -> bytes:
    document = Document()
    section = document.sections[0]
    section.top_margin = section.bottom_margin = Inches(0.75)
    document.add_heading("Department Performance Evaluation Report", 0)
    source_spaces = ", ".join(map(str, result.get("space_names", [result["space_name"]])))
    document.add_paragraph(f"Department: {result['department_name']} | {result.get('data_source', 'ClickUp')} Source Spaces: {source_spaces} | Cutoff: {result['cutoff']}")
    document.add_heading("Executive Summary", level=1)
    completion = _metric({"overall": result["kpis"].rename(columns={"KPI":"Metric"})}, "Task Completion Rate (%)")
    overdue = _metric({"overall": result["kpis"].rename(columns={"KPI":"Metric"})}, "Open Overdue Tasks", 0)
    document.add_paragraph(f"The department scope contains {len(result['tasks'])} tasks. Completion rate is {_display(completion)}% and {int(overdue or 0)} open overdue task(s) require review.")
    for heading, frame, columns in [
        ("Department KPI Summary", result["kpis"], ["KPI", "Value", "Numerator", "Denominator"]),
        ("Employee Workload Summary", result["employee_breakdown"], ["Assignee", "Total Assigned", "Completed", "Open", "WIP", "Overdue"]),
        ("Bottleneck Candidates", result["bottlenecks"], ["Status", "Tasks", "Average Hours", "Total Hours", "Interpretation"]),
        ("Tasks Requiring Attention", result["attention"], ["Task ID", "Task Name", "Assignee", "Current Status", "Priority", "Issue"]),
        ("Data Quality Notes", result["department_quality"], ["Issue Type", "Count", "Analysis Impact"]),
    ]:
        document.add_heading(heading, level=1)
        visible = frame[[c for c in columns if c in frame]].copy()
        if visible.empty:
            document.add_paragraph("No items were identified in this section.")
            continue
        table = document.add_table(rows=1, cols=len(visible.columns)); table.style = "Table Grid"
        for i, name in enumerate(visible.columns): table.rows[0].cells[i].text = name
        for values in visible.itertuples(index=False, name=None):
            cells = table.add_row().cells
            for i, value in enumerate(values): cells[i].text = _display(value)
    document.add_heading("Recommendations", level=1)
    if int(overdue or 0): document.add_paragraph("Prioritize open overdue tasks, confirm blockers and owners, and agree recovery dates.", style="List Bullet")
    if not result["employee_breakdown"].empty: document.add_paragraph("Review workload distribution before reassigning work; task counts alone do not measure individual contribution.", style="List Bullet")
    document.add_paragraph("Treat listed bottlenecks as candidates requiring management validation, not confirmed root causes.", style="List Bullet")
    stream = BytesIO(); document.save(stream); return stream.getvalue()
