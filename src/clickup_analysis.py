"""ClickUp-only performance analysis built from task and time-in-status data.

This module intentionally has no Jira imports. It evaluates the ClickUp task
snapshot that was selected and collected, while keeping unsupported historical
metrics unavailable instead of inventing them.
"""
from __future__ import annotations

import json
from io import BytesIO
from typing import Any

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

from clickup_filters import (
    CANCELLED_STATUSES,
    TERMINAL_STATUSES,
    assignees,
    created_by,
    is_closed,
    location_values,
    milliseconds_to_hours,
    status_name,
    status_type,
    tags,
    task_type,
    timestamp,
)


NOT_STARTED_STATUSES = {
    "backlog",
    "to do",
    "todo",
    "not started",
    "open",
    "planning",
    "ready",
}


def _status_minutes(payload: Any) -> dict[str, float]:
    """Flatten native ClickUp Total time in Status data by status."""
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
        values[status] = max(minutes, values.get(status, 0.0))
    return values


def _duration_hours(start, end):
    if pd.isna(start) or pd.isna(end):
        return None
    return (end - start).total_seconds() / 3600.0


def _timing_values(created, start, end, end_label: str):
    """Return trustworthy task-timing values without inventing negative time.

    ClickUp start dates are frequently date-only values.  A same-day start
    timestamp can therefore precede the creation timestamp by a few hours.
    In that case the effective start is the creation time.  A start date after
    the completion/cutoff, however, is a chronology conflict and the derived
    timing metrics must remain unavailable.
    """
    if pd.isna(start):
        return None, None, "No start date"
    if pd.notna(end) and start > end:
        return None, None, f"Start date after {end_label}"
    if pd.notna(created) and start.normalize() < created.normalize():
        return None, None, "Start date before creation date"

    effective_start = max(start, created) if pd.notna(created) and start < created else start
    return (
        _duration_hours(effective_start, end),
        _duration_hours(created, effective_start),
        "Available",
    )


def _first_location(task: dict) -> str:
    values = location_values(task)
    return values[-1] if values else "Unavailable"


def _due_category(completed: bool, open_task: bool, due, variance):
    if pd.isna(due) or variance is None or pd.isna(variance):
        return "No due date"
    if completed:
        if variance > 0:
            return "Completed late"
        if variance < 0:
            return "Completed early"
        return "Completed on time"
    if open_task:
        return "Open overdue" if variance > 0 else "Open not due"
    return "Not applicable"


def _safe_rate(values: pd.Series):
    usable = values.dropna()
    return None if usable.empty else float(usable.mean() * 100.0)


def _mean(frame: pd.DataFrame, column: str, mask=None):
    if frame.empty or column not in frame:
        return None
    values = pd.to_numeric(frame.loc[mask, column] if mask is not None else frame[column], errors="coerce").dropna()
    return None if values.empty else float(values.mean())


def _weekly_flow(task_frame: pd.DataFrame) -> pd.DataFrame:
    columns = ["Week Starting", "Tasks Created", "Tasks Completed", "Net Flow", "Cumulative Net Flow"]
    if task_frame.empty:
        return pd.DataFrame(columns=columns)
    created = task_frame.loc[task_frame["Created"].notna(), ["Created"]].copy()
    completed = task_frame.loc[task_frame["Completed"].notna(), ["Completed"]].copy()
    created["Week Starting"] = created["Created"].dt.normalize() - pd.to_timedelta(created["Created"].dt.weekday, unit="D")
    completed["Week Starting"] = completed["Completed"].dt.normalize() - pd.to_timedelta(completed["Completed"].dt.weekday, unit="D")
    created_counts = created.groupby("Week Starting").size().rename("Tasks Created")
    completed_counts = completed.groupby("Week Starting").size().rename("Tasks Completed")
    output = pd.concat([created_counts, completed_counts], axis=1).fillna(0).reset_index()
    if output.empty:
        return pd.DataFrame(columns=columns)
    output[["Tasks Created", "Tasks Completed"]] = output[["Tasks Created", "Tasks Completed"]].astype(int)
    output["Net Flow"] = output["Tasks Created"] - output["Tasks Completed"]
    output["Cumulative Net Flow"] = output["Net Flow"].cumsum()
    return output.sort_values("Week Starting").reset_index(drop=True)


def _due_status_summary(frame: pd.DataFrame, cutoff) -> pd.DataFrame:
    columns = ["Due Status", "Tasks", "Share of Open Tasks (%)"]
    if frame.empty:
        return pd.DataFrame(columns=columns)
    open_tasks = frame[frame["Open?"]].copy()
    if open_tasks.empty:
        return pd.DataFrame(columns=columns)

    def bucket(row):
        due = row["Due Date"]
        if pd.isna(due):
            return "No due date"
        if pd.notna(cutoff) and due < cutoff:
            return "Overdue"
        if pd.notna(cutoff) and due <= cutoff + pd.Timedelta(days=7):
            return "Due in next 7 days"
        return "Due later"

    output = open_tasks.assign(**{"Due Status": open_tasks.apply(bucket, axis=1)}).groupby("Due Status").size().rename("Tasks").reset_index()
    output["Share of Open Tasks (%)"] = output["Tasks"] / len(open_tasks) * 100.0
    return output.sort_values("Tasks", ascending=False).reset_index(drop=True)


def _assignee_summary(frame: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "Assignee", "Total Tasks", "Completed", "Open", "WIP", "Cancelled",
        "Completion Rate (%)", "On-Time Completion Rate (%)", "Completed Late",
        "Open Overdue", "Average Due Variance (days)",
    ]
    if frame.empty:
        return pd.DataFrame(columns=columns)
    rows = []
    for assignee, group in frame.groupby("Assignee", dropna=False):
        completed = group[group["Completed?"]]
        rows.append({
            "Assignee": assignee or "Unassigned",
            "Total Tasks": int(len(group)),
            "Completed": int(group["Completed?"].sum()),
            "Open": int(group["Open?"].sum()),
            "WIP": int(group["WIP?"].sum()),
            "Cancelled": int(group["Cancelled?"].sum()),
            "Completion Rate (%)": float(group["Completed?"].mean() * 100.0) if len(group) else None,
            "On-Time Completion Rate (%)": _safe_rate(completed["On Time?"]) if not completed.empty else None,
            "Completed Late": int((completed["Due Variance (days)"] > 0).sum()) if not completed.empty else 0,
            "Open Overdue": int(((group["Open?"]) & (group["Due Variance (days)"] > 0)).sum()),
            "Average Due Variance (days)": _mean(completed, "Due Variance (days)"),
        })
    return pd.DataFrame(rows, columns=columns).sort_values(["Total Tasks", "Assignee"], ascending=[False, True]).reset_index(drop=True)


def _status_summary(status_frame: pd.DataFrame) -> pd.DataFrame:
    columns = ["Status", "Tasks", "Average Hours", "Median Hours", "Total Hours"]
    if status_frame.empty:
        return pd.DataFrame(columns=columns)
    return (status_frame.groupby("Status", dropna=False)
            .agg(Tasks=("Task ID", "nunique"), Average_Hours=("Hours", "mean"),
                 Median_Hours=("Hours", "median"), Total_Hours=("Hours", "sum"))
            .reset_index()
            .rename(columns={"Average_Hours": "Average Hours", "Median_Hours": "Median Hours", "Total_Hours": "Total Hours"})
            .sort_values("Total Hours", ascending=False)
            .reset_index(drop=True))


def analyze_clickup(prepared_data):
    """Return a ClickUp-native analysis payload for the Streamlit dashboard."""
    payload = json.loads(prepared_data.history_json.decode("utf-8"))
    raw_tasks = payload.get("tasks") if isinstance(payload, dict) else []
    time_payloads = payload.get("time_in_status", {}) if isinstance(payload, dict) else {}
    cutoff = timestamp(prepared_data.cutoff)
    rows: list[dict[str, Any]] = []
    status_rows: list[dict[str, Any]] = []

    for task in raw_tasks if isinstance(raw_tasks, list) else []:
        task_id = str(task.get("id", ""))
        status = status_name(task)
        status_key = status.casefold()
        state = status_type(task)
        created = timestamp(task.get("date_created"))
        updated = timestamp(task.get("date_updated"))
        start = timestamp(task.get("start_date"))
        due = timestamp(task.get("due_date"))
        completed = timestamp(task.get("date_closed") or task.get("date_done"))
        cancelled = status_key in CANCELLED_STATUSES
        completed_flag = (status_key in TERMINAL_STATUSES or is_closed(task) or pd.notna(completed)) and not cancelled
        open_flag = not completed_flag and not cancelled
        wip_flag = open_flag and status_key not in NOT_STARTED_STATUSES
        end = completed if pd.notna(completed) else cutoff
        lead_time_hours = _duration_hours(created, end)
        timing_end = completed if completed_flag else cutoff
        timing_label = "completion date" if completed_flag else "evaluation cutoff"
        execution_hours, time_to_start_hours, timing_data_status = _timing_values(
            created, start, timing_end, timing_label
        )
        due_variance_days = None
        due_basis = "Unavailable"
        if pd.notna(due) and pd.notna(completed) and completed_flag:
            due_variance_days = (completed - due).total_seconds() / 86400.0
            due_basis = "Completed date vs due date"
        elif pd.notna(due) and open_flag and pd.notna(cutoff):
            due_variance_days = (cutoff - due).total_seconds() / 86400.0
            due_basis = "Evaluation cutoff vs due date"
        on_time = bool(completed <= due) if completed_flag and pd.notna(completed) and pd.notna(due) else None
        overdue_days = due_variance_days if open_flag and due_variance_days is not None and due_variance_days > 0 else None

        time_payload = time_payloads.get(task_id) if isinstance(time_payloads, dict) else None
        minutes_by_status = _status_minutes(time_payload)
        current_minutes = None
        current_since = None
        if isinstance(time_payload, dict) and isinstance(time_payload.get("current_status"), dict):
            current = time_payload["current_status"]
            current_minutes = (current.get("total_time") or {}).get("by_minute")
            current_since = (current.get("total_time") or {}).get("since")
            try:
                current_minutes = float(current_minutes)
            except (TypeError, ValueError):
                current_minutes = None

        row = {
            "Task ID": task_id,
            "Task Name": task.get("name"),
            "Assignee": ", ".join(assignees(task)),
            "Created By": created_by(task),
            "Priority": (task.get("priority") or {}).get("priority") if isinstance(task.get("priority"), dict) else task.get("priority"),
            "Task Type": task_type(task),
            "Tags": ", ".join(tags(task)) or "No tags",
            "Location/List": _first_location(task),
            "Current Status": status,
            "Status State": state or "Unavailable",
            "Created": created,
            "Updated": updated,
            "Start Date": start,
            "Due Date": due,
            "Completed": completed,
            "Completed?": completed_flag,
            "Cancelled?": cancelled,
            "Open?": open_flag,
            "WIP?": wip_flag,
            "Elapsed Hours": lead_time_hours,
            "Lead Time Hours": lead_time_hours,
            "Execution Hours": execution_hours,
            "Time to Start Hours": time_to_start_hours,
            "Timing Data Status": timing_data_status,
            "On Time?": on_time,
            "Due Variance (days)": due_variance_days,
            "Due Variance Basis": due_basis,
            "Due Variance Category": _due_category(completed_flag, open_flag, due, due_variance_days),
            "Overdue Days": overdue_days,
            "Time Estimate Hours": milliseconds_to_hours(task.get("time_estimate")),
            "Time Tracked Hours": milliseconds_to_hours(task.get("time_spent")),
            "Current Status Time (min)": current_minutes,
            "Current Status Since": timestamp(current_since),
            "Total Time in Status (min)": sum(minutes_by_status.values()) if minutes_by_status else None,
            "Time in Status": json.dumps(minutes_by_status, ensure_ascii=False) if minutes_by_status else None,
        }
        rows.append(row)
        for status_label, minutes in minutes_by_status.items():
            status_rows.append({"Status": status_label, "Task ID": task_id, "Minutes": minutes, "Hours": minutes / 60.0})

    columns = [
        "Task ID", "Task Name", "Assignee", "Created By", "Priority", "Task Type", "Tags", "Location/List",
        "Current Status", "Status State", "Created", "Updated", "Start Date", "Due Date", "Completed",
        "Completed?", "Cancelled?", "Open?", "WIP?", "Elapsed Hours", "Lead Time Hours", "Execution Hours",
        "Time to Start Hours", "Timing Data Status", "On Time?", "Due Variance (days)", "Due Variance Basis", "Due Variance Category",
        "Overdue Days", "Time Estimate Hours", "Time Tracked Hours", "Current Status Time (min)",
        "Current Status Since", "Total Time in Status (min)", "Time in Status",
    ]
    task_frame = pd.DataFrame(rows, columns=columns)
    status_frame = pd.DataFrame(status_rows, columns=["Status", "Task ID", "Minutes", "Hours"])
    total = len(task_frame)
    completed = task_frame[task_frame["Completed?"]] if total else task_frame
    open_tasks = task_frame[task_frame["Open?"]] if total else task_frame
    due_known_completed = completed[completed["Due Variance (days)"].notna()] if total else task_frame
    status_summary = _status_summary(status_frame)
    status_counts = (task_frame["Current Status"].fillna("Unavailable").value_counts().rename_axis("Status").reset_index(name="Tasks")
                     if total else pd.DataFrame(columns=["Status", "Tasks"]))
    weekly_flow = _weekly_flow(task_frame)
    due_status_summary = _due_status_summary(task_frame, cutoff)
    due_variance_summary = (task_frame[task_frame["Due Variance (days)"].notna()]
                            .groupby("Due Variance Category", dropna=False)
                            .agg(Tasks=("Task ID", "count"), Average_Variance_Days=("Due Variance (days)", "mean"),
                                 Median_Variance_Days=("Due Variance (days)", "median"))
                            .reset_index()
                            .rename(columns={"Average_Variance_Days": "Average Variance (days)", "Median_Variance_Days": "Median Variance (days)"})
                            .sort_values("Tasks", ascending=False).reset_index(drop=True)
                            if total else pd.DataFrame(columns=["Due Variance Category", "Tasks", "Average Variance (days)", "Median Variance (days)"]))
    assignee_summary = _assignee_summary(task_frame)
    late_completed = completed[completed["Due Variance (days)"] > 0].copy() if total else task_frame
    overdue_open = open_tasks[open_tasks["Due Variance (days)"] > 0].copy() if total else task_frame
    missing_due = task_frame[task_frame["Due Date"].isna()].copy() if total else task_frame
    status_duration_available = int(task_frame["Total Time in Status (min)"].notna().sum()) if total else 0
    future_start_tasks = (
        int(task_frame["Timing Data Status"].eq("Start date after evaluation cutoff").sum())
        if total else 0
    )
    timing_conflicts = (
        int(task_frame["Timing Data Status"].isin({
            "Start date after completion date", "Start date before creation date",
        }).sum())
        if total else 0
    )

    overall_values = [
        ("Total tasks", total),
        ("Completed tasks", int(task_frame["Completed?"].sum()) if total else 0),
        ("Open tasks", int(task_frame["Open?"].sum()) if total else 0),
        ("WIP tasks", int(task_frame["WIP?"].sum()) if total else 0),
        ("Cancelled tasks", int(task_frame["Cancelled?"].sum()) if total else 0),
        ("Completion rate (%)", float(task_frame["Completed?"].mean() * 100.0) if total else None),
        ("On-time completion rate (%)", _safe_rate(completed["On Time?"]) if total else None),
        ("Open overdue tasks", int(len(overdue_open))),
        ("Completed late tasks", int(len(late_completed))),
        ("Average due variance (days)", _mean(due_known_completed, "Due Variance (days)")),
        ("Average execution hours", _mean(completed, "Execution Hours")),
        ("Average lead time hours", _mean(completed, "Lead Time Hours")),
        ("Average time to start hours", _mean(task_frame, "Time to Start Hours")),
        ("Tasks with Total time in Status", status_duration_available),
    ]
    overall = pd.DataFrame(overall_values, columns=["Metric", "Value"])
    time_status_state = (
        "Available" if status_duration_available
        else ("Unavailable for this scope/account." if getattr(prepared_data, "clickup_time_status_error", "") else "Not requested")
    )
    analysis_context = pd.DataFrame([
        ("Process Name", prepared_data.space_name),
        ("Evaluation Scope", getattr(prepared_data, "filter_summary", "All tasks in the selected ClickUp Space")),
        ("Dataset Type", "ClickUp API collection"),
        ("Evaluation Cutoff", prepared_data.cutoff),
        ("Source Timezone", prepared_data.source_timezone),
        ("Task Count", total),
        ("Activity History", "Not collected; browser collector is disabled."),
        ("Total time in Status", time_status_state),
    ], columns=["Field", "Value"])
    quality = pd.DataFrame([
        ("Task rows collected", total, "OK" if total else "Unavailable"),
        ("Tasks with status-duration data", status_duration_available, "OK" if status_duration_available else "Unavailable"),
        ("Tasks missing created date", int(task_frame["Created"].isna().sum()) if total else 0,
         "Review" if total and task_frame["Created"].isna().any() else "OK"),
        ("Tasks missing due date", int(task_frame["Due Date"].isna().sum()) if total else 0,
         "Review" if total and task_frame["Due Date"].isna().any() else "OK"),
        ("Completed tasks missing completion date", int((task_frame["Completed?"] & task_frame["Completed"].isna()).sum()) if total else 0,
         "Review" if total and (task_frame["Completed?"] & task_frame["Completed"].isna()).any() else "OK"),
        ("Tasks scheduled after the evaluation cutoff", future_start_tasks,
         "Info" if future_start_tasks else "OK"),
        ("Tasks with timing chronology conflicts", timing_conflicts,
         "Review" if timing_conflicts else "OK"),
    ], columns=["Check", "Value", "Status"])
    metric_definitions = pd.DataFrame([
        ("Due Variance (days)", "Completed task: Completed date minus Due date. Open task: Evaluation cutoff minus Due date. Positive is late/overdue; negative is early/time remaining."),
        ("On-time completion rate", "Share of completed tasks with both a due date and a completion date where completion was on or before the due date."),
        ("Execution and time to start", "Calculated only when the recorded start date is on or before the completion date or evaluation cutoff. Date-only starts on the creation day use creation time as the effective start. Chronology conflicts remain unavailable."),
        ("WIP tasks", "Open tasks whose current status is not a common not-started status (Backlog, To Do, Planning, Ready, or Open)."),
        ("Status-duration data", "Read only from ClickUp Total time in Status when the ClickApp/API exposes it. Missing values stay unavailable."),
        ("Individual assignment", "Uses the assignee snapshot returned by ClickUp. It does not establish individual contribution or ownership at completion."),
    ], columns=["Metric", "Definition"])
    findings = pd.DataFrame([
        ("Completed late tasks", int(len(late_completed)), "Review deadlines, task estimates, and dependencies for these completed tasks."),
        ("Open overdue tasks", int(len(overdue_open)), "Prioritize current blockers and agree a recovery plan for open overdue work."),
        ("Tasks without due dates", int(len(missing_due)), "Add due dates where delivery timeliness is expected to be evaluated."),
        ("Tasks with timing chronology conflicts", timing_conflicts, "Review the task start and completion dates. Conflicting timing values are excluded from duration averages."),
        ("Tasks without status-duration data", total - status_duration_available, "Enable or verify Total time in Status only when the ClickUp plan and permissions support it."),
    ], columns=["Finding", "Tasks", "Recommended Follow-up"])
    return {
        "tasks": task_frame,
        "status_detail": status_frame,
        "status_summary": status_summary,
        "status_counts": status_counts,
        "weekly_flow": weekly_flow,
        "due_status_summary": due_status_summary,
        "due_variance_summary": due_variance_summary,
        "assignee_summary": assignee_summary,
        "late_completed_tasks": late_completed,
        "overdue_tasks": overdue_open,
        "missing_due_tasks": missing_due,
        "overall": overall,
        "analysis_context": analysis_context,
        "quality": quality,
        "metric_definitions": metric_definitions,
        "findings": findings,
        "cutoff": prepared_data.cutoff,
        "source_timezone": prepared_data.source_timezone,
        "space_name": prepared_data.space_name,
        "filter_summary": getattr(prepared_data, "filter_summary", "All tasks in the selected ClickUp Space"),
    }


def _write_sheet(workbook, name: str, frame: pd.DataFrame):
    sheet = workbook.create_sheet(name)
    for column, value in enumerate(frame.columns, 1):
        cell = sheet.cell(1, column, value)
        cell.fill = PatternFill("solid", fgColor="17324D")
        cell.font = Font(color="FFFFFF", bold=True)
    for row_number, row in enumerate(frame.itertuples(index=False, name=None), 2):
        for column, value in enumerate(row, 1):
            if isinstance(value, (dict, list)):
                value = json.dumps(value, ensure_ascii=False)
            if isinstance(value, pd.Timestamp):
                value = value.isoformat()
            sheet.cell(row_number, column, value)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for column_cells in sheet.columns:
        width = min(max(len(str(cell.value or "")) for cell in column_cells) + 2, 48)
        sheet.column_dimensions[column_cells[0].column_letter].width = width


def analysis_excel(result) -> bytes:
    """Create an English ClickUp analysis workbook with all dashboard tables."""
    workbook = Workbook()
    workbook.remove(workbook.active)
    sheets = [
        ("Executive_Summary", result["overall"]),
        ("Task_Evaluation", result["tasks"]),
        ("Assignee_Summary", result["assignee_summary"]),
        ("Due_Variance", result["due_variance_summary"]),
        ("Open_Due_Status", result["due_status_summary"]),
        ("Weekly_Flow", result["weekly_flow"]),
        ("Status_Summary", result["status_summary"]),
        ("Status_Detail", result["status_detail"]),
        ("Findings", result["findings"]),
        ("Data_Quality", result["quality"]),
        ("Metric_Definitions", result["metric_definitions"]),
        ("Analysis_Context", result["analysis_context"]),
    ]
    for name, frame in sheets:
        _write_sheet(workbook, name, frame)
    stream = BytesIO()
    workbook.save(stream)
    return stream.getvalue()
