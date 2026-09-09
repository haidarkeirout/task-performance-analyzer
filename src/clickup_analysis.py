"""ClickUp-only performance analysis built from task and time-in-status data."""
from __future__ import annotations

import json
from io import BytesIO

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill


TERMINAL_STATUSES = {"complete", "completed", "done", "closed"}
CANCELLED_STATUSES = {"cancelled", "canceled"}


def _timestamp(value):
    if value in (None, "", 0, "0"):
        return pd.NaT
    try:
        numeric = float(value)
        if numeric > 100_000_000_000:
            return pd.to_datetime(numeric, unit="ms", utc=True, errors="coerce")
    except (TypeError, ValueError):
        pass
    return pd.to_datetime(value, utc=True, errors="coerce")


def _status_minutes(payload):
    values = {}
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


def _hours(value):
    return None if value is None or pd.isna(value) else float(value) / 60.0


def analyze_clickup(prepared_data):
    """Return isolated ClickUp analysis tables and an Excel-ready result."""
    payload = json.loads(prepared_data.history_json.decode("utf-8"))
    tasks = payload.get("tasks") if isinstance(payload, dict) else []
    time_payloads = payload.get("time_in_status", {}) if isinstance(payload, dict) else {}
    cutoff = _timestamp(prepared_data.cutoff)
    rows = []
    status_rows = []
    for task in tasks if isinstance(tasks, list) else []:
        task_id = str(task.get("id", ""))
        status_obj = task.get("status") or {}
        status = str(status_obj.get("status") or "Unavailable") if isinstance(status_obj, dict) else str(status_obj)
        status_key = status.casefold()
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
        created = _timestamp(task.get("date_created"))
        due = _timestamp(task.get("due_date"))
        completed = _timestamp(task.get("date_closed") or task.get("date_done"))
        completed_flag = status_key in TERMINAL_STATUSES or pd.notna(completed)
        cancelled = status_key in CANCELLED_STATUSES
        open_flag = not completed_flag and not cancelled
        elapsed_hours = None
        if pd.notna(created):
            end = completed if pd.notna(completed) else cutoff
            if pd.notna(end):
                elapsed_hours = (end - created).total_seconds() / 3600.0
        overdue_days = None
        if open_flag and pd.notna(due) and pd.notna(cutoff) and due < cutoff:
            overdue_days = (cutoff - due).total_seconds() / 86400.0
        on_time = None
        if completed_flag and pd.notna(completed) and pd.notna(due):
            on_time = bool(completed <= due)
        assignees = task.get("assignees") or []
        assignee = ", ".join(str(a.get("username") or a.get("email") or a.get("id"))
                            for a in assignees if isinstance(a, dict)) or "Unassigned"
        row = {
            "Task ID": task_id,
            "Task Name": task.get("name"),
            "Assignee": assignee,
            "Priority": (task.get("priority") or {}).get("priority") if isinstance(task.get("priority"), dict) else task.get("priority"),
            "Current Status": status,
            "Created": created,
            "Due Date": due,
            "Completed": completed,
            "Completed?": completed_flag,
            "Cancelled?": cancelled,
            "Open?": open_flag,
            "Elapsed Hours": elapsed_hours,
            "On Time?": on_time,
            "Overdue Days": overdue_days,
            "Current Status Time (min)": current_minutes,
            "Current Status Since": current_since,
            "Total Time in Status (min)": sum(minutes_by_status.values()) if minutes_by_status else None,
            "Time in Status": json.dumps(minutes_by_status, ensure_ascii=False) if minutes_by_status else None,
        }
        rows.append(row)
        for status_name, minutes in minutes_by_status.items():
            status_rows.append({"Status": status_name, "Task ID": task_id, "Minutes": minutes, "Hours": minutes / 60.0})

    task_frame = pd.DataFrame(rows)
    if task_frame.empty:
        task_frame = pd.DataFrame(columns=[
            "Task ID", "Task Name", "Assignee", "Priority", "Current Status", "Created", "Due Date",
            "Completed", "Completed?", "Cancelled?", "Open?", "Elapsed Hours", "On Time?", "Overdue Days",
            "Current Status Time (min)", "Current Status Since", "Total Time in Status (min)", "Time in Status",
        ])
    status_frame = pd.DataFrame(status_rows)
    if status_frame.empty:
        status_summary = pd.DataFrame(columns=["Status", "Tasks", "Average Hours", "Median Hours", "Total Hours"])
    else:
        status_summary = (status_frame.groupby("Status", dropna=False)
                          .agg(Tasks=("Task ID", "nunique"), Average_Hours=("Hours", "mean"),
                               Median_Hours=("Hours", "median"), Total_Hours=("Hours", "sum"))
                          .reset_index()
                          .rename(columns={"Average_Hours": "Average Hours", "Median_Hours": "Median Hours", "Total_Hours": "Total Hours"})
                          .sort_values("Total Hours", ascending=False))

    total = len(task_frame)
    completed_count = int(task_frame["Completed?"].sum()) if total else 0
    open_count = int(task_frame["Open?"].sum()) if total else 0
    overdue_count = int(task_frame["Overdue Days"].notna().sum()) if total else 0
    durations = pd.to_numeric(task_frame["Elapsed Hours"], errors="coerce") if total else pd.Series(dtype=float)
    overall = pd.DataFrame([{
        "Metric": "Total tasks", "Value": total,
    }, {
        "Metric": "Completed tasks", "Value": completed_count,
    }, {
        "Metric": "Open tasks", "Value": open_count,
    }, {
        "Metric": "Cancelled tasks", "Value": int(task_frame["Cancelled?"].sum()) if total else 0,
    }, {
        "Metric": "Completion rate (%)", "Value": (completed_count / total * 100.0) if total else None,
    }, {
        "Metric": "On-time completion rate (%)", "Value": (task_frame["On Time?"].dropna().mean() * 100.0) if total and task_frame["On Time?"].notna().any() else None,
    }, {
        "Metric": "Open overdue tasks", "Value": overdue_count,
    }, {
        "Metric": "Average elapsed hours", "Value": durations.mean() if not durations.dropna().empty else None,
    }, {
        "Metric": "Tasks with Total time in Status", "Value": int(task_frame["Total Time in Status (min)"].notna().sum()) if total else 0,
    }])
    quality = pd.DataFrame([{
        "Check": "Task rows collected", "Value": total, "Status": "OK" if total else "Unavailable",
    }, {
        "Check": "Tasks with status-duration data", "Value": int(task_frame["Total Time in Status (min)"].notna().sum()) if total else 0,
        "Status": "OK" if total and task_frame["Total Time in Status (min)"].notna().any() else "Unavailable",
    }, {
        "Check": "Tasks missing created date", "Value": int(task_frame["Created"].isna().sum()) if total else 0,
        "Status": "Review" if total and task_frame["Created"].isna().any() else "OK",
    }])
    return {"tasks": task_frame, "status_detail": status_frame, "status_summary": status_summary,
            "overall": overall, "quality": quality, "cutoff": prepared_data.cutoff,
            "source_timezone": prepared_data.source_timezone, "space_name": prepared_data.space_name}


def _write_sheet(workbook, name, frame):
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


def analysis_excel(result) -> bytes:
    workbook = Workbook()
    workbook.remove(workbook.active)
    _write_sheet(workbook, "Task_Evaluation", result["tasks"])
    _write_sheet(workbook, "Status_Summary", result["status_summary"])
    _write_sheet(workbook, "Status_Detail", result["status_detail"])
    _write_sheet(workbook, "Overall_Summary", result["overall"])
    _write_sheet(workbook, "Data_Quality", result["quality"])
    stream = BytesIO()
    workbook.save(stream)
    return stream.getvalue()
