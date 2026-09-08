from __future__ import annotations

import json
from datetime import datetime, timezone
from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

from jira_export import PreparedData
from clickup_gateway import ClickUpActivityUnavailable


def _sheet(wb, name, headers, rows):
    ws = wb.create_sheet(name)
    for col, header in enumerate(headers, 1):
        cell = ws.cell(1, col, header)
        cell.fill = PatternFill("solid", fgColor="17324D")
        cell.font = Font(color="FFFFFF", bold=True)
    for r, row in enumerate(rows, 2):
        for c, value in enumerate(row, 1):
            ws.cell(r, c, json.dumps(value, ensure_ascii=False) if isinstance(value, (dict, list)) else value)
    ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions


def collect_data(gateway, tasks, space_name, fingerprint, source_timezone="Asia/Damascus", progress=None, space_id=None):
    cutoff = datetime.now(timezone.utc).isoformat(); rows=[]; activity_rows=[]; raw=[]
    activity_unavailable = False
    activity_notes = []
    for index, task in enumerate(tasks, 1):
        task_id = str(task.get("id", ""))
        if progress: progress(f"Collecting Activity: {index} of {len(tasks)} tasks...")
        try:
            activity = gateway.activity(task_id)
        except ClickUpActivityUnavailable as exc:
            # The task/list API is still valid; only the optional web Activity
            # History endpoint is unavailable for this account.
            activity = []
            activity_unavailable = True
            activity_notes.append([task_id, "Activity History unavailable", str(exc)])
        status=task.get("status") or {}; priority=task.get("priority") or {}; assignees=task.get("assignees") or []
        history_ids = [str(event.get("id") or event.get("history_id") or event.get("event_id") or event.get("activity_id"))
                       for event in activity
                       if event.get("id") or event.get("history_id") or event.get("event_id") or event.get("activity_id")]
        rows.append([space_id or task.get("space_id"), task_id, task.get("name"), ", ".join(str(a.get("username") or a.get("email") or a.get("id")) for a in assignees) or "Unassigned", priority.get("priority") if isinstance(priority,dict) else priority, status.get("status") if isinstance(status,dict) else status, task.get("date_created"), task.get("due_date"), task.get("date_closed"), task.get("time_estimate"), task.get("time_spent"), len(activity), ", ".join(history_ids)])
        for event in activity:
            event_id = event.get("id") or event.get("history_id") or event.get("event_id") or event.get("activity_id")
            activity_rows.append([space_id or task.get("space_id"), task_id, event_id, event.get("date") or event.get("timestamp"), event.get("user"), event.get("type"), event.get("field"), event.get("from"), event.get("to"), event.get("comment") or event.get("description")])
        raw.append([task_id,"task",json.dumps({**task, "selected_space_id": space_id},ensure_ascii=False)])
        raw.extend([[task_id,"activity",json.dumps(event,ensure_ascii=False)] for event in activity])
    wb=Workbook(); wb.remove(wb.active)
    _sheet(wb,"ClickUp_Data",["Space ID","Task ID","Task Name","Assignee","Priority","Current Status","Created","Due Date","Completed","Time Estimate (ms)","Time Spent (ms)","Activity Events","History IDs"],rows)
    _sheet(wb,"Activity",["Space ID","Task ID","History ID","Timestamp","User","Event Type","Field","From","To","Comment"],activity_rows)
    _sheet(wb,"Process_Context",["Field","Value"],[["Process Name",space_name],["Space ID",space_id or ""],["Evaluation Scope","Selected ClickUp Space"],["Dataset Type","ClickUp API collection"],["Evaluation Cutoff Date",cutoff],["Source Timezone",source_timezone],["Task Count",len(tasks)],["Activity API","Available" if not activity_unavailable else "Unavailable for this account"],["History IDs","Collected from Activity events when the endpoint is available." if not activity_unavailable else "Not returned by the public ClickUp API for this account."],["Activity Errors",len(activity_notes)]])
    _sheet(wb,"Collection_Notes",["Task ID","Issue","Detail"],activity_notes or [["","None","No ClickUp collection warnings."]])
    _sheet(wb,"Raw_JSON",["Task ID","Section","JSON"],raw)
    stream=BytesIO(); wb.save(stream)
    filename=f"ClickUp_{space_name.replace(' ','_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    prepared = PreparedData(stream.getvalue(),json.dumps({"tasks":tasks},ensure_ascii=False).encode(),cutoff,datetime.now(timezone.utc).isoformat(),space_name,fingerprint,len(tasks),filename,space_name,source_timezone)
    # PreparedData is shared with Jira; keep this ClickUp-only quality flag
    # dynamic so no Jira schema or export code needs to change.
    prepared.clickup_activity_available = not activity_unavailable
    prepared.clickup_activity_notes = activity_notes
    return prepared

