"""Department-first Jira collection for the Tech department.

The Tech department is Jira-owned. The collector intentionally reads every
accessible Jira project ("space") and every matching work item in the selected
period, then prepares the same source contract used by the existing Jira
analysis pipeline.
"""
from __future__ import annotations

import json
from datetime import datetime, time, timezone
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

from jira_export import PreparedData, build_workbook
from jira_gateway import CollectionError, JiraGateway


def _date_value(value):
    if isinstance(value, tuple):
        return value[0] if value else None
    return value


def _analysis_cutoff(end_date, source_timezone: str) -> str:
    """Return a cutoff that never extends beyond the collection start time."""
    period_timezone = ZoneInfo(source_timezone)
    period_end = datetime.combine(
        end_date, time.max, tzinfo=period_timezone
    ).astimezone(timezone.utc)
    collection_start = datetime.now(timezone.utc)
    return min(period_end, collection_start).isoformat()


def _jql_text(value: str) -> str:
    return str(value or "").replace("\\", "\\\\").replace('"', '\\"')


def _period_query(project_key: str, start_date, end_date) -> str:
    return (
        f'project = "{_jql_text(project_key)}" '
        f'AND created >= "{start_date}" AND created <= "{end_date}" '
        "ORDER BY created ASC"
    )


def _preview(items: list[dict], jira_url: str) -> pd.DataFrame:
    rows = []
    for item in items:
        fields = item.get("fields") or {}
        assignee = fields.get("assignee") or {}
        issue_type = fields.get("issuetype") or {}
        priority = fields.get("priority") or {}
        status = fields.get("status") or {}
        rows.append({
            "Work Item": f"{jira_url.rstrip('/')}/browse/{item.get('key', '')}",
            "Key": item.get("key"),
            "Summary": fields.get("summary"),
            "Space": (fields.get("project") or {}).get("name"),
            "Assignee": assignee.get("displayName", "Unassigned"),
            "Type": issue_type.get("name"),
            "Priority": priority.get("name"),
            "Status": status.get("name"),
            "Created": fields.get("created"),
            "Due Date": fields.get("duedate"),
        })
    return pd.DataFrame(rows)


def _collect(settings, spaces, start_date, end_date, progress, fingerprint: str) -> PreparedData:
    gateway = JiraGateway(settings)
    gateway.progress = progress
    try:
        source_timezone = gateway.settings.source_timezone
        cutoff = _analysis_cutoff(end_date, source_timezone)
        definitions = gateway.fields()
        collected = []
        histories = {}
        seen_keys = set()
        queries = []

        for space in spaces:
            project_key = str(space.get("key") or "").strip()
            project_id = str(space.get("id") or "").strip()
            if not project_key or not project_id:
                continue
            query = _period_query(project_key, start_date, end_date)
            queries.append(query)
            items = gateway.all_issues(
                query,
                progress=lambda message, name=space.get("name", project_key):
                    progress(f"{name}: {message}"),
            )
            for item in items:
                key = item.get("key")
                if not key or key in seen_keys:
                    raise CollectionError(
                        "Jira returned a missing or duplicate work item across spaces. "
                        "Please collect again."
                    )
                seen_keys.add(key)
                collected.append((space, item))

        if not collected:
            raise CollectionError(
                "No Jira work items match the Tech department and selected period."
            )

        complete_items = []
        for index, (space, item) in enumerate(collected, 1):
            key = item["key"]
            progress(
                f"Collecting Jira work items: {index} of {len(collected)} "
                f"({key})..."
            )
            snapshot, history = gateway.complete_issue(item)
            actual_project = str(
                (snapshot.get("fields") or {}).get("project", {}).get("id", "")
            )
            if actual_project != str(space["id"]):
                raise CollectionError(
                    "A Jira work item changed projects during collection. "
                    "Please start a new collection."
                )
            complete_items.append(snapshot)
            histories[key] = history

        collected_at = datetime.now(timezone.utc).isoformat()
        space_names = [
            str(space.get("name") or space.get("key"))
            for space in spaces
            if space.get("name") or space.get("key")
        ]
        combined_query = " OR ".join(f"({query})" for query in queries)
        workbook = build_workbook(
            complete_items,
            histories,
            definitions,
            cutoff=cutoff,
            collected_at=collected_at,
            query=combined_query,
            space_name="All Jira Tech Spaces",
            source_timezone=source_timezone,
            preferred_start=gateway.settings.start_date_field,
        )
        filename = (
            f"Jira_Tech_All_Spaces_"
            f"{datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        prepared = PreparedData(
            workbook,
            json.dumps(histories, ensure_ascii=False).encode("utf-8"),
            cutoff,
            collected_at,
            combined_query,
            fingerprint,
            len(complete_items),
            filename,
            "All Jira Tech Spaces",
            source_timezone,
        )
        prepared.space_names = space_names
        prepared.period_start = str(start_date)
        prepared.period_end = str(end_date)
        prepared.department_name = "Tech Development"
        prepared.department_id = "jira-tech-development"
        return prepared
    finally:
        gateway.close()


def render_jira_department_collection(settings):
    """Render Tech's all-space Jira collection behind the Department selector."""
    st.session_state["data_source"] = "Jira"
    st.subheader("Tech department data collection")
    st.caption(
        "Tech is connected to Jira. The system will scan every accessible Jira "
        "space and collect all matching work items for the selected period."
    )

    try:
        gateway = JiraGateway(settings)
        try:
            spaces = gateway.spaces()
        finally:
            gateway.close()
    except CollectionError as exc:
        st.error(str(exc))
        return None, False

    if not spaces:
        st.warning("No Jira spaces are available to the connected account.")
        return None, False

    left, right = st.columns(2)
    with left:
        start_date = _date_value(
            st.date_input("From date", value=None, key="jira_department_period_start")
        )
    with right:
        end_date = _date_value(
            st.date_input("To date", value=None, key="jira_department_period_end")
        )
    if start_date is None or end_date is None:
        st.info("Select both From and To dates to load the Tech department scope.")
        return None, False
    if start_date > end_date:
        st.error("The From date must be on or before the To date.")
        return None, False

    fingerprint = json.dumps({
        "department": "jira-tech-development",
        "from": str(start_date),
        "to": str(end_date),
        "spaces": [(str(s.get("id")), str(s.get("key"))) for s in spaces],
    }, sort_keys=True)
    previous = st.session_state.get("jira_department_prepared_data")
    if previous is not None and previous.fingerprint != fingerprint:
        # Keep the preview and date controls visible. A changed period is
        # collected again only when the user runs the analysis.
        st.session_state.pop("department_analysis", None)

    with st.spinner("Scanning every Jira space for Tech work items..."):
        preview_gateway = JiraGateway(settings)
        try:
            preview_items = []
            preview_spaces = []
            for space in spaces:
                key = str(space.get("key") or "").strip()
                if not key:
                    continue
                query = _period_query(key, start_date, end_date)
                items = preview_gateway.all_issues(query)
                preview_items.extend(items)
                preview_spaces.append(space)
        except CollectionError as exc:
            st.error(str(exc))
            return None, False
        finally:
            preview_gateway.close()

    st.caption(
        f"{len(preview_items)} Jira work items found across "
        f"{len(preview_spaces)} space(s)."
    )
    st.dataframe(
        _preview(preview_items, settings.jira_url),
        hide_index=True,
        use_container_width=True,
    )

    if st.button(
        "Done",
        type="primary",
        key="jira_department_done",
        disabled=not preview_items,
    ):
        try:
            with st.status(
                "Collecting Tech work items from Jira...",
                expanded=True,
            ) as progress:
                prepared = _collect(
                    settings,
                    spaces,
                    start_date,
                    end_date,
                    lambda message: progress.update(label=message),
                    fingerprint=fingerprint,
                )
                st.session_state["jira_department_prepared_data"] = prepared
                progress.update(
                    label="Tech Jira collection completed.",
                    state="complete",
                    expanded=False,
                )
        except Exception as exc:
            st.error(f"Tech Jira collection could not be completed: {exc}")

    prepared = st.session_state.get("jira_department_prepared_data")
    prepared_is_current = (
        prepared is not None
        and prepared.fingerprint == fingerprint
    )
    if prepared_is_current:
        st.success(
            f"Tech collection completed: {prepared.count} work items from "
            f"{len(getattr(prepared, 'space_names', []))} Jira space(s). "
            "The source is ready for analysis."
        )
        st.download_button(
            "Download Tech source Excel",
            prepared.xlsx,
            prepared.filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            on_click="ignore",
        )

    run_clicked = st.button(
        "Run Analysis",
        type="primary",
        disabled=prepared is None or not preview_items,
        key="jira_department_run",
    )
    if run_clicked:
        if not prepared_is_current:
            try:
                with st.spinner("Updating the Tech Jira analysis source..."):
                    prepared = _collect(
                        settings,
                        spaces,
                        start_date,
                        end_date,
                        lambda message: st.write(message),
                        fingerprint=fingerprint,
                    )
                    st.session_state["jira_department_prepared_data"] = prepared
            except Exception as exc:
                st.error(f"Tech Jira source could not be updated: {exc}")
                return None, False
        st.session_state["jira_department_run_requested"] = True
        st.rerun()

    run_requested = st.session_state.pop(
        "jira_department_run_requested", False
    )
    return prepared, run_requested
