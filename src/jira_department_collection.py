"""Department-first Jira collection for the Tech department.

The Tech department is Jira-owned. The collector intentionally reads every
accessible Jira project ("space") and every matching work item in the selected
period, then prepares the same source contract used by the existing Jira
analysis pipeline.
"""
from __future__ import annotations

import json
from uuid import uuid4
from datetime import datetime, time, timezone
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

from jira_export import PreparedData, build_workbook
from jira_gateway import CollectionError, JiraGateway
from resumable_jira import collect_jira_query


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
    # Candidate scope includes old work that is still open plus terminal work
    # changed during/after the period.  Exact active-life filtering is applied
    # after complete histories are collected; creation date alone is not a
    # valid department-performance scope.
    return (
        f'project = "{_jql_text(project_key)}" '
        f'AND created <= "{end_date}" '
        f'AND (statusCategory != Done OR updated >= "{start_date}") '
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


def _collect(
    settings,
    spaces,
    start_date,
    end_date,
    progress,
    fingerprint: str,
    seed_items_by_project: dict[str, list[dict]] | None = None,
    collection_id: str | None = None,
    fresh: bool = False,
) -> PreparedData:
    source_timezone = settings.source_timezone
    cutoff = _analysis_cutoff(end_date, source_timezone)
    complete_items = []
    histories = {}
    definitions = []
    seen_keys = set()
    queries = []
    collected_spaces = []

    for space in spaces:
        project_key = str(space.get("key") or "").strip()
        project_id = str(space.get("id") or "").strip()
        if not project_key or not project_id:
            continue
        seeds = (
            (seed_items_by_project or {}).get(project_key)
            if seed_items_by_project is not None else None
        )
        if seeds == []:
            continue
        query = _period_query(project_key, start_date, end_date)
        queries.append(query)
        name = str(space.get("name") or project_key)
        prepared_space = collect_jira_query(
            settings,
            space=space,
            query=query,
            fingerprint=f"{fingerprint}:{project_key}",
            collection_id=collection_id,
            fresh=fresh,
            cutoff=cutoff,
            seed_issues=seeds,
            progress=lambda fraction, message, label=name: progress(
                f"{label}: {message}"
            ),
        )
        collected_spaces.append(name)
        definitions = definitions or list(getattr(prepared_space, "definitions", []))
        for item in getattr(prepared_space, "issues", []):
            key = str(item.get("key") or "")
            if not key or key in seen_keys:
                raise CollectionError(
                    "Jira returned a missing or duplicate work item across spaces. "
                    "Please collect again."
                )
            seen_keys.add(key)
            complete_items.append(item)
        histories.update(getattr(prepared_space, "histories", {}))

    if not complete_items:
        raise CollectionError(
            "No Jira work items match the Tech department and selected period."
        )

    collected_at = datetime.now(timezone.utc).isoformat()
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
        preferred_start=settings.start_date_field,
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
    prepared.space_names = collected_spaces
    prepared.period_start = str(start_date)
    prepared.period_end = str(end_date)
    prepared.department_name = "Tech Development"
    prepared.department_id = "jira-tech-development"
    return prepared


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

    if "jira_department_collection_id" not in st.session_state:
        st.session_state["jira_department_collection_id"] = uuid4().hex
        st.session_state["jira_department_collection_fresh"] = True
    if st.button("Refresh Live Data", key="jira_department_refresh"):
        for key in (
            "jira_department_prepared_data",
            "jira_department_preview_cache",
            "department_analysis",
        ):
            st.session_state.pop(key, None)
        st.session_state["jira_department_collection_id"] = uuid4().hex
        st.session_state["jira_department_collection_fresh"] = True
        st.rerun()

    fingerprint = json.dumps({
        "department": "jira-tech-development",
        "from": str(start_date),
        "to": str(end_date),
        "spaces": [(str(s.get("id")), str(s.get("key"))) for s in spaces],
        "collection_id": st.session_state.get("jira_department_collection_id"),
    }, sort_keys=True)
    previous = st.session_state.get("jira_department_prepared_data")
    if previous is not None and previous.fingerprint != fingerprint:
        # Keep the preview and date controls visible. A changed period is
        # collected again only when the user runs the analysis.
        st.session_state.pop("department_analysis", None)

    preview_cache = st.session_state.get("jira_department_preview_cache") or {}
    cached_preview = preview_cache.get(fingerprint)
    if cached_preview is not None:
        preview_items_by_project = cached_preview
    else:
        preview_gateway = JiraGateway(settings)
        preview_progress = st.progress(
            0.0, text=f"Scanning Jira department Spaces: 0 / {len(spaces)}"
        )
        try:
            preview_items_by_project = {}
            active_spaces = [
                space for space in spaces if str(space.get("key") or "").strip()
            ]
            for index, space in enumerate(active_spaces):
                key = str(space.get("key") or "").strip()
                query = _period_query(key, start_date, end_date)

                def update_preview(message, position=index, label=key):
                    fraction = position / len(active_spaces) if active_spaces else 1.0
                    preview_progress.progress(
                        fraction,
                        text=f"{label}: {message}",
                    )

                preview_items_by_project[key] = preview_gateway.all_issues(
                    query, progress=update_preview
                )
                preview_progress.progress(
                    (index + 1) / len(active_spaces) if active_spaces else 1.0,
                    text=f"Scanned Jira department Spaces: {index + 1} / {len(active_spaces)}",
                )
            preview_cache = {fingerprint: preview_items_by_project}
            st.session_state["jira_department_preview_cache"] = preview_cache
            preview_progress.progress(
                1.0, text=f"Jira department collection complete: {sum(len(items) for items in preview_items_by_project.values())} tasks"
            )
        except CollectionError as exc:
            st.error(str(exc))
            return None, False
        finally:
            preview_gateway.close()

    preview_items = []
    preview_spaces = []
    for space in spaces:
        key = str(space.get("key") or "").strip()
        items = preview_items_by_project.get(key, [])
        preview_items.extend(items)
        if items:
            preview_spaces.append(space)

    if st.session_state.get("jira_department_collection_collected_at"):
        st.caption(f"Data collected at: {st.session_state['jira_department_collection_collected_at']}")
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
            progress = st.progress(0.0, text="Collecting Tech work items from Jira...")
            prepared = _collect(
                settings,
                spaces,
                start_date,
                end_date,
                lambda message: progress.progress(0.0, text=str(message)),
                fingerprint=fingerprint,
                collection_id=st.session_state.get("jira_department_collection_id"),
                fresh=st.session_state.get("jira_department_collection_fresh", False),
                seed_items_by_project=preview_items_by_project,
            )
            st.session_state["jira_department_prepared_data"] = prepared
            st.session_state["jira_department_collection_fresh"] = False
            st.session_state["jira_department_collection_collected_at"] = datetime.now(timezone.utc).isoformat()
            progress.progress(1.0, text="Tech Jira collection completed.")
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
                progress = st.progress(0.0, text="Updating the Tech Jira analysis source...")
                prepared = _collect(
                    settings,
                    spaces,
                    start_date,
                    end_date,
                    lambda message: progress.progress(0.0, text=str(message)),
                    fingerprint=fingerprint,
                    seed_items_by_project=preview_items_by_project,
                )
                st.session_state["jira_department_prepared_data"] = prepared
                progress.progress(1.0, text="Tech Jira source updated.")
            except Exception as exc:
                st.session_state["jira_department_collection_fresh"] = False
                st.error(f"Tech Jira source could not be updated: {exc}")
                return None, False
        st.session_state["jira_department_run_requested"] = True
        st.rerun()

    run_requested = st.session_state.pop(
        "jira_department_run_requested", False
    )
    return prepared, run_requested
