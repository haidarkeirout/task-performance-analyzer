"""English-only sign-in, space selection, filters, and collection screens."""
from __future__ import annotations

import pandas as pd
import streamlit as st

from automation_auth import SetupError, credentials_match, read_settings
from collection_job import CollectionJob
from jira_filters import (BASIC_IDS, FilterError, build_query, date_clause, field_catalog,
                          field_clause, plain_label, query_fingerprint, unquote_value)
from jira_gateway import CollectionError, JiraGateway


RESULT_KEYS = ("task_metrics", "process_data", "validation_log", "cutoff_text")


def invalidate_selection():
    job = st.session_state.pop("collection_job", None)
    if job is not None:
        job.cancel()
    for key in ("prepared_data", "prepared_fingerprint", "preview", "preview_query",
                "collection_error", "collection_notice", *RESULT_KEYS):
        st.session_state.pop(key, None)


def _login(settings):
    username = st.session_state.get("login_username", "")
    password = st.session_state.get("login_password", "")
    if credentials_match(username, password, settings):
        st.session_state.clear()
        st.session_state["auth_revision"] = settings.revision
    else:
        st.session_state["login_password"] = ""
        st.session_state["login_error"] = "Incorrect username or password. Please try again."


def _logout():
    invalidate_selection()
    st.session_state.clear()


def require_sign_in():
    try:
        try:
            values = st.secrets.to_dict()
        except FileNotFoundError:
            values = {}
        settings = read_settings(values)
    except (SetupError, ValueError):
        st.title("Jira Performance")
        st.info("The system is being configured. Please contact the administrator.")
        st.stop()
    if st.session_state.get("auth_revision") != settings.revision:
        if st.session_state.get("auth_revision"):
            invalidate_selection()
            st.session_state.clear()
        _, center, _ = st.columns([1, 1.3, 1])
        with center:
            st.title("Jira Performance")
            st.caption("Sign in to select a space and analyze its work items.")
            with st.form("sign_in"):
                st.text_input("Username", key="login_username")
                st.text_input("Password", type="password", key="login_password")
                st.form_submit_button("Sign In", type="primary", width="stretch",
                                      on_click=_login, args=(settings,))
            if st.session_state.get("login_error"):
                st.error(st.session_state["login_error"])
        st.stop()
    return settings


def _reset_connection():
    invalidate_selection()
    for key in ("jira_spaces", "jira_fields", "jira_identity", "jira_reference",
                "catalog_space", "suggestion_cache"):
        st.session_state.pop(key, None)


OPERATOR_LABELS = {"=": "Is", "!=": "Is not", "in": "Is any of", "not in": "Is none of",
                   "~": "Contains", "!~": "Does not contain", ">": "Greater than", "<": "Less than",
                   ">=": "At least", "<=": "At most", "is": "Is empty", "is not": "Is not empty",
                   "was": "Was", "was not": "Was not", "was in": "Was any of", "was not in": "Was none of",
                   "changed": "Changed"}


def _suggestions(gateway, field, search):
    cache = st.session_state.setdefault("suggestion_cache", {})
    key = (st.session_state.get("catalog_space"), field.id, search)
    if key not in cache:
        results = gateway.suggestions(field.jql, search)
        cache[key] = {unquote_value(item["value"]): plain_label(item.get("displayName", item["value"]))
                      for item in results if item.get("value") is not None}
    return cache[key]


def _generic_filter(gateway, field):
    prefix = "filter_" + field.id + "_"
    # Favor multi-select comparison without changing the server-advertised operator set.
    operators = list(field.operators)
    if "in" in operators:
        operators.insert(0, operators.pop(operators.index("in")))
    op = st.selectbox("Comparison", operators, format_func=lambda v: OPERATOR_LABELS.get(v, v.upper()),
                      key=prefix + "operator", on_change=invalidate_selection)
    if op in {"is", "is not", "changed"}:
        return field_clause(field, op)
    options = {"Values": "Values", "JQL expression": "JQL expression"}
    mode = st.radio("Value type", list(options), horizontal=True, key=prefix + "mode",
                    on_change=invalidate_selection)
    if mode == "JQL expression":
        value = st.text_input("Expression", key=prefix + "expression", on_change=invalidate_selection,
                              help="For example: currentUser() or membersOf(\"team-name\").")
        return field_clause(field, op, [value], value_mode=mode)
    if field.kind == "number":
        text = st.text_input("Value", key=prefix + "number", on_change=invalidate_selection,
                             help="Enter a number, or comma-separated numbers for an 'any of' comparison.")
        return field_clause(field, op, [x.strip() for x in text.split(",")])
    if op in {"~", "!~"}:
        value = st.text_input("Search text", key=prefix + "text", on_change=invalidate_selection)
        return field_clause(field, op, [value])
    suggestions = {}
    if field.auto or field.kind == "user":
        search = st.text_input("Search values", key=prefix + "lookup",
                               help="Type a name to request matching Jira values.")
        try:
            suggestions = _suggestions(gateway, field, search)
        except CollectionError:
            st.caption("Suggestions are unavailable. You can enter an exact Jira value below.")
    labels = st.session_state.setdefault(prefix + "labels", {})
    labels.update(suggestions)
    selected = st.session_state.get(prefix + "values", [])
    values = st.multiselect("Values", list(dict.fromkeys([*suggestions, *selected])),
                            format_func=lambda v: labels.get(v, v), accept_new_options=True,
                            key=prefix + "values", on_change=invalidate_selection,
                            help="Search the list or type an exact value and press Enter. Leave blank for any value.")
    return field_clause(field, op, values)


def _date_filter(gateway, field):
    prefix = "filter_" + field.id + "_"
    modes = ["Any time", "On", "Before", "After", "Between", "Within the last",
             "This week", "This month", "This year"]
    if "is" in field.operators:
        modes.append("Is empty")
    if "is not" in field.operators:
        modes.append("Is not empty")
    modes.append("Advanced comparison")
    mode = st.selectbox("Date range", modes, key=prefix + "date_mode", on_change=invalidate_selection)
    if mode == "Advanced comparison":
        return _generic_filter(gateway, field)
    start = end = None
    days = 7
    if mode in {"On", "Before", "After", "Between"}:
        start = st.date_input("From date" if mode == "Between" else "Date", value=None,
                              key=prefix + "from", on_change=invalidate_selection)
    if mode == "Between":
        end = st.date_input("To date", value=None, key=prefix + "to", on_change=invalidate_selection)
    if mode == "Within the last":
        days = st.number_input("Days", min_value=1, value=7, step=1,
                               key=prefix + "days", on_change=invalidate_selection)
    return date_clause(field, mode, start, end, days)


def _filter_popover(gateway, field, target):
    with target.popover(field.label, width="stretch"):
        st.caption(field.label)
        return _date_filter(gateway, field) if field.kind == "date" else _generic_filter(gateway, field)


def _work_items_frame(items, site_url):
    rows = []
    for item in items:
        fields = item.get("fields", {})
        rows.append({"Work item": site_url + "/browse/" + item["key"], "Summary": fields.get("summary"),
                     "Assignee": (fields.get("assignee") or {}).get("displayName", "Unassigned"),
                     "Type": (fields.get("issuetype") or {}).get("name"),
                     "Priority": (fields.get("priority") or {}).get("name"),
                     "Status": (fields.get("status") or {}).get("name"),
                     "Resolution": (fields.get("resolution") or {}).get("name", "Unresolved"),
                     "Created": fields.get("created"), "Due date": fields.get("duedate")})
    return pd.DataFrame(rows)


def _preview(gateway, query, site_url):
    if st.session_state.get("preview_query") != query:
        st.session_state.pop("preview", None)
        try:
            gateway.validate(query)
            page = gateway.search_page(query)
            st.session_state["preview"] = page
            st.session_state["preview_query"] = query
        except CollectionError as exc:
            st.error(str(exc))
            return False
    preview = st.session_state.get("preview")
    if preview is None:
        return False
    items = preview["issues"]
    if items:
        st.dataframe(_work_items_frame(items, site_url), hide_index=True, width="stretch",
                     column_config={"Work item": st.column_config.LinkColumn(
                         "Work item", display_text=r".*/browse/(.*)")})
    else:
        st.info("No work items match the selected filters.")
    cols = st.columns([3, 1, 1])
    cols[0].caption(
        f"{len(items)} work items shown. " +
        ("All matching work items are shown." if preview.get("isLast")
         else "Done collects all matching work items, including later pages.")
    )
    if cols[1].button("Refresh list", width="stretch"):
        st.session_state.pop("preview_query", None)
        st.rerun()
    if cols[2].button("Load more", disabled=preview.get("isLast", False), width="stretch"):
        try:
            page = gateway.search_page(query, preview.get("nextPageToken"))
            keys = {item["key"] for item in items}
            if any(item["key"] in keys for item in page["issues"]):
                raise CollectionError("The work item list changed. Refresh the list before continuing.")
            page["issues"] = items + page["issues"]
            st.session_state["preview"] = page
            st.rerun()
        except CollectionError as exc:
            st.error(str(exc))
    return True


def _render_collection_error(job, snapshot):
    detail = snapshot.get("error_detail") or {}
    completed = detail.get("completed", snapshot.get("completed", 0))
    total = detail.get("total", snapshot.get("total", 0))
    failed_issue = detail.get("failed_issue")
    stage = detail.get("stage") or snapshot.get("stage") or "Unknown stage"
    error_type = detail.get("type") or "CollectionError"
    message = detail.get("message") or snapshot.get("error") or "Unknown collection error."
    status = detail.get("http_status")
    reference = detail.get("reference") or snapshot.get("id")

    notice = (reference, failed_issue, stage, message, completed, total)
    if st.session_state.get("collection_notice") != notice:
        st.session_state["collection_notice"] = notice
        st.toast(
            f"Collection stopped at {stage}. {completed} of {total} work items are saved."
            if total else f"Collection stopped at {stage}.",
            icon="⚠️",
        )

    st.error("Collection stopped before all selected work items were collected.")
    st.markdown(f"**Problem:** {message}")
    details = [
        f"**Stage:** {stage}",
        f"**Progress saved:** {completed} / {total}" if total else f"**Progress saved:** {completed}",
        f"**Error type:** {error_type}",
    ]
    if failed_issue:
        details.insert(1, f"**Failed work item:** `{failed_issue}`")
    if status is not None:
        details.append(f"**Jira HTTP status:** `{status}`")
    details.append(f"**Reference:** `{reference}`")
    st.markdown("  \n".join(details))
    st.info("Completed work is preserved. Retry continues from the first unfinished work item; it does not restart from task 1.")

    retry_col, reset_col = st.columns(2)
    if retry_col.button("Retry Collection", type="primary", width="stretch",
                        key=f"retry_collection_{reference}"):
        job.start()
        st.session_state.pop("collection_notice", None)
        st.rerun()
    if reset_col.button("Start New Collection", width="stretch",
                        key=f"new_collection_{reference}"):
        invalidate_selection()
        st.rerun()


@st.fragment(run_every=1)
def _collection_monitor(job):
    """Advance one collection unit per fragment tick and keep the UI state explicit."""
    snapshot = job.snapshot()
    if snapshot["running"]:
        job.step()
        snapshot = job.snapshot()

    total = snapshot["total"]
    completed = snapshot["completed"]
    st.progress(
        completed / total if total else 0,
        text=f"{completed} of {total} tasks completed" if total else "Reading task list...",
    )
    st.caption(snapshot["message"])

    if snapshot["running"]:
        current = snapshot.get("current_issue")
        retry_attempt = snapshot.get("retry_attempt", 0)
        status_text = f"Collection is running. Elapsed: {snapshot['elapsed']} seconds."
        if current:
            status_text += f" Current work item: {current}."
        if retry_attempt:
            status_text += (
                f" Automatic retry {retry_attempt} of "
                f"{snapshot.get('max_auto_retries', 3)} is active."
            )
        status_text += f" Reference: {snapshot['id']}."
        st.info(status_text)
    elif snapshot["error"]:
        _render_collection_error(job, snapshot)
    elif snapshot["result"] is not None:
        if st.session_state.get("prepared_data") is not snapshot["result"]:
            st.session_state["prepared_data"] = snapshot["result"]
            st.session_state.pop("collection_notice", None)
            st.rerun()
        st.success("Collection complete. All selected work items were collected.")
        if st.button("Start New Collection", width="stretch",
                     key=f"new_collection_complete_{snapshot['id']}"):
            invalidate_selection()
            st.rerun()
    else:
        st.warning("Collection is paused. Start a new collection if you want to change the selected data.")


def render_collection(settings):
    gateway = JiraGateway(settings)
    try:
        return _render_collection(gateway, settings)
    finally:
        gateway.close()


def _render_collection(gateway, settings):
    header = st.columns([5, 1])
    header[0].caption("Select a space → Choose filters → Done → Run Analysis")
    header[1].button("Sign Out", on_click=_logout, width="stretch")
    st.subheader("Select a Space")
    try:
        if "jira_spaces" not in st.session_state:
            with st.spinner("Loading your Jira spaces..."):
                identity = gateway.identity()
                spaces = gateway.spaces()
                fields = gateway.fields()
            st.session_state.update(jira_identity=identity, jira_spaces=spaces, jira_fields=fields)
    except CollectionError as exc:
        st.error(str(exc))
        if st.button("Retry connection"):
            _reset_connection()
            st.rerun()
        return None, False
    spaces = {str(s["id"]): s for s in st.session_state["jira_spaces"]}
    if (st.session_state.get("selected_space") is not None and
            st.session_state["selected_space"] not in spaces):
        invalidate_selection()
        st.session_state["selected_space"] = None
    cols = st.columns([5, 1])
    selected = cols[0].selectbox(
        "Space", [None, *spaces],
        format_func=lambda x: "Choose a space" if x is None else f"{spaces[x]['name']} ({spaces[x]['key']})",
        key="selected_space", on_change=invalidate_selection,
    )
    if cols[1].button("Refresh spaces", width="stretch"):
        _reset_connection()
        st.rerun()
    if not spaces:
        st.info("No Jira spaces are available to the connected account.")
    if selected is None:
        return None, False
    space = spaces[selected]
    if st.session_state.get("catalog_space") != selected:
        invalidate_selection()
        for key in list(st.session_state):
            if key.startswith("filter_") or key in {"more_filters", "query_text", "search_text", "query_mode"}:
                del st.session_state[key]
        try:
            reference = gateway.filter_reference(selected)
            st.session_state["jira_reference"] = reference
            st.session_state["catalog_space"] = selected
        except CollectionError as exc:
            st.error(str(exc))
            return None, False
    catalog = field_catalog(st.session_state["jira_reference"], st.session_state["jira_fields"])
    st.subheader("All work items")
    st.caption(
        f"Selected space: {space['name']}. Date filters use the Jira account time zone: " +
        str(st.session_state["jira_identity"].get("timeZone", "Jira account default")) + "."
    )
    mode = st.radio("Search mode", ["Basic", "JQL"], horizontal=True,
                    key="query_mode", on_change=invalidate_selection)
    clauses, errors, query = [], [], None
    try:
        if mode == "Basic":
            text = st.text_input("Search work", key="search_text", on_change=invalidate_selection)
            cols = st.columns(5)
            for index, fid in enumerate(BASIC_IDS):
                try:
                    clauses.append(_filter_popover(gateway, catalog[fid], cols[index]))
                except FilterError as exc:
                    errors.append(str(exc))
            extras = sorted(
                [fid for fid in catalog if fid not in {*BASIC_IDS, "project", "text"}],
                key=lambda fid: catalog[fid].label.casefold(),
            )
            with cols[4].popover("More filters", width="stretch"):
                chosen = st.multiselect("Add filters", extras,
                                        format_func=lambda fid: catalog[fid].label,
                                        key="more_filters", on_change=invalidate_selection)
                st.caption("Available searchable fields are loaded from Jira for this space, including custom fields.")
            if chosen:
                for start in range(0, len(chosen), 3):
                    extra_cols = st.columns(3)
                    for index, fid in enumerate(chosen[start:start + 3]):
                        try:
                            clauses.append(_filter_popover(gateway, catalog[fid], extra_cols[index]))
                        except FilterError as exc:
                            errors.append(str(exc))
            query = build_query(space["key"], clauses, text=text)
        else:
            advanced = st.text_area(
                "JQL filter condition", key="query_text", on_change=invalidate_selection,
                placeholder='status = "In Progress" AND priority = High',
                help="The selected space is applied automatically. Enter a condition without ORDER BY.",
            )
            query = build_query(space["key"], advanced=advanced)
    except FilterError as exc:
        errors.append(str(exc))
    if errors:
        invalidate_selection()
        for error in errors:
            st.info(error)
        query = None
    fingerprint = query_fingerprint(selected, query, settings.revision) if query else None
    previous = st.session_state.get("prepared_data")
    if previous is not None and previous.fingerprint != fingerprint:
        invalidate_selection()
    can_collect = False
    if query:
        with st.expander("View query"):
            st.code(query, language="sql")
        can_collect = _preview(gateway, query, settings.jira_url)
    st.divider()

    job = st.session_state.get("collection_job")
    if job is not None and job.fingerprint != fingerprint:
        invalidate_selection()
        job = None

    # Done is only the initial confirmation action. Once a collection job exists,
    # its own explicit Running / Error / Complete state owns the controls.
    if job is None:
        if st.button("Done", type="primary", disabled=not can_collect, width="stretch",
                     help="Confirm the selected filters and collect all matching Jira work items."):
            invalidate_selection()
            job = CollectionJob(
                settings, space, query, fingerprint,
                st.session_state["jira_fields"], JiraGateway,
            )
            st.session_state["collection_job"] = job
            job.start()
            st.rerun()
    else:
        _collection_monitor(job)

    prepared = st.session_state.get("prepared_data")
    if prepared:
        st.success("Your data has been collected and is ready for analysis.")
        st.caption(f"{prepared.count} work items · {prepared.space_name} · Collected at {prepared.collected_at}")
        st.download_button("Download Source Excel", prepared.xlsx, prepared.filename,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           on_click="ignore")
    run_clicked = st.button("Run Analysis", type="primary", disabled=prepared is None)
    if prepared is None:
        if job is None:
            st.caption("Choose your filters and click Done to prepare the data before running analysis.")
        elif job.snapshot()["running"]:
            st.caption("Data collection is still running. Run Analysis will unlock when the Excel source is complete.")
        elif job.snapshot()["error"]:
            st.caption("Resolve the collection error above. Your saved progress will be reused when you retry.")
    return prepared, run_clicked
