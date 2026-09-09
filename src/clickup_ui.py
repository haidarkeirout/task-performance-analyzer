"""English-only sign-in, space selection, filters, and collection screens."""
from __future__ import annotations

import json

import pandas as pd
import streamlit as st

from automation_auth import SetupError, credentials_match, read_settings
from jira_export import collect_data
from jira_filters import (BASIC_IDS, FilterError, build_query, date_clause, field_catalog,
                          field_clause, plain_label, query_fingerprint, unquote_value)
from jira_gateway import CollectionError, JiraGateway
from clickup_gateway import ClickUpCollectionError, ClickUpGateway, ClickUpTimeStatusUnavailable
from clickup_export import (_current_status_info, _status_minutes,
                            collect_data as collect_clickup_data)
from clickup_filters import (MORE_FILTER_LABELS, assignees as clickup_assignees,
                             criteria_summary, filter_options as clickup_filter_options,
                             filter_tasks as filter_clickup_tasks, status_name as clickup_status_name,
                             timestamp as clickup_timestamp, widget_key)


RESULT_KEYS = ("task_metrics", "process_data", "validation_log", "cutoff_text")


def invalidate_selection():
    for key in ("prepared_data", "prepared_fingerprint", "preview", "preview_query", "collection_error",
                "clickup_prepared_data", "clickup_analysis", "clickup_run_analysis", *RESULT_KEYS):
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
    st.session_state.clear()


def require_sign_in():
    try:
        try:
            values = st.secrets.to_dict()
        except FileNotFoundError:
            values = {}
        settings = read_settings(values)
    except (SetupError, ValueError):
        st.title("Task Performance Intelligence")
        st.info("The system is being configured. Please contact the administrator.")
        st.stop()
    if st.session_state.get("auth_revision") != settings.revision:
        if st.session_state.get("auth_revision"):
            st.session_state.clear()
        _, center, _ = st.columns([1, 1.3, 1])
        with center:
            st.title("Task Performance Intelligence")
            st.caption("Sign in to select a space and analyze its work items.")
            with st.form("sign_in"):
                st.text_input("Username", key="login_username")
                st.text_input("Password", type="password", key="login_password")
                st.form_submit_button("Sign In", type="primary", use_container_width=True,
                                      on_click=_login, args=(settings,))
            if st.session_state.get("login_error"):
                st.error(st.session_state["login_error"])
        st.stop()
    return settings


def _reset_connection():
    invalidate_selection()
    for key in ("jira_spaces", "jira_fields", "jira_identity", "jira_reference", "catalog_space", "suggestion_cache"):
        st.session_state.pop(key, None)



def _clickup_gateway(settings, workspace_id=""):
    """Build the isolated ClickUp API gateway; no browser session is used."""
    return ClickUpGateway(
        settings.clickup_token,
        workspace_id=str(workspace_id or settings.clickup_workspace_id or ""),
    )

def _clickup_date_rule(label, key):
    """Render one ClickUp-style date condition and return a small rule object."""
    modes = ["Any time", "On", "Before", "After", "Between", "Is set", "Is not set"]
    mode = st.selectbox(label, modes, key=key + "_mode")
    rule = {"mode": mode}
    if mode in {"On", "Before", "After", "Between"}:
        rule["start"] = st.date_input(
            "From date" if mode == "Between" else "Date",
            key=key + "_start",
        )
    if mode == "Between":
        rule["end"] = st.date_input("To date", key=key + "_end")
    return rule


def _clickup_number_rule(label, key):
    """Render a duration/number condition using hours, as ClickUp does."""
    modes = ["Any value", "At least (hours)", "At most (hours)", "Between (hours)", "Has value", "No value"]
    mode = st.selectbox(label, modes, key=key + "_mode")
    rule = {"mode": mode}
    if mode in {"At least (hours)", "At most (hours)", "Between (hours)"}:
        minimum = st.number_input("Minimum hours", min_value=0.0, value=0.0, step=0.5,
                                  key=key + "_minimum")
        rule["minimum"] = minimum
        if mode == "At most (hours)":
            rule["maximum"] = st.number_input("Maximum hours", min_value=0.0, value=minimum, step=0.5,
                                                key=key + "_maximum")
        elif mode == "Between (hours)":
            rule["maximum"] = st.number_input("Maximum hours", min_value=minimum,
                                                value=max(minimum, 1.0), step=0.5,
                                                key=key + "_maximum")
    return rule


def _clickup_boolean_rule(label, key):
    return st.selectbox(label, ["Any", "Is", "Is not"], key=key + "_value")


def _clickup_time_status(settings, workspace_id, space_id, tasks):
    """Read status-duration data only after the user adds that specific filter."""
    task_ids = [str(task.get("id", "")) for task in tasks if task.get("id")]
    cache_key = f"{workspace_id}:{space_id}:{','.join(task_ids)}"
    if st.session_state.get("clickup_time_status_key") != cache_key:
        time_status_map = {}
        error = ""
        try:
            gateway = _clickup_gateway(settings, workspace_id)
            try:
                time_status_map = gateway.time_in_status(task_ids) or {}
            finally:
                gateway.close()
        except (ClickUpTimeStatusUnavailable, ClickUpCollectionError) as exc:
            error = str(exc)
        st.session_state["clickup_time_status_key"] = cache_key
        st.session_state["clickup_time_status_map"] = time_status_map
        st.session_state["clickup_time_status_error"] = error
    return (
        st.session_state.get("clickup_time_status_map", {}),
        st.session_state.get("clickup_time_status_error", ""),
    )


def _clickup_display_timestamp(value):
    value = clickup_timestamp(value)
    return "" if pd.isna(value) else value.strftime("%Y-%m-%d")


def _render_clickup_collection(settings):
    """Collect one filtered ClickUp Space without changing the Jira path.

    ClickUp filters operate against an API snapshot.  The paid Total time in
    Status endpoint is requested only when the corresponding More filter is
    selected; it is not a background collector and never touches Jira.
    """
    st.caption("Select a ClickUp Space, apply filters, review the matching tasks, and select Done.")
    try:
        gateway = _clickup_gateway(settings)
        try:
            workspaces = gateway.workspaces()
            workspace_id = settings.clickup_workspace_id or (str(workspaces[0]["id"]) if workspaces else "")
            if not workspace_id:
                st.error("No ClickUp Workspace was found for this connection.")
                return None, False
            spaces = gateway.spaces(workspace_id)
        finally:
            gateway.close()
    except ClickUpCollectionError as exc:
        st.error(str(exc))
        return None, False

    space_map = {str(item["id"]): item for item in spaces}
    selected = st.selectbox(
        "Space",
        [None, *space_map],
        format_func=lambda value: "Choose a space" if value is None else space_map[value].get("name", value),
        key="clickup_selected_space",
    )
    if selected is None:
        return None, False

    tasks_cache_key = f"{workspace_id}:{selected}"
    if st.session_state.get("clickup_tasks_cache_key") != tasks_cache_key:
        try:
            gateway = _clickup_gateway(settings, workspace_id)
            try:
                with st.spinner("Loading ClickUp tasks..."):
                    tasks = gateway.all_tasks_for_space(selected)
            finally:
                gateway.close()
        except ClickUpCollectionError as exc:
            st.error(str(exc))
            return None, False
        st.session_state["clickup_tasks_cache_key"] = tasks_cache_key
        st.session_state["clickup_tasks_cache"] = tasks
        # A different Space must not reuse a status-duration result.
        for key in ("clickup_time_status_key", "clickup_time_status_map", "clickup_time_status_error"):
            st.session_state.pop(key, None)
    tasks = st.session_state.get("clickup_tasks_cache", [])
    options = clickup_filter_options(tasks)

    st.subheader("ClickUp filters")
    quick_left, quick_middle, quick_right = st.columns(3)
    with quick_left:
        keyword = st.text_input("Search tasks", key=f"clickup_search_{selected}")
        selected_statuses = st.multiselect("Status", options["status"], key=f"clickup_status_{selected}")
    with quick_middle:
        selected_assignees = st.multiselect("Assignee", options["assignee"], key=f"clickup_assignee_{selected}")
        selected_priorities = st.multiselect("Priority", options["priority"], key=f"clickup_priority_{selected}")
    with quick_right:
        due_date = _clickup_date_rule("Due date", f"clickup_due_{selected}")
        match = st.selectbox("Match filters", ["All (AND)", "Any (OR)"], key=f"clickup_match_{selected}")

    criteria = {
        "keyword": keyword,
        "status": selected_statuses,
        "assignee": selected_assignees,
        "priority": selected_priorities,
        "due_date": due_date,
        "match": match,
        "extras": {},
    }
    time_status_map = {}
    time_status_error = ""

    with st.expander("More filters", expanded=False):
        more_filters = st.multiselect(
            "Add filter",
            MORE_FILTER_LABELS,
            key=f"clickup_more_filters_{selected}",
            help="These labels mirror ClickUp View and Dashboard filter families when the task API provides the field.",
        )
        if more_filters:
            st.caption("Additional conditions are combined with the selected filter-matching rule above.")
        date_labels = {"Start date", "Date created", "Date updated", "Date closed"}
        for label in [item for item in more_filters if item in date_labels]:
            criteria["extras"][label] = _clickup_date_rule(label, f"clickup_more_{widget_key(label)}_{selected}")
        if "Tags" in more_filters:
            criteria["extras"]["Tags"] = {
                "mode": st.selectbox("Tag condition", ["Has any of", "Has all of", "Has none of"],
                                       key=f"clickup_tag_mode_{selected}"),
                "values": st.multiselect("Tags", options["tag"], key=f"clickup_tags_{selected}"),
            }
        for label, values in (("Location/List", options["location"]), ("Task type", options["task_type"]),
                              ("Created by", options["created_by"])):
            if label in more_filters:
                criteria["extras"][label] = st.multiselect(label, values,
                                                            key=f"clickup_more_{widget_key(label)}_{selected}")
        for label in ("Time estimates", "Time tracked"):
            if label in more_filters:
                criteria["extras"][label] = _clickup_number_rule(label, f"clickup_more_{widget_key(label)}_{selected}")
        for label in ("Status is closed", "Recurring", "Milestone", "Dependencies", "Archived"):
            if label in more_filters:
                criteria["extras"][label] = _clickup_boolean_rule(label, f"clickup_more_{widget_key(label)}_{selected}")
        if "Custom Fields" in more_filters:
            selected_fields = st.multiselect("Custom field", list(options["custom_fields"]),
                                             key=f"clickup_custom_fields_{selected}")
            selected_values = {}
            for field_name in selected_fields:
                selected_values[field_name] = st.multiselect(
                    field_name,
                    options["custom_fields"][field_name],
                    key=f"clickup_custom_{widget_key(field_name)}_{selected}",
                )
            if selected_values:
                criteria["extras"]["Custom Fields"] = selected_values
        if "Total time in Status" in more_filters:
            time_status_map, time_status_error = _clickup_time_status(settings, workspace_id, selected, tasks)
            if time_status_error:
                st.warning(
                    "Total time in Status is not available for this ClickUp account. "
                    "The task filters and analysis remain available."
                )
            else:
                status_minutes = {
                    task_id: _status_minutes(payload)
                    for task_id, payload in time_status_map.items()
                }
                status_options = sorted({status for values in status_minutes.values() for status in values}, key=str.casefold)
                if not status_options:
                    st.info("ClickUp returned no Total time in Status values for this Space.")
                else:
                    selected_status = st.selectbox("Status duration for", status_options,
                                                   key=f"clickup_tis_status_{selected}")
                    rule = _clickup_number_rule("Total time in Status", f"clickup_tis_{selected}")
                    rule["status"] = selected_status
                    criteria["extras"]["Total time in Status"] = rule

    filtered_tasks = filter_clickup_tasks(tasks, criteria, time_status_map)
    filter_summary = criteria_summary(criteria)
    rows = []
    for task in filtered_tasks:
        task_id = str(task.get("id", ""))
        current_minutes, _ = _current_status_info(time_status_map.get(task_id))
        rows.append({
            "Task ID": task_id,
            "Task Name": task.get("name"),
            "Assignee": ", ".join(clickup_assignees(task)),
            "Priority": (task.get("priority") or {}).get("priority") if isinstance(task.get("priority"), dict) else task.get("priority"),
            "Status": clickup_status_name(task),
            "Due date": _clickup_display_timestamp(task.get("due_date")),
            "Current Status Time (min)": current_minutes,
        })
    selected_name = space_map[selected].get("name", selected)
    st.subheader("Matching tasks")
    st.caption(
        f"Space: {selected_name} · Space ID: {selected} · {len(rows)} of {len(tasks)} tasks shown · Scope: {filter_summary}"
    )
    st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
    if not filtered_tasks:
        st.warning("No tasks match the current filters. Change a filter before selecting Done.")

    st.divider()
    fingerprint = "clickup:" + selected + ":" + json.dumps(
        {"task_ids": [str(task.get("id", "")) for task in filtered_tasks], "criteria": criteria},
        sort_keys=True,
        default=str,
    )
    previous = st.session_state.get("clickup_prepared_data")
    if previous is not None and previous.fingerprint != fingerprint:
        st.session_state.pop("clickup_prepared_data", None)
        st.session_state.pop("clickup_analysis", None)
    if st.button("Done", type="primary", key="clickup_done", disabled=not filtered_tasks):
        st.session_state.pop("clickup_error", None)
        st.session_state.pop("clickup_prepared_data", None)
        try:
            with st.status("Preparing the selected ClickUp tasks for analysis...", expanded=True) as progress:
                prepared = collect_clickup_data(
                    None,
                    filtered_tasks,
                    selected_name,
                    fingerprint,
                    settings.source_timezone,
                    progress=lambda message: progress.update(label=message),
                    space_id=selected,
                    # Passing a map, even an empty one, prevents a hidden API collector call.
                    time_status_data={
                        str(task.get("id", "")): time_status_map[str(task.get("id", ""))]
                        for task in filtered_tasks
                        if str(task.get("id", "")) in time_status_map
                    },
                    time_status_error=time_status_error,
                    filter_summary=filter_summary,
                    filter_criteria=criteria,
                )
                st.session_state["clickup_prepared_data"] = prepared
                st.session_state["clickup_run_analysis"] = True
                progress.update(label="ClickUp task collection completed.", state="complete", expanded=False)
        except Exception as exc:
            st.session_state["clickup_error"] = str(exc)
    if st.session_state.get("clickup_error"):
        st.error(st.session_state["clickup_error"])
    prepared = st.session_state.get("clickup_prepared_data")
    if prepared:
        st.success("ClickUp task collection completed. Jira was not used or changed.")
        if "Total time in Status" in criteria["extras"] and not getattr(prepared, "clickup_time_status_available", False):
            st.info("Total time in Status was unavailable for this selected scope; unavailable values remain blank.")
        st.download_button(
            "Download ClickUp source Excel",
            prepared.xlsx,
            prepared.filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            on_click="ignore",
        )
    return prepared, bool(st.session_state.pop("clickup_run_analysis", False))


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
    modes = ["Any time", "On", "Before", "After", "Between", "Within the last", "This week", "This month", "This year"]
    if "is" in field.operators: modes.append("Is empty")
    if "is not" in field.operators: modes.append("Is not empty")
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
    with target.popover(field.label, use_container_width=True):
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
    if preview is None: return False
    items = preview["issues"]
    if items:
        st.dataframe(_work_items_frame(items, site_url), hide_index=True, use_container_width=True,
                      column_config={"Work item": st.column_config.LinkColumn("Work item", display_text=r".*/browse/(.*)")})
    else:
        st.info("No work items match the selected filters.")
    cols = st.columns([3, 1, 1])
    cols[0].caption(f"{len(items)} work items shown. " +
                    ("All matching work items are shown." if preview.get("isLast") else "Done collects all matching work items, including later pages."))
    if cols[1].button("Refresh list", use_container_width=True):
        st.session_state.pop("preview_query", None)
        st.rerun()
    if cols[2].button("Load more", disabled=preview.get("isLast", False), use_container_width=True):
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


def render_collection(settings):
    # Keep the source selector outside either connector branch. Streamlit reruns
    # the script after every widget interaction; rendering the selector only in
    # the Jira branch could make a ClickUp selection fall back to Jira.
    source = st.radio("Data source", ["Jira", "ClickUp"], horizontal=True,
                      key="data_source", on_change=invalidate_selection)
    if source == "ClickUp":
        return _render_clickup_collection(settings)
    gateway = JiraGateway(settings)
    try:
        return _render_collection(gateway, settings)
    finally:
        gateway.close()


def _render_collection(gateway, settings):
    header = st.columns([5, 1])
    header[0].caption("Select a space → Choose filters → Done → Run Analysis")
    header[1].button("Sign Out", on_click=_logout, use_container_width=True)
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
    if st.session_state.get("selected_space") is not None and st.session_state["selected_space"] not in spaces:
        invalidate_selection()
        st.session_state["selected_space"] = None
    cols = st.columns([5, 1])
    selected = cols[0].selectbox("Space", [None, *spaces],
                                 format_func=lambda x: "Choose a space" if x is None else f"{spaces[x]['name']} ({spaces[x]['key']})",
                                 key="selected_space", on_change=invalidate_selection)
    if cols[1].button("Refresh spaces", use_container_width=True):
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
    st.caption(f"Selected space: {space['name']}. Date filters use the Jira account time zone: " +
               str(st.session_state["jira_identity"].get("timeZone", "Jira account default")) + ".")
    mode = st.radio("Search mode", ["Basic", "JQL"], horizontal=True, key="query_mode", on_change=invalidate_selection)
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
            extras = sorted([fid for fid in catalog if fid not in {*BASIC_IDS, "project", "text"}], key=lambda fid: catalog[fid].label.casefold())
            with cols[4].popover("More filters", use_container_width=True):
                chosen = st.multiselect("Add filters", extras, format_func=lambda fid: catalog[fid].label,
                                        key="more_filters", on_change=invalidate_selection)
                st.caption("Available searchable fields are loaded from Jira for this space, including custom fields.")
            if chosen:
                for start in range(0, len(chosen), 3):
                    extra_cols = st.columns(3)
                    for index, fid in enumerate(chosen[start:start+3]):
                        try:
                            clauses.append(_filter_popover(gateway, catalog[fid], extra_cols[index]))
                        except FilterError as exc:
                            errors.append(str(exc))
            query = build_query(space["key"], clauses, text=text)
        else:
            advanced = st.text_area("JQL filter condition", key="query_text", on_change=invalidate_selection,
                                     placeholder='status = "In Progress" AND priority = High',
                                     help="The selected space is applied automatically. Enter a condition without ORDER BY.")
            query = build_query(space["key"], advanced=advanced)
    except FilterError as exc:
        errors.append(str(exc))
    if errors:
        invalidate_selection()
        for error in errors: st.info(error)
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
    if st.button("Done", type="primary", disabled=not can_collect, help="Collect all matching work items and prepare their Excel file."):
        invalidate_selection()
        try:
            with st.status("Collecting your data...", expanded=True) as status:
                prepared = collect_data(gateway, space, query, fingerprint, st.session_state["jira_fields"],
                                         progress=lambda message: status.update(label=message))
                st.session_state["prepared_data"] = prepared
                status.update(label="Data collection completed.", state="complete", expanded=False)
        except (CollectionError, ValueError) as exc:
            st.session_state["collection_error"] = str(exc)
        except Exception:
            st.session_state["collection_error"] = "Data collection could not be completed. Please try again or contact the administrator."
    if st.session_state.get("collection_error"):
        st.error(st.session_state["collection_error"])
    prepared = st.session_state.get("prepared_data")
    if prepared:
        st.success("Your data has been collected and is ready for analysis.")
        st.caption(f"{prepared.count} work items · {prepared.space_name} · Collected at {prepared.collected_at}")
        st.download_button("Download Source Excel", prepared.xlsx, prepared.filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", on_click="ignore")
    run_clicked = st.button("Run Analysis", type="primary", disabled=prepared is None)
    if prepared is None:
        st.caption("Choose your filters and click Done to prepare the data before running analysis.")
    return prepared, run_clicked
