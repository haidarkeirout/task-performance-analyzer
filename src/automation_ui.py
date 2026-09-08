"""English-only sign-in, space selection, filters, and collection screens."""
from __future__ import annotations

import pandas as pd
import streamlit as st

from automation_auth import SetupError, credentials_match, read_settings
from jira_export import collect_data
from jira_filters import (BASIC_IDS, FilterError, build_query, date_clause, field_catalog,
                          field_clause, plain_label, query_fingerprint, unquote_value)
from jira_gateway import CollectionError, JiraGateway
from collection_store import CollectionStore, persistence_owner_key
from collection_job import CollectionJob
from clickup_gateway import ClickUpCollectionError, ClickUpGateway
from clickup_export import collect_data as collect_clickup_data


RESULT_KEYS = ("task_metrics", "process_data", "validation_log", "cutoff_text")


def invalidate_selection():
    for key in ("prepared_data", "prepared_fingerprint", "preview", "preview_query", "collection_error", *RESULT_KEYS):
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
        st.title("Jira Performance")
        st.info("The system is being configured. Please contact the administrator.")
        st.stop()
    if st.session_state.get("auth_revision") != settings.revision:
        if st.session_state.get("auth_revision"):
            st.session_state.clear()
        _, center, _ = st.columns([1, 1.3, 1])
        with center:
            st.title("Jira Performance")
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

def _collection_store(settings):
    if "_collection_store" not in st.session_state:
        st.session_state["_collection_store"] = CollectionStore.configured()
    return st.session_state["_collection_store"]

def _restore_collection_prompt(settings):
    if "collection_job" in st.session_state:
        return
    store = _collection_store(settings)
    if store is None:
        return
    payload = store.latest_resumable(persistence_owner_key(settings))
    if not payload:
        return
    st.info("A previous Jira collection is available.")
    if st.button("Resume Previous Collection", key="resume_previous_collection"):
        st.session_state["collection_job"] = CollectionJob.from_persisted(
            settings, payload, JiraGateway, store=store,
            owner_key=persistence_owner_key(settings))
        st.rerun()

def _advance_collection_job(settings):
    job = st.session_state.get("collection_job")
    if job is None:
        return False
    if job.result is not None:
        st.session_state["prepared_data"] = job.result
        return False
    if job.error:
        st.error(job.error)
        if st.button("Retry Collection", key="retry_collection"):
            job.start()
            job.step()
            st.rerun()
        return True
    if job.running:
        job.step()
        snap = job.snapshot()
        st.info(f"{snap['message']} ({snap['completed']} of {snap['total']})")
        if job.result is not None:
            st.session_state["prepared_data"] = job.result
        return job.result is None
    if st.button("Resume Collection", key="resume_collection"):
        job.start()
        job.step()
        st.rerun()
    return True


def _render_clickup_collection(settings):
    st.caption("Select a ClickUp Space → review its tasks → Done")
    try:
        gateway = ClickUpGateway(settings.clickup_token)
        try:
            workspaces = gateway.workspaces()
            workspace_id = settings.clickup_workspace_id or (str(workspaces[0]["id"]) if workspaces else "")
            if not workspace_id:
                st.error("لم يتم العثور على Workspace في ClickUp.")
                return None, False
            spaces = gateway.spaces(workspace_id)
        finally:
            gateway.close()
    except ClickUpCollectionError as exc:
        st.error(str(exc))
        return None, False
    space_map = {str(item["id"]): item for item in spaces}
    selected = st.selectbox("Space", [None, *space_map],
                            format_func=lambda value: "Choose a space" if value is None else space_map[value].get("name", value),
                            key="clickup_selected_space")
    if selected is None:
        return None, False
    try:
        gateway = ClickUpGateway(settings.clickup_token)
        try:
            with st.spinner("جاري تحميل مهام ClickUp..."):
                tasks = gateway.all_tasks_for_space(selected)
        finally:
            gateway.close()
    except ClickUpCollectionError as exc:
        st.error(str(exc))
        return None, False
    rows = []
    for task in tasks:
        status = task.get("status") or {}
        priority = task.get("priority") or {}
        assignees = task.get("assignees") or []
        rows.append({"Task ID": task.get("id"), "Task Name": task.get("name"),
                     "Assignee": ", ".join(str(a.get("username") or a.get("email") or a.get("id")) for a in assignees) or "Unassigned",
                     "Priority": priority.get("priority") if isinstance(priority, dict) else None,
                     "Status": status.get("status") if isinstance(status, dict) else None,
                     "Due date": task.get("due_date")})
    st.subheader("All work items")
    st.caption(f"Selected space: {space_map[selected].get('name', selected)} · {len(rows)} work items shown")
    st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
    st.divider()
    fingerprint = f"clickup:{selected}:{len(tasks)}"
    if st.button("Done", type="primary", key="clickup_done"):
        st.session_state.pop("clickup_error", None)
        try:
            with st.status("Collecting ClickUp Activity...", expanded=True) as status:
                prepared = collect_clickup_data(ClickUpGateway(settings.clickup_token), tasks,
                    space_map[selected].get("name", selected), fingerprint, settings.source_timezone,
                    progress=lambda message: status.update(label=message))
                st.session_state["prepared_data"] = prepared
                status.update(label="ClickUp data collection completed.", state="complete", expanded=False)
        except Exception as exc:
            st.session_state["clickup_error"] = str(exc)
    if st.session_state.get("clickup_error"):
        st.error(st.session_state["clickup_error"])
    prepared = st.session_state.get("prepared_data")
    if prepared:
        st.success("ClickUp data has been collected and is ready for analysis.")
        st.download_button("Download Source Excel", prepared.xlsx, prepared.filename,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", on_click="ignore")
    return prepared, False


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
    _restore_collection_prompt(settings)
    if _advance_collection_job(settings):
        return st.session_state.get("prepared_data"), False
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
            store = _collection_store(settings)
            job = CollectionJob(settings, space, query, fingerprint,
                                st.session_state["jira_fields"], JiraGateway,
                                store=store, owner_key=persistence_owner_key(settings))
            st.session_state["collection_job"] = job
            job.start()
            job.step()
            if job.result is not None:
                st.session_state["prepared_data"] = job.result
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

