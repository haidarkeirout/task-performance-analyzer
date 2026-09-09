"""Source-selecting UI wrapper that keeps Jira persistence and ClickUp isolated."""
from __future__ import annotations

import streamlit as st

import jira_ui as _jira_ui
from clickup_ui import _render_clickup_collection

# Re-export these symbols for existing tests and callers.
CollectionError = _jira_ui.CollectionError
CollectionStore = _jira_ui.CollectionStore
PersistenceError = _jira_ui.PersistenceError
persistence_owner_key = _jira_ui.persistence_owner_key
JiraGateway = _jira_ui.JiraGateway


class CollectionJob(_jira_ui.CollectionJob):
    """Keep an active persistent Jira collection continuous on Streamlit Cloud.

    A normal collection must behave as one uninterrupted operation after Done.
    If Streamlit leaves the in-session object idle while the durable job is still
    running, restart the same job automatically from its saved checkpoint.
    """

    def snapshot(self):
        snapshot = super().snapshot()
        if (
            snapshot["status"] == "idle"
            and snapshot.get("persisted_status") == "running"
            and snapshot.get("result") is None
            and not snapshot.get("error")
            and not self.cancelled
        ):
            self.start()
            snapshot = super().snapshot()
        return snapshot


require_sign_in = _jira_ui.require_sign_in
invalidate_selection = _jira_ui.invalidate_selection


def _auto_restore_running_jira_job(settings):
    """Rehydrate and continue a running Jira job without user interaction.

    This covers full-script reruns or Streamlit session-object recreation. Only a
    durable job whose persisted status is still ``running`` is auto-resumed.
    Error jobs keep the explicit Retry flow and completed jobs keep the existing
    result-rebuild flow.
    """
    if st.session_state.get("collection_job") is not None:
        return

    store = CollectionStore.configured()
    if store is None:
        return

    keep_store = False
    try:
        owner_key = persistence_owner_key(settings)
        payload = store.latest_resumable(owner_key)
        if not payload or payload.get("status") != "running":
            return

        job = CollectionJob.from_persisted(
            settings,
            payload,
            JiraGateway,
            store=store,
            owner_key=owner_key,
        )
        if job is None:
            return

        job.start()
        st.session_state["collection_job"] = job
        st.session_state["persistent_candidate_checked"] = True
        st.session_state.pop("persistent_candidate", None)
        st.session_state.pop("persistent_store_error", None)
        keep_store = True
    except PersistenceError as exc:
        st.session_state["persistent_store_error"] = str(exc)
    finally:
        if not keep_store:
            store.close()


def render_collection(settings):
    """Route to the selected connector without sharing collection state."""
    source = st.radio(
        "Data source",
        ["Jira", "ClickUp"],
        horizontal=True,
        key="data_source",
    )
    if source == "ClickUp":
        return _render_clickup_collection(settings)

    # Tests patch the wrapper-level JiraGateway. Keep the implementation synced
    # before rendering so production and deterministic test transports behave alike.
    _jira_ui.JiraGateway = JiraGateway
    _jira_ui.CollectionStore = CollectionStore
    _jira_ui.CollectionJob = CollectionJob

    # Normal Jira collection is continuous after Done. If Streamlit recreated the
    # session object, restore the still-running durable job before rendering Jira.
    _auto_restore_running_jira_job(settings)
    return _jira_ui.render_collection(settings)
