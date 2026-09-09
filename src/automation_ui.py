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
    """Keep an active persistent Jira job alive across Streamlit fragment state loss.

    Streamlit Cloud can occasionally leave the in-session object non-running after
    a successful bounded collection step even though the durable job is still
    marked running. A restored browser/app session remains explicitly paused and
    still requires the existing Resume action.
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
    return _jira_ui.render_collection(settings)
