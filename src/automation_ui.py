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
    """Keep an active persistent Jira collection continuous on Streamlit Cloud."""

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
    """Rehydrate and continue a durable running Jira job without user interaction."""
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


def _render_monitor_snapshot(job, snapshot):
    """Render one live collection snapshot for the session-owned job."""
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
        _jira_ui._render_collection_error(job, snapshot)
    elif snapshot["result"] is not None:
        if st.session_state.get("prepared_data") is not snapshot["result"]:
            st.session_state["prepared_data"] = snapshot["result"]
            st.session_state.pop("collection_notice", None)
            st.rerun()
        st.success("Collection complete. All selected work items were collected and checkpointed.")
        if st.button(
            "Start New Collection",
            width="stretch",
            key=f"new_collection_complete_{snapshot['id']}",
        ):
            invalidate_selection()
            st.session_state.pop("persistent_candidate_checked", None)
            st.session_state.pop("persistent_candidate", None)
            st.rerun()
    elif snapshot["status"] == "paused":
        st.info(
            f"A persistent collection was restored. {completed} of {total} work items are already saved. "
            "Resume continues from the first unfinished work item."
            if total
            else "A persistent collection was restored and is ready to resume."
        )
        resume_col, reset_col = st.columns(2)
        if resume_col.button(
            "Resume Collection",
            type="primary",
            width="stretch",
            key=f"resume_restored_{snapshot['id']}",
        ):
            job.start()
            st.rerun()
        if reset_col.button(
            "Start New Collection",
            width="stretch",
            key=f"new_restored_{snapshot['id']}",
        ):
            invalidate_selection()
            st.session_state.pop("persistent_candidate_checked", None)
            st.session_state.pop("persistent_candidate", None)
            st.rerun()
    else:
        st.warning("Collection is paused.")


@st.fragment(run_every=1)
def _session_collection_monitor():
    """Advance the canonical session job; never receive a mutable job argument.

    Streamlit fragments can rerun independently of their parent script. Reading the
    job from session_state on every fragment tick prevents the fragment from using a
    stale serialized/captured job while the parent page sees a newer live object.
    """
    job = st.session_state.get("collection_job")
    if job is None:
        st.warning("Collection session is being restored...")
        return

    snapshot = job.snapshot()
    if snapshot["running"]:
        job.step()
        snapshot = job.snapshot()

    _render_monitor_snapshot(job, snapshot)


def _render_active_job_session(job):
    """Render Jira collection using the session-owned job and argument-free fragment."""
    live_job = st.session_state.get("collection_job") or job
    st.subheader("Data Collection")
    space_name = live_job.space.get("name") or live_job.space.get("key") or "Jira space"
    st.caption(f"Persistent collection for {space_name} · Reference: {live_job.id}")

    # Critical: do not pass CollectionJob into st.fragment. The fragment reads the
    # canonical object directly from session_state on every tick.
    _session_collection_monitor()

    prepared = st.session_state.get("prepared_data")
    if prepared:
        st.success("Your data has been collected and is ready for analysis.")
        st.caption(
            f"{prepared.count} work items · {prepared.space_name} · "
            f"Collected at {prepared.collected_at}"
        )
        st.download_button(
            "Download Source Excel",
            prepared.xlsx,
            prepared.filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            on_click="ignore",
        )

    run_clicked = st.button("Run Analysis", type="primary", disabled=prepared is None)
    if prepared is None:
        snapshot = live_job.snapshot()
        if snapshot["running"]:
            st.caption(
                "Data collection is still running. Run Analysis will unlock when the Excel source is complete."
            )
        elif snapshot["error"]:
            st.caption("Resolve the collection error above. Saved progress will be reused when you retry.")
        elif snapshot["status"] == "paused":
            st.caption("Resume the saved collection to rebuild or finish the analysis source.")
    return prepared, run_clicked


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

    # Tests patch wrapper-level objects, so keep the Jira implementation synchronized.
    _jira_ui.JiraGateway = JiraGateway
    _jira_ui.CollectionStore = CollectionStore
    _jira_ui.CollectionJob = CollectionJob
    _jira_ui._render_active_job = _render_active_job_session

    _auto_restore_running_jira_job(settings)
    return _jira_ui.render_collection(settings)
