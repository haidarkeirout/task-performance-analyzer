"""Non-invasive Streamlit bridge for the Employee AI context prototype."""

from __future__ import annotations

import secrets
from typing import Any, Mapping, MutableMapping

from .client import EmployeeAIWebhookError, ask_employee_ai, configured_webhook_url
from .context_builder import (
    build_clickup_employee_context,
    build_combined_employee_context,
    build_jira_employee_context,
)
from .schemas import ContextValidationError, EmployeeAIContext


def _record_value(record: Any, name: str) -> Any:
    if isinstance(record, Mapping):
        return record.get(name)
    return getattr(record, name, None)


def _employee_from_state(state: Mapping[str, Any]) -> dict[str, Any]:
    snapshot = state.get("employee_snapshot") or {}
    record = snapshot.get("employee_record") if isinstance(snapshot, Mapping) else None
    prepared = state.get("employee_prepared")
    display_name = (
        _record_value(record, "name")
        or getattr(prepared, "employee_name", None)
        or "Selected Employee"
    )
    position = (
        _record_value(record, "position")
        or _record_value(record, "jira_position")
        or _record_value(record, "clickup_position")
    )
    # The subject key remains local.  It can contain the source-qualified
    # selection key because only its HMAC digest enters the public context.
    subject_key = state.get("employee_snapshot_key") or str(display_name)
    return {
        "display_name": display_name,
        "department": _record_value(record, "department"),
        "position": position,
        "subject_key": subject_key,
    }


def _filters_from_state(state: Mapping[str, Any], *, combined: bool) -> dict[str, Any]:
    if combined:
        return {
            "company": list(state.get("employee_filter_companies") or []),
            "projects_spaces": list(state.get("employee_filter_spaces") or []),
            "statuses": list(state.get("employee_filter_statuses") or []),
            "task_types": list(state.get("employee_filter_task_types") or []),
            "priorities": list(state.get("employee_filter_priorities") or []),
            "due_date_states": list(state.get("employee_filter_due_states") or []),
        }
    return {
        "company": state.get("employee_filter_company") or "All",
        "projects_spaces": list(state.get("employee_filter_projects") or []),
        "statuses": list(state.get("employee_filter_status") or []),
        "task_types": list(state.get("employee_filter_task_type") or []),
        "priorities": list(state.get("employee_filter_priority") or []),
        "due_date_states": list(state.get("employee_filter_due_state") or []),
    }


def _collection_from_state(state: Mapping[str, Any]) -> dict[str, Any]:
    return {
        "collection_id": state.get("employee_collection_id"),
        "collected_at": state.get("employee_collection_collected_at"),
        "analysis_completed_at": state.get("employee_analysis_completed_at"),
    }


def _subject_secret(state: MutableMapping[str, Any]) -> str:
    """Keep a session-local pseudonym key out of the exported context."""

    key = "_employee_ai_subject_secret"
    if not state.get(key):
        state[key] = secrets.token_hex(32)
    return str(state[key])


def build_employee_context_from_state(
    state: MutableMapping[str, Any],
    *,
    source_mode: str,
    analysis_result: Any,
    dashboard_values: Mapping[str, Any] | None = None,
    question: str = "",
    max_evidence_rows: int = 25,
) -> EmployeeAIContext:
    """Build the active filtered Employee context without changing analysis data."""

    mode = str(source_mode).strip().casefold()
    common = {
        "employee": _employee_from_state(state),
        "collection": _collection_from_state(state),
        "analysis_filters": _filters_from_state(state, combined=mode == "combined"),
        "collection_spaces": list(state.get("employee_selected_spaces") or []),
        "question": question,
        "subject_secret": _subject_secret(state),
        "max_evidence_rows": max_evidence_rows,
    }
    if mode == "jira":
        if dashboard_values is None:
            raise ContextValidationError("Rendered Jira dashboard values are required")
        return build_jira_employee_context(
            **common,
            dashboard_values=dashboard_values,
            task_metrics=analysis_result,
            validation_warnings=state.get("validation_log") or [],
        )
    if mode == "clickup":
        return build_clickup_employee_context(**common, analysis_result=analysis_result)
    if mode == "combined":
        return build_combined_employee_context(**common, analysis_result=analysis_result)
    raise ContextValidationError(f"Unsupported Employee source mode: {source_mode}")


def _streamlit_secrets(st: Any) -> Mapping[str, Any]:
    try:
        return st.secrets.to_dict()
    except (FileNotFoundError, AttributeError):
        return {}


def _clear_stale_ai_answer(state: MutableMapping[str, Any], fingerprint: str) -> None:
    if state.get("employee_ai_answer_fingerprint") != fingerprint:
        for key in (
            "employee_ai_answer",
            "employee_ai_answer_question",
            "employee_ai_answer_fingerprint",
        ):
            state.pop(key, None)


def render_employee_ai_assistant(st: Any, context: EmployeeAIContext) -> None:
    """Render a small read-only assistant only after an Employee analysis exists."""

    payload = context.to_dict()
    state = st.session_state
    state["employee_ai_context"] = payload
    fingerprint = str(payload["request"]["context_fingerprint"])
    _clear_stale_ai_answer(state, fingerprint)

    try:
        webhook_url = configured_webhook_url(_streamlit_secrets(st))
    except EmployeeAIWebhookError:
        # Do not expose configuration details or interrupt the Employee dashboard.
        return
    if not webhook_url:
        return

    with st.popover("✨ Ask AI", use_container_width=False):
        st.caption("Ask about the current Employee analysis. The assistant is read-only.")
        question = st.text_area(
            "Your question",
            placeholder="For example: What may explain the open overdue tasks?",
            key=f"employee_ai_question_{fingerprint}",
            height=90,
        )
        if st.button("Ask Assistant", key=f"employee_ai_submit_{fingerprint}", type="primary"):
            try:
                with st.spinner("Reviewing the current analysis..."):
                    result = ask_employee_ai(
                        context=context,
                        question=question,
                        webhook_url=webhook_url,
                    )
                state["employee_ai_answer"] = result["answer"]
                state["employee_ai_answer_question"] = question.strip()
                state["employee_ai_answer_fingerprint"] = result["context_fingerprint"]
            except EmployeeAIWebhookError as exc:
                st.warning(str(exc))

        if state.get("employee_ai_answer_fingerprint") == fingerprint:
            question_text = state.get("employee_ai_answer_question")
            if question_text:
                st.caption(f"Question: {question_text}")
            st.markdown(str(state.get("employee_ai_answer") or ""))


def safely_render_employee_ai_assistant(
    st: Any,
    *,
    source_mode: str,
    analysis_result: Any,
    dashboard_values: Mapping[str, Any] | None = None,
) -> None:
    """Contain prototype failures so the existing dashboard remains available."""

    try:
        context = build_employee_context_from_state(
            st.session_state,
            source_mode=source_mode,
            analysis_result=analysis_result,
            dashboard_values=dashboard_values,
        )
        render_employee_ai_assistant(st, context)
    except Exception:
        st.session_state.pop("employee_ai_context", None)
        # The optional assistant must never disturb the core Employee analysis.


def safely_render_employee_context_preview(
    st: Any,
    *,
    source_mode: str,
    analysis_result: Any,
    dashboard_values: Mapping[str, Any] | None = None,
) -> None:
    """Backward-compatible alias retained for the prototype integration points."""

    safely_render_employee_ai_assistant(
        st,
        source_mode=source_mode,
        analysis_result=analysis_result,
        dashboard_values=dashboard_values,
    )
