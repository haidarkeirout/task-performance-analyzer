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
            "employee_ai_answer_fingerprint",
            "employee_ai_messages",
            "employee_ai_open",
        ):
            state.pop(key, None)
        state["employee_ai_answer_fingerprint"] = fingerprint


def _render_employee_ai_widget(
    st: Any,
    *,
    context: EmployeeAIContext,
    webhook_url: str,
    fingerprint: str,
) -> None:
    """Render a floating side drawer without obscuring the analysis page."""

    @st.fragment
    def assistant_widget() -> None:
        st.markdown(
            """
            <style>
            .st-key-employee_ai_launcher {
                position: fixed;
                z-index: 1000000;
                bottom: 24px;
                left: 24px;
                width: 64px;
            }
            .st-key-employee_ai_launcher button {
                width: 64px;
                height: 64px;
                min-height: 64px;
                padding: 0;
                border: 0;
                border-radius: 19px;
                color: #ffffff;
                background: #f26a2d;
                box-shadow: 0 12px 24px rgba(15, 23, 42, 0.22);
                font-size: 1.8rem;
                font-weight: 400;
                line-height: 1;
            }
            .st-key-employee_ai_launcher button:hover {
                background: #db5720;
                color: #ffffff;
            }
            .st-key-employee_ai_drawer {
                position: fixed;
                z-index: 999999;
                bottom: 104px;
                left: 24px;
                width: min(500px, calc(100vw - 48px));
                max-height: calc(100vh - 132px);
                overflow: auto;
                padding: 0.8rem 1rem 1rem;
                border: 1px solid #e5e7eb;
                border-radius: 20px;
                background: #ffffff;
                box-shadow: 0 22px 60px rgba(15, 23, 42, 0.24);
            }
            .st-key-employee_ai_drawer [data-testid="stVerticalBlockBorderWrapper"] {
                border: 0;
            }
            .st-key-employee_ai_drawer_close {
                position: absolute;
                z-index: 1;
                top: 0.95rem;
                right: 1rem;
                width: 2.3rem;
            }
            .employee-ai-header {
                margin: -0.8rem -1rem 1rem;
                padding: 1.25rem 1.4rem;
                border-radius: 19px 19px 0 0;
                color: #f8fafc;
                background: linear-gradient(120deg, #111827, #26344e);
            }
            .employee-ai-header__eyebrow {
                margin: 0 0 0.35rem;
                color: #fbbf24;
                font-size: 0.72rem;
                font-weight: 700;
                letter-spacing: 0.14em;
            }
            .employee-ai-header__title {
                margin: 0;
                color: #ffffff;
                font-size: 1.35rem;
                font-weight: 700;
            }
            .employee-ai-header__copy {
                margin: 0.45rem 0 0;
                color: #cbd5e1;
                font-size: 0.9rem;
            }
            .employee-ai-empty {
                display: flex;
                min-height: 250px;
                align-items: center;
                justify-content: center;
                padding: 2rem;
                border: 1px solid #e5e7eb;
                border-radius: 16px;
                background: linear-gradient(145deg, #ffffff, #f8fafc);
                color: #64748b;
                text-align: center;
            }
            .employee-ai-empty h3 {
                margin: 0.35rem 0;
                color: #1e293b;
                font-size: 1.15rem;
            }
            .employee-ai-empty p {
                margin: 0;
            }
            .st-key-employee_ai_drawer [data-testid="stChatInput"] {
                margin-top: 0.8rem;
            }
            .st-key-employee_ai_drawer [data-testid="stChatInput"] textarea {
                min-height: 2.8rem;
            }
            .st-key-employee_ai_drawer_close button {
                min-height: 2.2rem;
                padding: 0.1rem 0.5rem;
                border: 0;
                color: #ffffff;
                background: transparent;
                font-size: 1.8rem;
                line-height: 1;
            }
            .st-key-employee_ai_drawer_close button:hover {
                color: #fbbf24;
                background: transparent;
            }
            @media (max-width: 640px) {
                .st-key-employee_ai_launcher {
                    bottom: 16px;
                    left: 16px;
                }
                .st-key-employee_ai_drawer {
                    bottom: 96px;
                    left: 16px;
                    width: calc(100vw - 32px);
                    max-height: calc(100vh - 120px);
                }
            }
            </style>
            """,
            unsafe_allow_html=True,
        )

        state = st.session_state
        is_open = bool(state.get("employee_ai_open"))
        with st.container(key="employee_ai_launcher"):
            launcher_label = "×" if is_open else "✦"
            launcher_help = "Close AI assistant" if is_open else "Open AI assistant"
            if st.button(launcher_label, key="employee_ai_launcher_button", help=launcher_help):
                state["employee_ai_open"] = not is_open
                st.rerun(scope="fragment")

        if not is_open:
            return

        with st.container(key="employee_ai_drawer"):
            st.markdown(
                """
                <section class="employee-ai-header">
                  <p class="employee-ai-header__eyebrow">EMPLOYEE / AI</p>
                  <p class="employee-ai-header__title">AI Assistant</p>
                  <p class="employee-ai-header__copy">Ask questions about this employee's current analysis.</p>
                </section>
                """,
                unsafe_allow_html=True,
            )
            with st.container(key="employee_ai_drawer_close"):
                if st.button("×", key="employee_ai_drawer_close_button", help="Close AI assistant"):
                    state["employee_ai_open"] = False
                    st.rerun(scope="fragment")

            messages = state.setdefault("employee_ai_messages", [])
            conversation = st.container(height=390, border=True, key="employee_ai_conversation")
            with conversation:
                if not messages:
                    st.markdown(
                        """
                        <section class="employee-ai-empty">
                          <div>
                            <div style="font-size: 2rem;">◌</div>
                            <h3>Start the conversation</h3>
                            <p>Your questions and the assistant's answers will appear here.</p>
                          </div>
                        </section>
                        """,
                        unsafe_allow_html=True,
                    )
                for message in messages:
                    role = message.get("role")
                    content = str(message.get("content") or "")
                    if role in {"user", "assistant"} and content:
                        with st.chat_message(role):
                            st.markdown(content)

            question = st.chat_input(
                "Ask about the current analysis…",
                key=f"employee_ai_question_{fingerprint}",
            )
            if question:
                messages.append({"role": "user", "content": question})
                try:
                    with st.spinner("Reviewing the current analysis..."):
                        result = ask_employee_ai(
                            context=context,
                            question=question,
                            webhook_url=webhook_url,
                        )
                    messages.append({"role": "assistant", "content": result["answer"]})
                    st.rerun(scope="fragment")
                except EmployeeAIWebhookError as exc:
                    st.error(str(exc))

    assistant_widget()

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

    _render_employee_ai_widget(
        st,
        context=context,
        webhook_url=webhook_url,
        fingerprint=fingerprint,
    )


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
