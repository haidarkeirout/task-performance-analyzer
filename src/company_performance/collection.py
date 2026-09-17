"""Automatic dual-source collection for Company Performance.

Company scope is intentionally source-agnostic: all Jira projects and all
ClickUp Spaces are collected, then each source-space is assigned a stable
project label for the company-level analysis.
"""

from __future__ import annotations

import re
from typing import Any

import streamlit as st

from clickup_gateway import ClickUpCollectionError
from jira_gateway import CollectionError
from project_ui import (
    _collect_clickup_space,
    _collect_jira_space,
    _load_catalogs,
)

from .application import CompanyPreparedItem


_PREFIX_RE = re.compile(
    r"^(?:website|project|platform|application|app|web|mobile)"
    r"(?:\s*[-_:|/]\s*|\s+)",
    re.IGNORECASE,
)


def canonical_project_name(space_name: str | None) -> str:
    """Return the company-facing project label for one source Space name.

    This deliberately uses a conservative rule instead of fuzzy matching:
    exact names stay unchanged, while common descriptive prefixes such as
    Website are removed. Ambiguous names remain visible as separate projects
    for later review.
    """
    value = " ".join(str(space_name or "").strip().split())
    while True:
        cleaned = _PREFIX_RE.sub("", value, count=1).strip(" -_:|/")
        if cleaned == value:
            break
        value = cleaned
    return value or "Unmapped Project"


def _source_identity(source: str, item: dict[str, Any], index: int) -> str:
    return f"{source.casefold()}:{item.get('id') or item.get('key') or index}"


def _project_names(sources: list[tuple[str, dict[str, Any]]]) -> dict[str, str]:
    """Resolve conservative project names without silently merging ambiguity.

    One Jira project and one ClickUp Space with the same canonical label are
    treated as the same cross-tool project.  Any collision within a source is
    ambiguous, so every colliding Space keeps a visibly distinct label.
    """
    candidates: list[tuple[str, str, str, str]] = []
    for index, (source, item) in enumerate(sources, 1):
        source_space = str(item.get("name") or item.get("key") or item.get("id") or f"{source} Space {index}")
        candidates.append(
            (_source_identity(source, item, index), source, source_space, canonical_project_name(source_space))
        )

    groups: dict[str, list[tuple[str, str, str, str]]] = {}
    for candidate in candidates:
        groups.setdefault(candidate[3].casefold(), []).append(candidate)

    resolved: dict[str, str] = {}
    for group in groups.values():
        sources_in_group = [candidate[1].casefold() for candidate in group]
        ambiguous = len(sources_in_group) != len(set(sources_in_group))
        for identity, source, source_space, canonical in group:
            resolved[identity] = (
                f"{canonical} [{source}: {source_space}]" if ambiguous else canonical
            )
    return resolved


def collect_company_spaces(settings) -> tuple[CompanyPreparedItem, ...]:
    """Collect every accessible Jira project and ClickUp Space."""
    jira_spaces, clickup_spaces = _load_catalogs(settings)
    sources: list[tuple[str, dict[str, Any]]] = [
        *[("Jira", item) for item in jira_spaces.values()],
        *[("ClickUp", item) for item in clickup_spaces.values()],
    ]
    if not sources:
        jira_error = st.session_state.get("project_jira_error", "")
        clickup_error = st.session_state.get("project_clickup_error", "")
        details = " ".join(value for value in (jira_error, clickup_error) if value)
        raise CollectionError(
            "No Jira projects or ClickUp Spaces are available for this connection."
            + (f" {details}" if details else "")
        )

    project_names = _project_names(sources)
    overall = st.progress(0, text=f"Collecting Company spaces: 0 / {len(sources)}")
    saved = dict(st.session_state.get("company_partial_prepared_items") or {})
    valid_identities = {
        _source_identity(source, item, index)
        for index, (source, item) in enumerate(sources, 1)
    }
    saved = {key: value for key, value in saved.items() if key in valid_identities}
    failures: list[str] = []

    for index, (source, item) in enumerate(sources, 1):
        identity = _source_identity(source, item, index)
        source_space = str(
            item.get("name")
            or item.get("key")
            or item.get("id")
            or f"{source} Space {index}"
        )
        project_name = project_names[identity]
        if identity in saved:
            overall.progress(
                index / len(sources),
                text=f"Reusing completed {source} Space {index} / {len(sources)} — {source_space}",
            )
            continue
        overall.progress(
            (index - 1) / len(sources),
            text=f"Collecting {source} Space {index} / {len(sources)} — {source_space}",
        )

        def update_source_progress(fraction: float, message: str) -> None:
            try:
                normalized_fraction = max(0.0, min(1.0, float(fraction)))
            except (TypeError, ValueError):
                normalized_fraction = 0.0
            overall.progress(
                ((index - 1) + normalized_fraction) / len(sources),
                text=f"{source_space}: {message}",
            )

        try:
            fingerprint = f"company:{source.casefold()}:{item.get('id') or item.get('key') or source_space}"
            if source == "Jira":
                source_data = _collect_jira_space(
                    settings,
                    item,
                    fingerprint,
                    progress=update_source_progress,
                )
            else:
                source_data = _collect_clickup_space(
                    settings,
                    item,
                    fingerprint,
                    progress=update_source_progress,
                )
            saved[identity] = CompanyPreparedItem(
                prepared=source_data,
                source_tool=source,
                source_space=source_space,
                project_name=project_name,
            )
            st.session_state["company_partial_prepared_items"] = dict(saved)
        except (CollectionError, ClickUpCollectionError) as exc:
            failures.append(f"{source} — {source_space}: {exc}")

    prepared = [saved[key] for key in sorted(saved)]
    overall.progress(
        1.0,
        text=f"Company collection complete: {len(prepared)} / {len(sources)} spaces",
    )
    st.session_state["company_collection_errors"] = tuple(failures)
    if failures or len(prepared) != len(sources):
        raise CollectionError(
            "Company collection is incomplete, so no partial dashboard was created. "
            "Retry to continue only the failed Spaces."
            + (f" {' | '.join(failures)}" if failures else "")
        )
    st.session_state.pop("company_partial_prepared_items", None)
    return tuple(prepared)


def render_company_collection(settings):
    """Automatically collect the full Company source scope once per session."""
    if st.session_state.get("company_analysis") is not None:
        return None, False
    if st.session_state.get("company_prepared_items"):
        return None, False

    if not st.session_state.get("company_collection_attempted"):
        st.session_state["company_collection_attempted"] = True
        try:
            st.session_state["company_prepared_items"] = collect_company_spaces(settings)
        except (CollectionError, ClickUpCollectionError) as exc:
            st.session_state["company_collection_error"] = str(exc)

    error = st.session_state.get("company_collection_error", "")
    if error:
        st.error(error)
        if st.button("Retry Company Collection", key="retry_company_collection"):
            for key in (
                "company_collection_attempted",
                "company_collection_error",
                "company_collection_errors",
                "company_prepared_items",
            ):
                st.session_state.pop(key, None)
            st.rerun()
    return None, False
