"""Automatic dual-source collection for Company Performance.

Company scope is intentionally source-agnostic: all Jira projects and all
ClickUp Spaces are collected, then each source-space is assigned a stable
project label for the company-level analysis.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from typing import Any, Mapping, MutableMapping

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

ALL_COMPANIES_ID = "__all_companies__"
COMPANY_SOURCE_MAPPINGS_KEY = "company_source_mappings"


@dataclass(frozen=True)
class CompanySourceRef:
    """One stable source object discovered for a company."""

    source_tool: str
    source_id: str
    source_name: str
    item: Mapping[str, Any]


@dataclass(frozen=True)
class CompanyCatalogEntry:
    """A user-selectable company assembled from source Space/Project records."""

    company_id: str
    company_name: str
    sources: tuple[CompanySourceRef, ...]
    warnings: tuple[str, ...] = ()

    @property
    def has_jira(self) -> bool:
        return any(source.source_tool == "Jira" for source in self.sources)

    @property
    def has_clickup(self) -> bool:
        return any(source.source_tool == "ClickUp" for source in self.sources)

    @property
    def source_summary(self) -> str:
        tools = []
        if self.has_jira:
            tools.append("Jira")
        if self.has_clickup:
            tools.append("ClickUp")
        return " + ".join(tools) or "No source"


def _company_match_key(value: str | None) -> str:
    """Build a conservative cross-source key without fuzzy name matching."""
    return " ".join(str(value or "").strip().split()).casefold()


def _slug(value: str | None) -> str:
    # ``\w`` keeps Arabic and other Unicode company names usable as stable
    # session identifiers instead of collapsing every non-Latin name to the
    # same literal ``company`` value.
    cleaned = re.sub(r"[^\w]+", "-", str(value or "").casefold(), flags=re.UNICODE).strip("-")
    return cleaned or "company"


def company_option_label(entry: CompanyCatalogEntry) -> str:
    return f"{entry.company_name} ({entry.source_summary})"


def _source_mapping_key(source_tool: str, source_id: str) -> str:
    """Return the stable, platform-qualified key used by the mapping registry."""
    return f"{str(source_tool).casefold()}:{str(source_id)}"


def _source_mapping_keys(ref: CompanySourceRef) -> tuple[str, ...]:
    """Return primary and platform-native aliases for one source record."""
    candidates = [ref.source_id]
    if ref.source_tool == "Jira":
        candidates.append(str(ref.item.get("key") or ""))
    return tuple(
        dict.fromkeys(
            _source_mapping_key(ref.source_tool, candidate)
            for candidate in candidates
            if str(candidate).strip()
        )
    )


def _mapping_record(value: Any) -> tuple[str, str, tuple[str, ...]] | None:
    """Normalize a stored mapping while accepting the initial string format."""
    if isinstance(value, str):
        company_id = value.strip()
        if not company_id:
            return None
        return company_id, company_id, ()
    if not isinstance(value, Mapping):
        return None
    company_id = str(value.get("company_id") or "").strip()
    company_name = str(value.get("company_name") or company_id).strip()
    if not company_id:
        return None
    aliases = value.get("aliases") or ()
    if isinstance(aliases, str):
        aliases = (aliases,)
    return company_id, company_name, tuple(str(alias).strip() for alias in aliases if str(alias).strip())


def _source_mapping_registry() -> dict[str, Any]:
    raw = st.session_state.get(COMPANY_SOURCE_MAPPINGS_KEY) or {}
    return dict(raw) if isinstance(raw, Mapping) else {}


def remember_company_source_mapping(
    session_state: MutableMapping[str, Any],
    source_entry: CompanyCatalogEntry,
    target_entry: CompanyCatalogEntry,
) -> dict[str, Any]:
    """Persist an explicit cross-platform company mapping in session state.

    Jira keys/project IDs and ClickUp Space IDs are intentionally kept in
    separate namespaces.  A mapping joins those IDs to the selected company;
    it never assumes that an ID from one platform can equal an ID from the
    other platform.
    """
    if source_entry.company_id == target_entry.company_id:
        raise ValueError("A company source cannot be mapped to itself.")
    registry = dict(session_state.get(COMPANY_SOURCE_MAPPINGS_KEY) or {})
    aliases = tuple(
        dict.fromkeys(
            [
                target_entry.company_name,
                *(source.source_name for source in source_entry.sources),
                *(source.source_name for source in target_entry.sources),
            ]
        )
    )
    record = {
        "company_id": target_entry.company_id,
        "company_name": target_entry.company_name,
        "aliases": aliases,
    }
    for entry in (source_entry, target_entry):
        for source in entry.sources:
            for mapping_key in _source_mapping_keys(source):
                registry[mapping_key] = record
    session_state[COMPANY_SOURCE_MAPPINGS_KEY] = registry
    return registry


def clear_company_catalog_cache() -> None:
    """Invalidate the catalog after a mapping is changed in the UI."""
    st.session_state.pop("company_catalog", None)
    st.session_state.pop("company_catalog_revision", None)


def render_company_source_mapping(catalog: tuple[CompanyCatalogEntry, ...]) -> None:
    """Render an optional UI for joining differently named source records.

    Exact normalized names continue to merge automatically.  This control is
    only needed when, for example, Jira calls a project ``Najm Al-Shamal
    Website`` while ClickUp calls its Space ``Najm Al-Shamal``.
    """
    source_candidates = [entry for entry in catalog if len(entry.sources) == 1]
    target_candidates = list(catalog)
    if not source_candidates or len(target_candidates) < 2:
        return
    with st.expander("Map Jira / ClickUp sources to one company", expanded=False):
        st.caption(
            "Use this only when the same company has different names. "
            "Platform IDs remain independent and are linked explicitly."
        )
        labels = {
            entry.company_id: (
                f"{entry.company_name} — "
                + ", ".join(
                    f"{source.source_tool}:{source.source_id}"
                    for source in entry.sources
                )
            )
            for entry in target_candidates
        }
        source_id = st.selectbox(
            "Source to map",
            [entry.company_id for entry in source_candidates],
            format_func=lambda value: labels[value],
            key="company_mapping_source",
        )
        target_options = [entry.company_id for entry in target_candidates if entry.company_id != source_id]
        target_id = st.selectbox(
            "Existing company target",
            target_options,
            format_func=lambda value: labels[value],
            key="company_mapping_target",
        )
        if st.button("Save company source mapping", key="save_company_source_mapping"):
            source_entry = next(entry for entry in source_candidates if entry.company_id == source_id)
            target_entry = next(entry for entry in target_candidates if entry.company_id == target_id)
            remember_company_source_mapping(st.session_state, source_entry, target_entry)
            clear_company_catalog_cache()
            for key in (
                "company_selected_company",
                "company_active_selection",
                "company_prepared_items",
                "company_collection_attempted",
                "company_collection_error",
                "company_collection_errors",
                "company_partial_prepared_items",
                "company_analysis",
            ):
                st.session_state.pop(key, None)
            st.success("Company source mapping saved. The catalog will be rebuilt.")
            st.rerun()


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


def _source_id(source: str, item: Mapping[str, Any], index: int) -> str:
    return str(item.get("id") or item.get("key") or f"{source.casefold()}-{index}")


def discover_company_catalog(settings) -> tuple[CompanyCatalogEntry, ...]:
    """Discover selectable companies from the currently connected sources.

    A ClickUp Space and a Jira Project with the same exact normalized name are
    grouped into one company.  A saved explicit mapping can join different
    names by their platform-qualified IDs (and Jira key aliases). The source
    IDs remain attached to the entry; names are only used to build the initial
    catalog, never as task identity. Ambiguous duplicate names within one
    source are kept separate.
    """
    mapping_registry = _source_mapping_registry()
    mapping_signature = json.dumps(mapping_registry, sort_keys=True, default=str)
    revision = f"{getattr(settings, 'revision', '')}|mappings:{mapping_signature}"
    if st.session_state.get("company_catalog_revision") == revision:
        cached = st.session_state.get("company_catalog")
        if cached is not None:
            return tuple(cached)

    jira_spaces, clickup_spaces = _load_catalogs(settings)
    raw_sources: list[tuple[str, dict[str, Any], int]] = [
        *[("Jira", item, index) for index, item in enumerate(jira_spaces.values(), 1)],
        *[("ClickUp", item, index) for index, item in enumerate(clickup_spaces.values(), 1)],
    ]
    raw_refs: list[CompanySourceRef] = []
    for source, item, index in raw_sources:
        source_name = str(
            item.get("name")
            or item.get("key")
            or item.get("id")
            or f"{source} Company {index}"
        ).strip()
        source_id = _source_id(source, item, index)
        raw_refs.append(CompanySourceRef(source, source_id, source_name, dict(item)))

    mapped_targets: dict[str, tuple[str, str, tuple[str, ...]]] = {}
    mapped_aliases: dict[str, set[str]] = {}
    group_names: dict[str, str] = {}
    for ref in raw_refs:
        mapped = next(
            (
                _mapping_record(mapping_registry.get(mapping_key))
                for mapping_key in _source_mapping_keys(ref)
                if mapping_key in mapping_registry
            ),
            None,
        )
        if mapped:
            company_id, company_name, aliases = mapped
            group_key = f"mapped:{company_id}"
            for mapping_key in _source_mapping_keys(ref):
                mapped_targets[mapping_key] = mapped
            group_names[group_key] = company_name
            mapped_aliases.setdefault(group_key, set()).update(
                _company_match_key(alias) for alias in aliases
            )
    groups: dict[str, list[CompanySourceRef]] = {}
    for ref in raw_refs:
        mapped = next(
            (mapped_targets.get(mapping_key) for mapping_key in _source_mapping_keys(ref) if mapping_key in mapped_targets),
            None,
        )
        if mapped:
            group_key = f"mapped:{mapped[0]}"
        else:
            normalized_name = _company_match_key(ref.source_name)
            matching_groups = [
                group_key
                for group_key, aliases in mapped_aliases.items()
                if normalized_name in aliases
            ]
            # An alias is applied only when it identifies one mapping target;
            # conflicting aliases stay separate instead of being guessed.
            group_key = matching_groups[0] if len(set(matching_groups)) == 1 else f"name:{normalized_name}"
            group_names.setdefault(group_key, ref.source_name)
        groups.setdefault(group_key, []).append(ref)

    entries: list[CompanyCatalogEntry] = []
    for group_key, refs in groups.items():
        if not refs:
            continue
        company_name = group_names.get(group_key, refs[0].source_name)
        counts: dict[str, int] = {}
        for ref in refs:
            counts[ref.source_tool] = counts.get(ref.source_tool, 0) + 1
        duplicate_source = any(count > 1 for count in counts.values())
        if duplicate_source:
            for ref in refs:
                entries.append(
                    CompanyCatalogEntry(
                        company_id=f"{_slug(company_name)}--{ref.source_tool.casefold()}-{_slug(ref.source_id)}",
                        company_name=f"{company_name} [{ref.source_tool}: {ref.source_name}]",
                        sources=(ref,),
                        warnings=("Duplicate source name kept separate",),
                    )
                )
            continue
        entries.append(
            CompanyCatalogEntry(
                company_id=_slug(company_name),
                company_name=company_name,
                sources=tuple(sorted(refs, key=lambda ref: (ref.source_tool, ref.source_id))),
            )
        )

    entries.sort(key=lambda entry: entry.company_name.casefold())
    st.session_state["company_catalog_revision"] = revision
    st.session_state["company_catalog"] = tuple(entries)
    return tuple(entries)


def _selected_catalog_entries(
    catalog: tuple[CompanyCatalogEntry, ...], selected_company: str,
) -> tuple[CompanyCatalogEntry, ...]:
    if selected_company == ALL_COMPANIES_ID:
        return catalog
    selected = tuple(entry for entry in catalog if entry.company_id == selected_company)
    if not selected:
        raise CollectionError("The selected company is no longer available. Refresh the company catalog.")
    return selected


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


def collect_company_spaces(
    settings,
    selected_company: str = ALL_COMPANIES_ID,
) -> tuple[CompanyPreparedItem, ...]:
    """Collect only the selected company, or every company for All Companies."""
    catalog = discover_company_catalog(settings)
    selected_entries = _selected_catalog_entries(catalog, selected_company)
    source_refs = [
        (entry, ref)
        for entry in selected_entries
        for ref in entry.sources
    ]
    sources: list[tuple[str, dict[str, Any]]] = [
        (ref.source_tool, dict(ref.item)) for _, ref in source_refs
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
    company_names: dict[str, str] = {}
    source_ids: dict[str, str] = {}
    source_kinds: dict[str, str] = {}
    for index, (entry, ref) in enumerate(source_refs, 1):
        identity = _source_identity(ref.source_tool, dict(ref.item), index)
        company_names[identity] = entry.company_name
        source_ids[identity] = ref.source_id
        source_kinds[identity] = "jira_project" if ref.source_tool == "Jira" else "clickup_space"

    previous_selection = st.session_state.get("company_collection_selection")
    if previous_selection != selected_company:
        st.session_state.pop("company_partial_prepared_items", None)
        st.session_state["company_collection_selection"] = selected_company
    selected_names = [entry.company_name for entry in selected_entries]
    st.session_state["company_collection_company_name"] = (
        "All Companies" if selected_company == ALL_COMPANIES_ID else selected_names[0]
    )
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
        company_name = company_names.get(identity, project_names[identity])
        project_name = company_name
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
            source_key = str(item.get("id") or item.get("key") or source_space)
            fingerprint = (
                f"company:{selected_company}:{source.casefold()}:{source_key}"
            )
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
                company_id=selected_company if selected_company != ALL_COMPANIES_ID else _slug(company_name),
                company_name=company_name,
                source_id=source_ids.get(identity),
                source_kind=source_kinds.get(identity),
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


def render_company_collection(settings, selected_company: str = ALL_COMPANIES_ID):
    """Automatically collect the full Company source scope once per session."""
    if st.session_state.get("company_analysis") is not None:
        return None, False
    if st.session_state.get("company_collection_selection") != selected_company:
        for key in (
            "company_prepared_items",
            "company_collection_attempted",
            "company_collection_error",
            "company_collection_errors",
            "company_partial_prepared_items",
        ):
            st.session_state.pop(key, None)
    if st.session_state.get("company_prepared_items"):
        return None, False

    if not st.session_state.get("company_collection_attempted"):
        st.session_state["company_collection_attempted"] = True
        try:
            st.session_state["company_prepared_items"] = collect_company_spaces(
                settings, selected_company=selected_company
            )
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
                "company_collection_selection",
            ):
                st.session_state.pop(key, None)
            st.rerun()
    return None, False
