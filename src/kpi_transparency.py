"""Shared presentation helpers for explaining Completion Rate population."""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from typing import Any, Iterable, Mapping


@dataclass(frozen=True)
class CompletionRatePopulation:
    """The task population behind the Completion Rate denominator."""

    total_tasks: int
    included_tasks: int
    reasons: tuple[tuple[str, int], ...] = ()

    @property
    def excluded_tasks(self) -> int:
        return max(0, self.total_tasks - self.included_tasks)

    def tooltip(self) -> str:
        return f"{self.excluded_tasks} task(s) excluded from Completion Rate"


def build_population(
    total_tasks: int,
    included_tasks: int,
    reasons: Mapping[str, int] | None = None,
) -> CompletionRatePopulation:
    """Create a consistent population object and reconcile reason counts."""
    total = max(0, int(total_tasks))
    included = min(total, max(0, int(included_tasks)))
    excluded = total - included
    counts = Counter()
    for reason, count in (reasons or {}).items():
        value = max(0, int(count))
        if value:
            counts[str(reason)] += value

    explained = sum(counts.values())
    if explained < excluded:
        counts["Other KPI exclusion"] += excluded - explained
    elif explained > excluded:
        # The source-specific adapter may report overlapping quality flags.
        # Keep the displayed total mathematically consistent with the KPI.
        overflow = explained - excluded
        for reason in sorted(counts, reverse=True):
            reduction = min(overflow, counts[reason])
            counts[reason] -= reduction
            overflow -= reduction
            if overflow == 0:
                break

    ordered = tuple(sorted(
        ((reason, count) for reason, count in counts.items() if count),
        key=lambda item: item[0].casefold(),
    ))
    return CompletionRatePopulation(total, included, ordered)


def population_from_snapshots(snapshots: Iterable[Any]) -> CompletionRatePopulation:
    """Build the population used by the Company/Project snapshot dashboards."""
    items = [
        item for item in snapshots
        if getattr(item, "in_scope", True)
        and getattr(
            getattr(getattr(item, "task", None), "parent_classification", None),
            "value",
            None,
        ) != "Container Parent"
    ]
    included = []
    reasons: Counter[str] = Counter()
    for item in items:
        task = item.task
        if not getattr(task, "counted_in_kpis", False):
            classification = getattr(task.parent_classification, "value", "Subtask")
            reasons["Subtask" if classification == "Subtask" else classification] += 1
        elif getattr(
            getattr(item, "status_at_period_end", None), "value", None
        ) == "Unknown":
            reasons["Unknown status/history"] += 1
        else:
            included.append(item)
    return build_population(len(items), len(included), reasons)


def render_population_card(
    column: Any,
    population: CompletionRatePopulation,
    *,
    title: str = "Tasks Included in Completion Rate",
    detail_label: str | None = None,
) -> None:
    """Render a metric with hover explanation and clickable reason details."""
    column.metric(title, population.included_tasks, help=population.tooltip())
    label = detail_label or (
        f"{population.excluded_tasks} task(s) excluded from Completion Rate"
        if population.excluded_tasks
        else "No tasks excluded from Completion Rate"
    )
    popover = getattr(column, "popover", None)
    if callable(popover):
        details = popover(label, use_container_width=True, disabled=population.excluded_tasks == 0)
        with details:
            _render_details(details, population)
    else:
        details = column.expander(label, expanded=False)
        with details:
            _render_details(details, population)


def _render_details(container: Any, population: CompletionRatePopulation) -> None:
    container.markdown("**Exclusion details**")
    if not population.reasons:
        container.info("No tasks were excluded from this KPI.")
        return
    for reason, count in population.reasons:
        container.write(f"- {reason}: {count}")
    container.caption(f"Total excluded: {population.excluded_tasks}")
