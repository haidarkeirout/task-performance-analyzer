# Existing analysis reference

The automation release retains these calculation and dashboard definitions.

## Existing outputs

- One shared aggregation for dashboard, Excel and Word.
- Review-return denominators are distinct reviewed tasks with complete history.
  Rework, replanning, re-evaluation, and their union expose numerator/denominator.
- Open overdue rate counts only open tasks with known due dates.
- Stage statistics sum repeat visits per task and include elapsed and business time.
- Terminal-state residence is excluded from stage bottleneck comparisons.
- Open task age, current status age, overdue days and evidence-linked follow-up.
- Verified no-transition histories are valid; absent/partial/invalid histories
  produce unavailable metrics and explicit unknown status counts.
- Cutoff reconstruction uses initial status even before the first transition;
  future-created tasks are excluded and listed in Data Quality.
- Excel includes process context, transition log, stage summary, open work,
  definitions, quality findings, and recommendations in addition to task metrics
  and aggregate views. It contains static calculated values; rerun the app to update.
- Word focuses on process outcomes and includes a snapshot assignment distribution.
  Unassigned is a task group, never an individual achievement profile.
- The dashboard adds an `Open Tasks by Due Status` native chart and table for
  overdue, within-due-date, and missing-due-date open tasks. It separately shows
  tasks whose status cannot be verified, plus a separate count/table for completed
  tasks that finished after the supplied due date.
- The Executive Dashboard adds a native `Weekly Task Flow` line chart with
  `Tasks Opened` and `Tasks Completed`, grouped into Monday-starting weeks. The
  view supports all available weeks, the last 4 weeks, the last 12 weeks
  (quarter), and the last 52 weeks (year), and exports the weekly table with
  net and cumulative flow values.
- The Executive Dashboard shows leadership-level averages for completed work
  (execution time, lead time, and time to start) and delay/planning risk (late
  completion days, open overdue days, and start variance days). Averages with
  no qualifying tasks are shown as unavailable rather than misleading zeros.
- Assignment Summary and Task Detail show created date, planned start date,
  actual start date, start schedule variance in days, due date, completion date,
  priority, current-status age, and overdue days. Positive start variance means
  the verified start was later than the planned date; negative means earlier.
- Overdue work is shown as a task-level table with assignee, priority, status,
  dates, and current age so leadership can identify follow-up items immediately.

## Definitions and boundaries

All rates use 0–100. Zero denominators are unavailable. Unknown-status tasks stay
in the selected in-scope total but not known-open/completed/rejected counts.
History-invalid tasks are excluded from transition metrics; coverage is disclosed.

This is a cumulative analysis from creation to cutoff, **not** a period-only event
filter. The original observation period remains source context, clearly separated
from the actual cutoff used.

Durations are process residence, not worklogs, productivity, or proof of labor.
Due dates, planned starts and assignees use the collected snapshot. Historical
schedule changes and completion ownership are **not reconstructed**; timeliness
must be read against that snapshot rather than an original baseline. No targets,
causes, employee rankings or performance scores are inferred.

The standard calendar remains Sunday–Thursday 09:00–17:00 Asia/Damascus, with
configured holidays. This update uses the existing workflow names: Idea, In Triage,
To Do, In Progress, In Review, Done, Rejected. Review returns are direct status
pairs; transition names such as "Request Changes" are not extra status nodes.

