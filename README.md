# Jira process performance analyzer — v2.0

Streamlit application for cumulative process evaluation through an explicit cutoff.
This update retains native Streamlit bar charts (no Plotly dependency).

## Deploy this update

Extract the ZIP. Upload **its contents** to the existing repository root, preserving
`app.py`, `src/`, `configs/`, `requirements.txt`, `tests/`, and this README.
Commit the upload to the branch used by Streamlit (`main`). The entry point remains
`app.py`. The app caption should show **Process analysis v2.0** after the update.

Do not upload a wrapper directory around these files: `src` and `configs` must be
siblings of `app.py`. No credentials or user datasets are included in this bundle.

## Run an analysis

1. Upload the Jira workbook.
2. Select a history source:
   - **Jira API**: enter site URL, email and API token. The client fetches the
     changelog pages and records coverage metadata. Credentials are not exported.
   - **Workbook transitions**: use the workbook's `Workflow_Events` sheet. Enter
     the history coverage end, and confirm completeness only if the log includes
     all transitions from creation through that moment and `Status` is its snapshot.
   - **History JSON**: use JSON produced by the updated `src/jira_client.py`.
     Legacy JSON without coverage metadata is considered unverified. Retrieve new
     history from Jira instead of labeling an unknown partial log complete.
3. Enter the exact evaluation cutoff (including time and timezone). Coverage must
   reach it. The default is the session's initial current time, not automatically
   the date in `Process_Context`. That original date is displayed for reference.
4. Review the process name, scope and dataset type, then run analysis.
5. Use Process Analysis for stage times, review-return rates, open work, evidence,
   definitions and transition audit trail. Download Excel and Word from Downloads.

Dates without times in source context are ambiguous. For testing the original
September 1–3 simulation, local reconciliation used `2026-09-03T23:59:59+03:00`
for cutoff and declared coverage. This is an explicit test assumption, not an
inferred timestamp or an assertion about later Jira activity.

## What changed

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
- Assignment Summary and Task Detail show created date, planned start date,
  actual start date, start schedule variance in days, due date, completion date,
  priority, current-status age, and overdue days. Positive start variance means
  the verified start was later than the planned date; negative means earlier.
- Overdue work is shown as a task-level table with assignee, priority, status,
  dates, and current age so leadership can identify follow-up items immediately.

## Definitions and boundaries

All rates use 0–100. Zero denominators are unavailable. Unknown-status tasks stay
in the uploaded in-scope total but not known-open/completed/rejected counts.
History-invalid tasks are excluded from transition metrics; coverage is disclosed.

This is a cumulative analysis from creation to cutoff, **not** a period-only event
filter. The original observation period remains source context, clearly separated
from the actual cutoff used.

Durations are process residence, not worklogs, productivity, or proof of labor.
Due dates, planned starts and assignees use the uploaded snapshot. Historical
schedule changes and completion ownership are **not reconstructed**; timeliness
must be read against that snapshot rather than an original baseline. No targets,
causes, employee rankings or performance scores are inferred.

The standard calendar remains Sunday–Thursday 09:00–17:00 Asia/Damascus, with
configured holidays. This update uses the existing workflow names: Idea, In Triage,
To Do, In Progress, In Review, Done, Rejected. Review returns are direct status
pairs; transition names such as "Request Changes" are not extra status nodes.

## Verification

Run with project dependencies installed:

```bash
python -m unittest discover -s tests -v
```

Set `JIRA_TEST_WORKBOOK` to the original six-sheet `Raw_Data_Jira(3).xlsx` to run
reconciliation and the offline application/export tests. Otherwise those tests
are skipped. No tests call Jira or require credentials.

Reconciliation at the declared test cutoff:

- 20 tasks; 10 done, 8 open, 2 rejected; WIP 4.
- 13 reviewed tasks; 2 rework events on 1 task (7.6923%).
- 2 replanning events on 2 tasks (15.3846%).
- 1 re-evaluation event on 1 task (7.6923%).
- 3 distinct tasks with any review exception (23.0769%).
- 87 actual transition records; SCRUM-27 has a blank event placeholder, not a transition.
- Mean elapsed execution 14.0845278h; mean elapsed lead time 30.6909167h.

Local tests cover calculations and both exports. A live Jira call and browser
rendering on Streamlit are deployment checks, not covered by the offline tests.
