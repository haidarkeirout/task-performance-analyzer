# Performance Management

Read-only Streamlit analytics for Jira and ClickUp at four scopes:

- Employee Performance
- Project Performance
- Department Performance
- Company Performance

The application collects source records, preserves their audit history, reconstructs
the selected analysis period, and produces dashboards plus Excel and Word reports.
It does not create or edit Jira or ClickUp work items.

## Reliability and calculation rules

- Jira checkpoints are persisted after every work item and resume from the first
  unfinished item after a rerun, process restart, or temporary API failure.
- ClickUp pagination checkpoints resume at the failed page during the active session.
- Company analysis is all-or-nothing: a failed Space blocks the dashboard while
  successfully collected Spaces remain available for the next retry.
- Calendar boundaries are interpreted in `SOURCE_TIMEZONE` (default
  `Asia/Damascus`) and normalized to UTC for comparisons.
- Historical results never substitute a current ClickUp status for an unknown past
  status. Affected metrics are displayed as unavailable.
- Parent and standalone tasks are KPI-eligible; subtasks remain visible but are not
  counted twice in parent-level rates.
- Generated Excel source strings are written as text, not executable formulas.

Metric definitions and known source limitations are documented in
`docs/ANALYSIS_REFERENCE.md` and in every generated report.

## Configuration

Copy the keys from `config_examples/streamlit_secrets.example.toml` into Streamlit
Secrets or equivalent environment variables. Never commit real credentials.

Required application settings:

```toml
APP_USERNAME = "<username>"
APP_PASSWORD_HASH = "<PBKDF2 hash>" # preferred; APP_PASSWORD is also supported
```

Configure Jira, ClickUp, or both:

```toml
JIRA_BASE_URL = "https://your-site.atlassian.net"
JIRA_EMAIL = "<service account email>"
JIRA_API_TOKEN = "<token>"

CLICKUP_API_TOKEN = "pk_..."
CLICKUP_WORKSPACE_ID = "<workspace id>"
```

The employee directory URL can be overridden with `EMPLOYEE_DIRECTORY_URL`.
Use a view-only workbook and avoid storing sensitive HR information in a public
share. The app caches the last validated directory briefly and reports when a
temporary failure forces use of that copy.

Persistent Jira resume uses Supabase settings when overriding the bundled service:

```toml
SUPABASE_COLLECTION_URL = "https://<project>.supabase.co"
SUPABASE_COLLECTION_PUBLISHABLE_KEY = "<publishable key>"
```

The required database contract is documented in `supabase/README.md`. Access is
capability-scoped by a server-derived owner key; never expose that owner key.

## Local development

```bash
python -m pip install -r requirements.txt
PYTHONPATH=src:tests python -m unittest discover -s tests -p 'test_*.py' -v
python -m streamlit run app.py
```

Python 3.12 is used in CI. Pushes to `main` and `feature/**`, plus pull requests,
run compilation and the full offline regression suite.

## Deployment checklist

1. Store credentials only in Streamlit Secrets.
2. Prefer `APP_PASSWORD_HASH`; failed sign-ins are throttled after five attempts.
3. Confirm the service accounts have read-only access to the intended projects.
4. Run a known Jira and ClickUp sample at all four analysis scopes.
5. Reconcile the source count, cutoff, period, and Data Quality section.
6. Download and open both Excel and Word reports.

Task/history data and generated reports live in the Streamlit session except for
Jira collection checkpoints in the configured persistent store. Sign Out clears
the active session but deliberately does not delete a resumable server checkpoint.
