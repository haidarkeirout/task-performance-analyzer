# Performance Management — verification

Verified locally on 2026-09-17 with no live Jira, ClickUp, employee-directory, or
Supabase credentials.

## Results

```text
120 tests passed; 2 fixture-dependent tests skipped
Python 3.12 · Streamlit 1.64.0 · pandas 2.2.3 · openpyxl 3.1.5
requests 2.34.2 · python-docx 1.2.0
```

Run from the repository root:

```bash
PYTHONPATH=src:tests python -m unittest discover -s tests -p 'test_*.py' -v
```

The two skipped tests require the private `JIRA_TEST_WORKBOOK` fixture. They are
not bundled with the repository.

## Covered areas

- Sign-in configuration, PBKDF2 passwords, failed-login throttling, and secret-safe
  error messages.
- Jira pagination, history coverage, concurrent edits, rate-limit retry, and
  durable per-item checkpoint recovery.
- ClickUp pagination recovery and source-specific status/time-in-status handling.
- Employee, Project, Department, and Company period reconstruction, timezone
  boundaries, parent/subtask KPI rules, and unknown historical statuses.
- Company all-or-nothing collection, retrying only failed Spaces, conservative
  project mapping, and preview filtering.
- Excel/Word generation, report caching, Excel formula-injection protection, fixed
  workbook dependencies, and current Streamlit mode smoke tests.
- Full Python compilation and the root-only repository layout.

## Live deployment checks

Offline tests cannot verify tenant permissions, Supabase RPC provisioning, or the
OneDrive sharing policy. After deployment, run one known Jira and ClickUp sample at
each analysis scope, reconcile task counts/cutoff/period, and open both report
formats. Confirm the persistent-store retry by interrupting a small Jira run.
