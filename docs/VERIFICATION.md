# Automation v3 — verification

Verified locally on 2026-09-08. This package has not been deployed to the user's
Streamlit app and contains no live Jira credentials.

## Results

**30 tests passed**, including the original workbook regression tests.

Runtime: Python 3.12.13, Streamlit 1.63.0, pandas 2.2.3, openpyxl 3.1.5,
requests 2.34.2, python-docx 1.2.0.

Command, from the project root:

```bash
JIRA_TEST_WORKBOOK='/path/to/Raw_Data_Jira(3).xlsx' python -m unittest discover -s tests -v
```

The private source workbook is not distributed with this package. Without that
variable, the four original-data tests are skipped; all other tests run offline.

## Coverage

- Correct and incorrect passwords, missing server configuration, optional password
  hashing, credential rotation, and no Jira calls before sign-in.
- The actual Streamlit UI using AppTest: Sign In, Space selection, Basic filters,
  More filters, custom numeric comparison, Done, source download, Run Analysis,
  all five result tabs, Excel/DOCX downloads, filter-change invalidation, Sign Out.
- Failed collection does not create a ready dataset or enable Run Analysis.
- Cursor pagination for search and offset pagination for spaces, histories,
  comments, and worklogs; server-capped page sizes and duplicate/partial pages.
- Concurrent-edit retry, equal-timestamp transition order, rate-limit retry,
  safe error messages, and disabled HTTP redirects for authenticated requests.
- Project-scoped query construction, quoted values, supported operators,
  custom fields, inclusive date ranges, and invalid JQL-condition rejection.
- Real XLSX generation, expected input headers, labels, planned-start mapping,
  duplicate field-name handling, formula-like text, and lossless long JSON chunks.
- Entire per-task metric frame and overall aggregation are identical when the
  original source fixture is converted into the new API-generated input format
  and processed at the same declared cutoff.
- Original workflow, cutoff, weekly flow, overdue denominators, history coverage,
  repeated stages, review returns, reconciliation, and both report exports.
- Python source compilation passed.

## Preserved code

Eight original analysis/client/configuration files remain byte-identical to the
working baseline. Their SHA-256 hashes are in `PRESERVED_ANALYSIS_SHA256.json`.
All 18 original functions in `app.py`, including `run_analysis`, the dashboard
functions, and the download builders, are AST-identical to the baseline.
Only the app's collection/sign-in orchestration is replaced.

## Live deployment check still required

Jira requests were tested against deterministic simulated API responses, not the
user's live tenant. The source workbook regression uses real supplied fixture
contents. The Streamlit screen flow was executed with AppTest, not a cloud browser.

After configuring Secrets and uploading this package, verify the connected space
list, custom More filters, and one known filtered selection on the actual Jira
site. Account permissions and token scopes govern which records/fields are
available. Third-party Jira filter widgets can differ from the metadata-driven
controls; JQL is available for supported conditions.

Jira search is not a transactionally frozen database snapshot. Collection detects
reported duplicate/page-total changes and per-issue edits and freezes an analysis
cutoff; edits to other query membership during a long collection can still affect
which tasks Jira returns. Recollect after such changes.

Existing Streamlit dashboard code emits deprecation notices for
`use_container_width` under this runtime. These did not fail the tests and that
original dashboard code was preserved.
