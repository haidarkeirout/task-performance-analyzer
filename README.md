# Jira Performance — Automation v3.1

Select a Jira space, apply filters, and prepare its data for the existing process
performance analysis. All application controls and messages are in English.

## Collection reliability update

Collection now runs in a session-owned background worker. Interface reruns do not
interrupt its HTTP requests. Completed task/history pairs are checkpointed in
server memory, and the progress counter counts completed tasks only. Each stage
and retry is shown with a diagnostic reference and recorded in server logs without
credentials or raw response bodies.

After a temporary failure, **Done** resumes the same selection. **Start new
collection** discards its checkpoints. A change of filters, space, or credentials
cancels the old job. Checkpoints older than 30 minutes are discarded on retry.
Checkpoints survive interface reruns within the same session; they do not survive
a server restart, a lost browser session, or Sign Out. After such an event, start a
new collection. This release does not add an external database.

The agreed Excel columns are fixed in `src/export_columns.py`. Required analysis
columns and full histories remain included. Missing/empty fields are listed in
`Source_Data_Quality`; values are never invented. Duplicate custom-field names
are distinguished by field ID when available. Filtering still selects tasks; it
does not control which columns are collected. Project metadata is fetched once.
Watcher identities are collected when Jira grants access, otherwise left blank
with a source-quality note; this does not block unrelated task analysis.

For an existing v3 installation, keep the existing Secrets and upload all contents
of this update, preserving the folder paths. No new credentials or dependencies
are needed. The sign-in page and analysis calculations are preserved.

## Daily use

1. **Sign In** with the single application account.
2. **Select a Space** from the spaces visible to the connected Jira account.
3. Review **All work items** and select filters. Use **More filters** to add other
   searchable Jira fields, including your custom fields. **JQL** is also available.
4. Click **Done**. The app collects every matching page and each task's full
   accessible history, then generates a real `.xlsx` workbook.
5. The message **Your data has been collected and is ready for analysis.** appears.
   **Download Source Excel** is optional; the file is already held by the app.
6. Click **Run Analysis**. The existing five dashboard tabs and the analysis Excel
   and Word downloads appear. No file download/reupload is required.

Changing the space or filters clears the previous prepared data and results.
Click Done again to collect the new selection. Sign Out clears the session.

## Set up once on the existing Streamlit app

**First save the connection settings, then upload the source update.** This lets
the new sign-in screen work as soon as the updated application starts.

### 1. Save the account and Jira connection

Open your app's **Manage app → Settings → Secrets**. Paste the following and
replace all five example values. If other secrets already exist, keep them and
add or update these keys without duplicate definitions.

```toml
APP_USERNAME = "<choose the application username>"
APP_PASSWORD = "<choose the application password>"
JIRA_BASE_URL = "https://your-site.atlassian.net"
JIRA_EMAIL = "<the connected Jira account email>"
JIRA_API_TOKEN = "<that Jira account's API token>"
```

The first two values define your one application account. They are separate from
your Jira sign-in. The last three values connect the backend to Jira and are never
requested from dashboard users. Store the real values in Streamlit Secrets only;
do not place them in a GitHub file or send them in a screenshot.

There is a placeholder template at `config_examples/streamlit_secrets.example.toml`.
It is documentation, and is not automatically loaded as a real configuration.

Use the Jira site root URL, with `https://` and no `/issues` or `/projects` suffix.
If your token has scopes and requires Atlassian's API gateway, also add:

```toml
JIRA_CLOUD_ID = "<the cloud ID for the same Jira site>"
```

The app then uses `api.atlassian.com/ex/jira/{cloudId}` for API requests while
keeping the normal site URL for task links. The connected account must be able
to browse the intended spaces and issues. API access is limited to the records
and fields that account can read. Replace the stored token when it expires or
is revoked; dashboard users never enter it.

### 2. Upload the update

Extract `jira_automation_v3_1.zip`. Upload its **contents** to the existing repository
root using **Add file → Upload files**, preserving all subfolders. Drag folders
as folders so file paths stay intact.

At the root you should have `app.py`, `README.md`, `requirements.txt`, `src/`,
`configs/`, `tests/`, `docs/`, and `config_examples/`. The archive also includes
the repository ignore rules. Do not place everything inside an extra parent folder.

Commit to the existing deployment branch, `main`. The Streamlit entry point stays
`app.py`. Wait for its normal redeployment, then open the app and sign in using
the application account configured in step 1. The new caption says
**Data collection v3.1** after sign-in.

### 3. Check the live connection

Select your existing Performance Analysis space. Check its basic filters and
**More filters**, collect a small known selection with **Done**, and run analysis.
The source Excel contains the selected JQL, task count, and collection timestamps
so you can reconcile the scope. The app does not create or edit Jira tasks.

## Optional administrator settings

These stay in Secrets; there is no settings sidebar for dashboard users.

| Setting | Default / purpose |
| --- | --- |
| `SOURCE_TIMEZONE` | `Asia/Damascus`; interpretation of source data and exported transition times. The existing business calendar remains configured in `configs/`. |
| `JIRA_START_DATE_FIELD_ID` | Auto-detect a field named Start date. Set its exact ID, such as `customfield_10015`, if your site has multiple populated fields with that name. |
| `JIRA_CLOUD_ID` | Empty; provide it for a scoped API token that uses the Atlassian gateway. |
| `APP_PASSWORD_HASH` | Optional PBKDF2 alternative to `APP_PASSWORD`; see `src/automation_auth.py`. The simple setup above does not require this. |

Changing the server account or connection configuration invalidates existing
authenticated sessions. Data and results are held per Streamlit session, with
no cross-user data cache. A new session requires sign-in and collection again.

## Filters and collection

- Basic controls are Space, Search work, Assignee, Type, Status, and Created.
- More filters is populated from Jira's project-scoped searchable-field metadata.
  It includes the other returned system/custom fields and their supported
  comparisons, rather than a fixed hand-written shortlist.
- Value suggestions come from Jira. Search values by name or enter an exact value.
  Numeric fields accept numbers; date fields offer ranges. The JQL option supports
  additional conditions and functions. The selected space always stays applied.
- These controls are implemented inside this application. Jira-specific app
  widgets or behaviors may require JQL; native Jira menus are not embedded or
  screen-scraped. Live parity of a tenant's custom fields must be checked on that site.
- Created/date filters select **tasks**. They do not truncate the selected tasks'
  status histories. Existing analysis still evaluates cumulatively through cutoff.
- Date filters follow the connected Jira account's time zone, shown above the list.
- Preview can load more pages. Done collects all matching pages, including those
  not yet displayed. Results use a cutoff frozen when collection starts.
- Incomplete pages, duplicate records, or a detected concurrent edit stop collection
  with an English message. Partial exports are not marked ready for analysis.

## Source Excel

The workbook is created from Jira REST API data. It is a genuine XLSX file, not a
CSV renamed to `.xlsx`. It contains all returned accessible fields. Its layout is
adapted to the existing analysis contract and is not a byte-for-byte recreation
of Jira's native CSV export.

| Sheet | Contents |
| --- | --- |
| `Jira_Data` | Task fields, using the headers required by the existing analyzer. |
| `Workflow_Events` | All retrieved status transitions for the selected tasks. |
| `Source_Data_Quality` | Missing or empty fixed source columns. |
| `History_Coverage` | Per-task completeness, coverage time, and initial status. |
| `Field_Changes` | Retrieved changes to all fields. |
| `Field_Catalog` | Jira field IDs, names, schemas, and corresponding Excel headers. |
| `Raw_JSON` | Structured task/history data preserved in numbered chunks, including full comments/worklogs available to the account. |
| `Process_Context` | Space, selected query, collection time, and evaluation cutoff. |

Attachment fields contain metadata and links; attachment binaries are not downloaded.
Long values are split into columns/chunks to respect Excel's cell limits. Source
text is written as text, including values that start with an equals sign.

## Existing analysis retained

The analysis engines, configuration files, report builder, and all original
dashboard functions are unchanged. The new collection adapter supplies an Excel
buffer and verified history JSON to the same `run_analysis` function.

Executive Dashboard, Process Analysis, Individual Achievements, Task Detail,
Data Quality, management averages, Weekly Task Flow, and both output reports use
the existing calculation rules. See `docs/ANALYSIS_REFERENCE.md` for definitions.

The existing calendar is Sunday–Thursday 09:00–17:00 Asia/Damascus. Its existing
workflow uses Idea, In Triage, To Do, In Progress, In Review, Done, and Rejected.
This automation release does not remap different workflows or change those rules.

## Verification and local development

```bash
python -m pip install -r requirements.txt
python -m unittest discover -s tests -v
python -m streamlit run app.py
```

For local development, put the five real configuration values in
`.streamlit/secrets.toml`, which is ignored by Git. The Cloud deployment uses the
Secrets setting described above.

Set `JIRA_TEST_WORKBOOK` to the original `Raw_Data_Jira(3).xlsx` to additionally run
the source-data reconciliation tests. Tests use simulated Jira responses and never
send requests to Jira. See `docs/VERIFICATION.md` for the release checks and limits.

Official references: [Streamlit Secrets](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/secrets-management),
[Jira authentication](https://developer.atlassian.com/cloud/jira/platform/basic-auth-for-rest-apis/),
[Jira search](https://developer.atlassian.com/cloud/jira/platform/rest/v3/api-group-issue-search/),
[Jira filter metadata and validation](https://developer.atlassian.com/cloud/jira/platform/rest/v3/api-group-jql/).

## Isolated ClickUp Activity History collector

ClickUp task inventory still uses the read-only ClickUp API. When the optional
browser session is configured, Activity History is collected through a headless
Chromium context using the same web requests as ClickUp's Activity panel. This
path is isolated to ClickUp; it is not imported by the Jira collector, Jira
analysis, or Jira Excel export.

For staging, keep the existing `CLICKUP_API_TOKEN` and `CLICKUP_WORKSPACE_ID`,
then configure one authenticated browser source in Streamlit Secrets:

```toml
# Local machine: a persistent profile created after signing in to ClickUp.
CLICKUP_BROWSER_PROFILE_DIR = "/path/to/clickup_probe_profile"

# Hosted deployment: storage state exported from that authenticated profile.
# Use a TOML multiline string for the JSON value.
CLICKUP_STORAGE_STATE_JSON = """<Playwright storage-state JSON>"""
```

`packages.txt` installs Chromium and `requirements.txt` installs the optional
Playwright client. Do not commit the profile or storage state to GitHub. If no
browser source is configured, ClickUp task collection still works and the app
falls back to the existing public/web request path; Jira remains unchanged.
