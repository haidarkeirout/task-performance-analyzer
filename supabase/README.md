# Persistent collection store contract

The application calls these PostgREST RPC functions with the publishable key:

| RPC | Purpose |
|---|---|
| `collection_begin` | Create or rehydrate a collection job and its fingerprint. |
| `collection_seed_items` | Store the complete ordered Jira issue list. |
| `collection_save_item` | Atomically persist one issue payload and its complete history. |
| `collection_update_job` | Update status, stage, progress, diagnostics, and result metadata. |
| `collection_load_job` | Load one owner-scoped job. |
| `collection_latest_resumable` | Find the latest running/error/paused job for one owner. |
| `collection_latest_for_fingerprint` | Find a recent completed job for one owner and selection fingerprint. |

The database implementation must keep jobs and items in private tables and expose
only the RPCs. Every RPC must require the 64-character `p_owner_key` capability,
filter every read and write by it, and never return another owner's rows. Direct
table access for `anon` and `authenticated` must remain revoked. The publishable
key is not a secret; the owner key is derived server-side and must never be sent
to the browser or logged.

The expected job response contains `job_id`, `owner_key`-scoped `space`, `query`,
`fingerprint`, `definitions`, `cutoff`, status fields, and an ordered `items`
array. Each item contains `position`, `issue_key`, `seed_item`, `completed`,
`item_data`, and `history_data`. `collection_save_item` must commit the item before
the client acknowledges it in memory.

Before deploying a fresh Supabase project, provision these functions through the
Supabase SQL editor or an audited migration, enable RLS on the backing tables,
revoke direct Data API table access, and run a test collection plus a forced retry.
This repository intentionally does not guess or silently apply database DDL to a
production project.
