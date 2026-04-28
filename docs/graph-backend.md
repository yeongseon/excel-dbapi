# Microsoft Graph Backend

Access remote Excel workbooks on OneDrive and SharePoint via the
[Microsoft Graph API](https://learn.microsoft.com/en-us/graph/overview).

> **Status**: Experimental — API may change in future releases.

## Installation

```bash
# Base Graph support (bring your own token provider)
pip install excel-dbapi[graph]

# Includes Azure Identity for DefaultAzureCredential and other Azure credential flows
pip install excel-dbapi[graph-azure]
```

Choose `graph-azure` when you want excel-dbapi to acquire tokens through
`azure-identity` directly.

## DSN Formats

excel-dbapi supports three ID-based DSN patterns:

- `msgraph://drives/{drive_id}/items/{item_id}` — generic Graph endpoint
- `sharepoint://sites/{site_name}/drives/{drive_id}/items/{item_id}` — SharePoint workbook
- `onedrive://me/drive/items/{item_id}` — signed-in user's OneDrive workbook

> **Note**: Path-based DSNs (e.g. `sharepoint://.../Shared Documents/path/to/file.xlsx`)
> are not implemented. Use the ID-based forms above.

## Quick Start

```python
from excel_dbapi import connect

conn = connect(
    "msgraph://drives/{drive_id}/items/{item_id}",
    engine="graph",
    credential=your_credential,
)
cursor = conn.cursor()
cursor.execute("SELECT * FROM Sheet1")
print(cursor.fetchall())
conn.close()
```

## Authentication

The Graph backend requires a bearer token provider. Accepted credential types:

| Type | Example |
|---|---|
| Static token string | `credential="eyJ0eX..."` |
| Zero-arg callable | `credential=lambda: get_token()` |
| `TokenProvider` protocol | Object with `.get_token()` method |
| Azure Identity credential | `credential=DefaultAzureCredential()` |

### Token Provider Strategies

- **Managed identity / workload identity** in cloud runtimes (preferred)
- **Client credentials** for service-to-service access
- **On-behalf-of / delegated flows** for user-scoped access

Minimum practical scope for workbook edits: `Files.ReadWrite.All`

### Credential Handling

- Store client secrets/cert references in a secret manager (Key Vault, AWS Secrets Manager, etc.)
- Rotate credentials regularly
- Grant least privilege by tenant/site/drive where feasible

## Read vs Write

The Graph backend is **read-only by default**. To enable writes:

```python
conn = connect(
    "msgraph://drives/{drive_id}/items/{item_id}",
    engine="graph",
    credential=your_credential,
    readonly=False,
)
```

Key write behavior:

- Writes are **immediate** — `persistChanges=true` is used for the session
- **No transactions**: `autocommit=False` raises `NotSupportedError`; `rollback()` is not available
- **Metadata sync**: best-effort. If metadata sync fails after a successful worksheet mutation, the workbook change is kept and a warning is logged

## Connection Configuration

```python
conn = connect(
    "msgraph://drives/{drive_id}/items/{item_id}",
    engine="graph",
    credential=your_credential,
    readonly=False,
    timeout=30.0,
    max_retries=4,
    backoff_factor=0.5,
)
```

| Option | Default | Description |
|---|---|---|
| `timeout` | `30.0` | HTTP request timeout in seconds |
| `max_retries` | `3` | Maximum retry attempts for safe methods |
| `backoff_factor` | `0.5` | Exponential backoff factor (`factor * 2^attempt`) |
| `conflict_strategy` | `"fail"` | `"fail"` (If-Match) or `"force"` (last-writer-wins) |

### Recommended Starting Points

- **Interactive workloads**: `timeout=10–20`, `max_retries=2–3`, `backoff_factor=0.25–0.5`
- **Batch workloads**: `timeout=30–60`, `max_retries=4–6`, `backoff_factor=0.5–1.0`

## Error Handling

### Authentication and Authorization

- **401**: raised as `InterfaceError` — refresh/reacquire token and retry
- **403**: raised as `InterfaceError` — validate app scopes, consent state, and tenant alignment

### eTag Conflicts (Concurrent Writes)

The backend supports optimistic concurrency via `conflict_strategy`:

- `"fail"` (default): sends `If-Match` on writes; raises `OperationalError` on 412 conflicts
- `"force"`: bypasses `If-Match`; last-writer-wins

Use `"fail"` in production unless data overwrite is explicitly acceptable.

### Rate Limiting and Transient Failures

Retryable status codes: `429` (rate limited), `503` (service unavailable), `504` (gateway timeout).

- Retries are automatic for safe methods (`GET`, `HEAD`, `OPTIONS`) only
- `Retry-After` header is honored (capped at 60 seconds)
- Non-idempotent writes are not automatically retried

### Session Lifecycle

Workbook sessions are managed automatically. Stale/expired sessions are reopened
transparently with a single retry (including for mutating methods).

## Production Checklist

- [ ] Azure AD app registration completed with required Graph scopes
- [ ] Admin consent granted and verified in target tenant(s)
- [ ] Token refresh/caching strategy implemented and load-tested
- [ ] Secrets moved to a managed secret store (no plaintext in repo)
- [ ] `timeout`, `max_retries`, and `backoff_factor` tuned for workload
- [ ] `conflict_strategy` chosen (`"fail"` recommended for integrity)
- [ ] Concurrency policy documented (single writer or conflict-managed)
- [ ] Observability in place (status/retry/conflict/session metrics)
- [ ] Alerting configured for auth and throttling regressions
- [ ] Rollback and incident response playbook validated

For detailed production operations guidance, see [Graph API Production Guide](GRAPH_PRODUCTION.md).

## Further Reading

- [Graph API Production Guide](GRAPH_PRODUCTION.md) — deployment patterns, monitoring, diagnostics
- [Engine Selection Guide](engines.md) — comparing all three backends
- [Usage Guide](USAGE.md) — configuration, advanced patterns
- [Engine Benchmarks](BENCHMARKS.md) — performance characteristics
