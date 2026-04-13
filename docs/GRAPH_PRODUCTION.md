# Graph API Production Guide

This guide covers production deployment patterns for the `graph` engine in `excel-dbapi`.

## Authentication Setup

### 1) Azure AD app registration

1. Register an application in Microsoft Entra ID (Azure AD).
2. Grant Microsoft Graph application/delegated permissions required for workbook access.
3. For write scenarios, ensure write-capable scopes are granted and admin consented.

Minimum practical scope for workbook edits:

- `Files.ReadWrite.All`

The graph client surfaces a permission diagnostic on 403 errors and explicitly points to `Files.ReadWrite.All`.

### 2) Token provider strategies

`GraphClient` requires a bearer token provider (`TokenProvider`). Production patterns:

- **Managed identity / workload identity** in cloud runtimes (preferred)
- **Client credentials** for service-to-service access
- **On-behalf-of / delegated flows** for user-scoped access

Guidance:

- Cache tokens with their expiry and refresh before expiration.
- Keep refresh logic outside request hot paths when possible.
- Avoid embedding static long-lived secrets in environment variables for multi-tenant deployments.

### 3) Credential handling

- Store client secrets/cert references in a secret manager (Key Vault, AWS Secrets Manager, etc.).
- Rotate credentials regularly.
- Grant least privilege by tenant/site/drive where feasible.

## Connection Configuration

`GraphBackend` exposes network controls through `connect(..., engine="graph", ...)`:

- `timeout` (default `30.0` seconds)
- `max_retries` (default `3`)
- `backoff_factor` (default `0.5`)

These values flow to `GraphClient`, which:

- Retries only safe methods (`GET`, `HEAD`, `OPTIONS`)
- Retries status codes `429`, `503`, `504`
- Honors `Retry-After` when present (capped at 60 seconds)
- Uses exponential backoff (`backoff_factor * 2**attempt`)

Recommended starting points:

- **Interactive workloads**: `timeout=10-20`, `max_retries=2-3`, `backoff_factor=0.25-0.5`
- **Batch workloads**: `timeout=30-60`, `max_retries=4-6`, `backoff_factor=0.5-1.0`

Example:

```python
from excel_dbapi import connect

conn = connect(
    "msgraph://drives/<drive_id>/items/<item_id>",
    engine="graph",
    credential=my_token_provider,
    readonly=False,
    timeout=30.0,
    max_retries=4,
    backoff_factor=0.5,
)
```

## Error Handling

### Authentication and authorization

- **401** is raised as `InterfaceError` with guidance to re-authenticate.
- **403** is raised as `InterfaceError` with permission scope guidance.

Operational response:

- Refresh/reacquire token and retry.
- Validate app scopes and consent state.
- Validate tenant and resource audience alignment.

### eTag conflicts (concurrent writes)

`GraphBackend` supports optimistic concurrency via `conflict_strategy`:

- `fail` (default): sends `If-Match` on writes and raises `OperationalError` on 412 conflicts
- `force`: bypasses `If-Match` precondition and writes last-writer-wins

Use `fail` in production unless data overwrite is explicitly acceptable.

### Rate limiting and transient service failures

Graph retryable codes:

- `429` (rate limited)
- `503` (service unavailable)
- `504` (gateway timeout)

For non-idempotent writes, retries are intentionally not automatic. Implement caller-side recovery policies for writes when business logic allows retry.

### Session lifecycle failures

Workbook sessions are managed by `WorkbookSession` and reopened automatically on stale session errors. The backend retries once after session re-open, including mutating methods, to recover from invalid/expired session IDs.

## Performance Tuning

### Request shaping

- Prefer set-based SQL operations over row-by-row calls.
- Keep worksheet schemas stable to maximize targeted row patching.
- Minimize high-frequency workbook open/close cycles; reuse connections where possible.

### Write behavior

`GraphBackend` attempts optimized writes:

- Patches only changed row ranges when shape is compatible
- Uses row delete endpoints for removals when possible
- Falls back to full rewrite when change ratio is high

This reduces payload size and request count for sparse updates.

### Connection pooling and process model

- Reuse backend/client instances per worker where possible.
- If deploying many workers, bound worker concurrency to avoid Graph throttling.
- Tune `max_retries` and caller-level retries together to avoid retry storms.

### Batch planning

- Partition large update workloads by workbook/sheet.
- Prefer moderate batches (not one giant transaction).
- Stagger job start times for scheduled batch pipelines.

## Monitoring and Diagnostics

Capture and monitor:

- Response status distribution (2xx/4xx/5xx)
- Retry counts and cumulative retry wait time
- 401/403 auth failures by tenant/app
- 412 conflict rate (for concurrency contention)
- Session reopen events (stale session frequency)

Log recommendations:

- Correlate each operation with workbook identifiers (`drive_id`, `item_id`) and request IDs.
- Include exception class (`InterfaceError`, `OperationalError`) and status code.
- Redact access tokens and sensitive workbook metadata.

Alerting suggestions:

- Spike in 429 or sustained 503/504
- Elevated 401/403 after deployment/rotation
- Increased 412 conflict ratio for write-heavy workloads

## Production Deployment Checklist

- [ ] Azure AD app registration completed with required Graph scopes
- [ ] Admin consent granted and verified in target tenant(s)
- [ ] Token refresh/caching strategy implemented and load-tested
- [ ] Secrets moved to a managed secret store (no plaintext in repo)
- [ ] `timeout`, `max_retries`, and `backoff_factor` tuned for workload
- [ ] `conflict_strategy` chosen (`fail` recommended for integrity)
- [ ] Concurrency policy documented (single writer or conflict-managed)
- [ ] Observability in place (status/retry/conflict/session metrics)
- [ ] Alerting configured for auth and throttling regressions
- [ ] Rollback and incident response playbook validated
