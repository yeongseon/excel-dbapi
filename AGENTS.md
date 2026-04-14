# AGENTS.md — Agent Rules for excel-dbapi

## Naming Conventions

All titles (issues, PRs, commits) MUST follow the same convention:

```
type: concise description of what and why
```

### Allowed `type` prefixes

| Type | When to use |
|------|-------------|
| `feat` | New feature or capability |
| `fix` | Bug fix |
| `refactor` | Code restructuring with no behavior change |
| `test` | Adding or updating tests only |
| `docs` | Documentation-only changes |
| `chore` | Maintenance, CI, dependencies, tooling |
| `perf` | Performance improvement |

### Title rules

1. **Describe the actual change, not the process.**
   - ✅ `fix: wrap non-DB-API exceptions as OperationalError at connect time`
   - ✅ `feat: add ABS, ROUND, REPLACE scalar functions`
   - ❌ `fix: close oracle round 12 gaps`
   - ❌ `fix: address review feedback`
   - ❌ `fix: resolve round 6 edge cases`

2. **No priority tags or internal labels in titles.**
   - ❌ `[P1] feat: add CTE support`
   - ❌ `[Bug]: connection fails on empty file`
   - ✅ `feat: add CTE support (WITH clause)`
   - ✅ `fix: connection fails on empty xlsx file`

3. **Be specific enough that someone can understand the change from the title alone.**
   - ❌ `fix: various bug fixes`
   - ❌ `fix: address multiple issues`
   - ✅ `fix: reset cursor state before execute and validate closed connection`

4. **One logical change per commit/issue/PR.** If a commit touches multiple unrelated areas, split it.

### Issue titles

- Use the same `type: description` format.
- The title must describe the **problem or feature**, not the solution process.
- Examples:
  - `fix: connection.execute() autocommit does not persist changes`
  - `feat: add window functions and aggregate FILTER`
  - `test: add stress tests for Graph retries and large-sheet workloads`
  - `docs: reconcile docs with implemented feature set`

### PR titles

- Match the linked issue title or summarize the change in the same `type: description` format.
- If one PR closes multiple issues, use the primary change as the title.

### Commit messages

- Subject line: `type: description` (lowercase type, imperative mood, no period at end).
- Keep subject under 72 characters.
- Body (optional): explain **why**, not what. The diff shows what changed.

## Code Conventions

- Python 3.10+, strict mypy typing.
- No `as any`, `@ts-ignore`, or type suppression equivalents (`type: ignore` without specific error code).
- Follow existing patterns in the codebase — check 2-3 similar files before writing new code.
- Run `make check-all` and `make test` before considering work complete.
