# Agents Guide

This repository may be worked on by human contributors and coding agents. Use this guide to keep changes consistent.

## Scope
- Keep changes small and focused on the requested task.
- Avoid refactors unless explicitly requested.
- Prefer existing patterns and conventions in this repo.

## Directory Map
- `src/`: Package source code.
- `tests/`: Test suite.
- `docs/`: Product and developer documentation.
- `scripts/`: Release and maintenance scripts.

## Versions and Releases
- Current development track is the **0.x** series.
- `v2.0.0` was published in error and is marked as withdrawn in the changelog.
- Update `pyproject.toml` for version changes and keep `CHANGELOG.md` in sync.
- Use SemVer: MAJOR.MINOR.PATCH.
- Target for 1.0.0: feature-complete read/write support, stable API, and documented behavior.

## Release Process
- Update `CHANGELOG.md` first, then bump `pyproject.toml`.
- Use `make release-patch`, `make release-minor`, or `make release-major` to generate tags.
- Do not create or modify git tags manually unless explicitly requested.

## Documentation
- Keep `docs/PRD.md` aligned with `docs/ROADMAP.md`.
- Use ROADMAP as the source of truth for milestone ordering.
- Update `README.md` when user-facing features change.

## Testing
- Run tests when changes affect logic; note if tests are not run.
- Do not delete or disable failing tests.
- Use `make test` to run the full test suite (sets up PYTHONPATH and coverage).
- Use `make setup` before first test run to create the virtualenv and install dependencies.
- Preferred commands: `make test` for tests; `ruff`/`mypy` when type or lint changes are involved.

## Deprecation Policy
- Announce deprecations in `CHANGELOG.md`.
- Keep deprecated behavior for at least one minor version before removal.

## Commit Hygiene
- Do not create commits unless explicitly asked.

## PR Conventions
- Keep PRs focused and small.
- Summaries should explain the why, not just the what.
