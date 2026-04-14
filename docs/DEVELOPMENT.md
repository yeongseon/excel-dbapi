# Development Guide

## Getting Started

### 1. Clone the repository

```bash
git clone https://github.com/yeongseon/excel-dbapi.git
cd excel-dbapi
```

### 2. Set up development environment

```bash
make install
```

This creates a virtual environment (`.venv/`), installs all dependencies in editable mode, and sets up pre-commit hooks.
The `dev` extra includes `pre-commit` and `build`, so `make install` and
`make build` work in a fresh virtual environment.

Or manually:

```bash
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
pip install -e ".[dev,pandas]"
```

### 3. Verify setup

```bash
make check-all   # lint + typecheck + tests
```

## Makefile Commands

| Command | Description |
|---------|-------------|
| `make install` | Bootstrap venv and install dev dependencies |
| `make format` | Format code with ruff |
| `make lint` | Run ruff linting checks |
| `make typecheck` | Run mypy strict type checking |
| `make test` | Run test suite with pytest |
| `make cov` | Run tests with coverage report (HTML + terminal) |
| `make check` | Run lint + typecheck |
| `make check-all` | Run lint + typecheck + tests |
| `make build` | Build distribution packages (sdist + wheel) |
| `make clean` | Remove build artifacts |
| `make clean-all` | Deep clean (caches, coverage, mypy cache) |

## Code Quality

### Linting and Formatting

**ruff** handles both linting and formatting:

```bash
make format   # Auto-format
make lint     # Check lint rules
```

### Type Checking

**mypy** is configured in strict mode:

```bash
make typecheck
```

All code must pass `mypy --strict` before merging.

## Running Tests

```bash
# All tests
make test

# With coverage report
make cov

# Specific file
.venv/bin/python -m pytest tests/test_parser.py -v

# Specific test
.venv/bin/python -m pytest tests/test_parser.py::test_select_basic -v
```

Test coverage target: **95%+** (see latest CI report for current percentage).

## Project Structure

```
excel-dbapi/
├── src/excel_dbapi/
│   ├── __init__.py          # Package entry, version, DB-API globals
│   ├── connection.py        # ExcelConnection (PEP 249)
│   ├── cursor.py            # ExcelCursor (PEP 249)
│   ├── parser.py            # SQL parser (tokenizer + AST)
│   ├── executor.py          # Query executor
│   ├── exceptions.py        # PEP 249 exception hierarchy
│   ├── reflection.py        # Dialect reflection helpers
│   ├── sanitize.py          # Formula injection defense
│   ├── engines/             # Backend engine implementations
│   ├── openpyxl/            # Openpyxl-specific engine
│   └── py.typed             # PEP 561 marker
├── tests/                   # 1,235+ tests (coverage tracked in CI)
├── docs/
│   ├── SQL_SPEC.md          # Formal SQL grammar (EBNF)
│   ├── USAGE.md             # Usage guide
│   ├── QUICKSTART_10_MIN.md # 10-minute quickstart
│   ├── OPERATIONS.md        # Concurrency & engine notes
│   ├── DEVELOPMENT.md       # This file
│   └── ROADMAP.md           # Project roadmap
├── pyproject.toml           # Project metadata (hatchling)
├── Makefile                 # Development commands
├── README.md
├── CHANGELOG.md
├── CONTRIBUTING.md
└── LICENSE
```

## Release Process

excel-dbapi uses **GitHub Releases** with **Trusted Publisher (OIDC)** for PyPI publishing. No API token required.

### Steps

1. Update `CHANGELOG.md` with new version entries.
2. Bump version in `pyproject.toml` and `src/excel_dbapi/__init__.py`.
3. Commit and push:
   ```bash
   git add pyproject.toml src/excel_dbapi/__init__.py CHANGELOG.md
   git commit -m "chore: bump version to X.Y.Z"
   git push origin main
   ```
4. Create a **GitHub Release** (via the GitHub UI or `gh release create vX.Y.Z`).
5. The `publish-pypi.yml` workflow triggers automatically, builds, validates, and publishes to PyPI.

### Verify

- PyPI: https://pypi.org/project/excel-dbapi/
- GitHub Releases: https://github.com/yeongseon/excel-dbapi/releases

## Continuous Integration

The CI pipeline (`.github/workflows/ci.yml`) runs on every push and pull request:

1. **Linting**: ruff
2. **Type checking**: mypy (strict mode)
3. **Testing**: pytest on Python 3.10, 3.11, 3.12, 3.13
4. **Coverage**: Upload to Codecov

All checks must pass before merging.

## Contributing Guidelines

- Write tests for all new features and bug fixes
- Maintain or improve code coverage (target: **95%+**)
- Follow the existing code style (enforced by ruff)
- Add type hints for all functions (enforced by mypy strict mode)
- Update documentation for user-facing changes
- Keep commits atomic with clear messages (`feat:`, `fix:`, `docs:`, `chore:`)
