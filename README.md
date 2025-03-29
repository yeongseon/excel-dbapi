# excel-dbapi

A DBAPI-like interface for Excel files with extensible engines.

---

## Development Setup

If you're contributing to this project or running it in development mode, follow the steps below.

### 1. Clone the repository

```bash
git clone https://github.com/your-username/exceldb.git
cd exceldb
```

### 2. Create and activate virtual environment

```bash
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
```

### 3. Install dependencies (with dev tools)

```bash
pip install --upgrade pip setuptools hatchling
pip install .[dev]
hatch develop
```

> 💡 This installs runtime and development dependencies (e.g. `pytest`, `black`, `mypy`, `ruff`, etc.)  
> and sets up the project in editable mode using Hatch.

---

### 4. Run tests, linters, and formatters

You can use the provided `Makefile` for convenient commands:

```bash
make test       # Run all tests
make format     # Auto-format code (black + isort)
make lint       # Run static analysis (ruff + mypy)
make sec-check  # Run security check (bandit)
make dead-code  # Detect unused code (vulture)
make build      # Build the package
make clean      # Remove build artifacts
```

---

### ✅ 5. Pre-commit setup

We use `pre-commit` to ensure consistent code quality before every commit.

#### Install and activate hooks

```bash
pre-commit install
```

#### Run checks on all files manually (first time)

```bash
pre-commit run --all-files
```

#### Pre-commit will automatically run the following tools:
- `black`: code formatter
- `isort`: import sorter
- `ruff`: code style checker
- `mypy`: type checker
- `bandit`: security analyzer
- `vulture`: unused code detector

---

### 🧹 6. Clean build artifacts

```bash
make clean
```

---

### 💡 Notes

- Python 3.9+ is required.
- We use [Hatchling](https://hatch.pypa.io/latest/) as the build backend.
- Editable installs are handled with `hatch develop`.
- Development dependencies are managed via `[project.optional-dependencies].dev` in `pyproject.toml`.

