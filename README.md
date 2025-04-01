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

Alternatively, you can use the provided Makefile to automatically set up everything:

```bash
make setup
```

---

### 4. Run tests, linters, and formatters

You can use the provided `Makefile` for convenient commands:

```bash
make test       # Run all tests
make format     # Auto-format code (black + isort)
make lint       # Run static analysis (ruff + mypy)
make build      # Build the package
make clean      # Remove build artifacts
```

You can also use additional quality checks:

```bash
make changelog         # Generate changelog
make release-patch     # Release a patch version
make release-minor     # Release a minor version
make release-major     # Release a major version
make commit m="msg"    # Commit and push changes
```

---

### ✅ Pre-commit setup

We use `pre-commit` to ensure consistent code quality before every commit.

#### Install and activate hooks

```bash
make precommit
```

#### Run checks on all files manually (optional)

```bash
make lint
```

The following tools are automatically run before each commit:
- `black`: code formatter
- `isort`: import sorter
- `ruff`: code style checker
- `mypy`: type checker

---

### 📄 Documentation

The project documentation is located in the `docs/` folder.

You can create an initial documentation file by running:

```bash
make docs
```

The `docs/` folder is structured to include:
- User Guide
- API Reference
- Example Use Cases

---

### 🐳 Development Environment with Docker

You can run a lightweight development environment using Docker:

```bash
make docker-dev    # Start development container
make docker-stop   # Stop development container
```

The Docker image uses **Python 3.9** to ensure compatibility with the minimum supported version.

---

### 🧹 Clean build artifacts

```bash
make clean
```

---

### 💡 Notes

- Python 3.9+ is required.
- We use [Hatchling](https://hatch.pypa.io/latest/) as the build backend.
- Editable installs are handled with `hatch develop`.
- Development dependencies are managed via `[project.optional-dependencies].dev` in `pyproject.toml`.
- A `Makefile` is provided to automate common development tasks.
- A lightweight Docker development environment is available.

---
