# Development Guide

## Getting Started

### 1. Clone the repository

```bash
git clone https://github.com/your-username/excel-dbapi.git
cd excel-dbapi
```

### 2. Create and activate virtual environment

```bash
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
```

### 3. Install dependencies

```bash
pip install --upgrade pip setuptools hatchling
pip install .[dev]
hatch develop
```

### 4. Using Makefile

| Command            | Description                                                  |
|--------------------|--------------------------------------------------------------|
| `make install`    | Install dependencies in development mode                     |
| `make format`     | Format code with black and isort                             |
| `make lint`       | Run code linting (ruff, mypy)                                |
| `make test`       | Run unit tests                                               |
| `make build`      | Build the package                                            |
| `make clean`      | Remove build artifacts                                       |
| `make release-patch` | Create a patch release                                    |
| `make release-minor` | Create a minor release                                    |
| `make release-major` | Create a major release                                    |

---

## 🚀 Release Automation

This project uses GitHub Actions for automated releases.

### Release commands

| Command                 | Description                                   |
|------------------------|-----------------------------------------------|
| `make release-patch`   | Create a patch release (e.g. `0.1.1 → 0.1.2`) |
| `make release-minor`   | Create a minor release (e.g. `0.1.2 → 0.2.0`) |
| `make release-major`   | Create a major release (e.g. `0.2.0 → 1.0.0`) |

### PyPI Authentication

Required secret:

| Secret Name      | Description                               |
|------------------|-------------------------------------------|
| `PYPI_API_TOKEN` | PyPI API token for publishing the package |

---

## Docker Development Environment

```bash
make docker-dev    # Start development container
make docker-stop   # Stop development container
```
