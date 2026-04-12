VENV_DIR := .venv
PYTHON := $(VENV_DIR)/bin/python
PIP := $(VENV_DIR)/bin/pip

.PHONY: bootstrap
bootstrap:
	@if [ ! -d "$(VENV_DIR)" ]; then \
		echo "Creating virtual environment..."; \
		python3 -m venv $(VENV_DIR); \
	fi
	@$(PIP) install --upgrade pip > /dev/null
	@$(PIP) install -e ".[dev,pandas]" > /dev/null
	@echo "Development environment ready."

.PHONY: install
install: bootstrap
	@if [ -n "$$CI" ]; then \
		echo "CI detected: skipping pre-commit hook installation"; \
	else \
		$(MAKE) precommit-install; \
	fi

.PHONY: format
format:
	@$(PYTHON) -m ruff format src/ tests/

.PHONY: lint
lint:
	@$(PYTHON) -m ruff check src/ tests/

.PHONY: typecheck
typecheck:
	@$(PYTHON) -m mypy --strict src/

.PHONY: test
test:
	@$(PYTHON) -m pytest tests/ -v

.PHONY: cov
cov:
	@$(PYTHON) -m pytest tests/ -v --cov=excel_dbapi --cov-report=xml --cov-report=term-missing --cov-report=html
	@echo "Open htmlcov/index.html to view coverage report."

.PHONY: check
check:
	@$(MAKE) lint
	@$(MAKE) typecheck
	@echo "Lint and type check passed."

.PHONY: check-all
check-all:
	@$(MAKE) check
	@$(MAKE) test
	@echo "All checks passed including tests."

.PHONY: precommit
precommit:
	@$(PYTHON) -m pre_commit run --all-files

.PHONY: precommit-install
precommit-install:
	@$(PYTHON) -m pre_commit install

.PHONY: build
build:
	@$(PYTHON) -m build

.PHONY: clean
clean:
	@rm -rf *.egg-info dist build __pycache__ .pytest_cache

.PHONY: clean-all
clean-all: clean
	@find . -type d -name "__pycache__" -exec rm -rf {} +
	@find . -type f \( -name "*.pyc" -o -name "*.pyo" \) -delete
	@rm -rf .mypy_cache .ruff_cache .pytest_cache .coverage coverage.xml htmlcov .DS_Store site

.PHONY: help
help:
	@echo "Available commands:" && \
	grep -E '^\.PHONY: ' Makefile | cut -d ':' -f2 | xargs -n1 echo "  - make"
