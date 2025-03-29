.PHONY: install format lint type-check sec-check dead-code test build clean

.PHONY: install-dev install-prod

install-dev:
	pip install .[dev]
	hatch develop

install-prod:
	pip install .


format:
	black .
	isort .

lint:
	ruff check .
	mypy .

type-check:
	mypy .

sec-check:
	bandit -r src

dead-code:
	vulture src

test:
	pytest

build:
	python -m build

clean:
	rm -rf dist build *.egg-info

publish-test:
	twine upload --repository testpypi dist/*

publish:
	twine upload dist/*
