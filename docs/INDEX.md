
# Documentation Index

- [Usage Guide](USAGE.md)
- [Development Guide](DEVELOPMENT.md)
- [Project Roadmap](ROADMAP.md)
- [10-Minute Quickstart](QUICKSTART_10_MIN.md)
- [Operations Notes](OPERATIONS.md)
- [Public Roadmap](PUBLIC_ROADMAP.md)
- [Versioning Policy](VERSIONING.md)

## Limitations

- `PandasEngine` rewrites workbooks and may drop formatting, charts, and formulas.
- `OpenpyxlEngine` loads with `data_only=True`, so formulas are evaluated to values when reading.
- `rollback()` is only available when `autocommit=False`.
