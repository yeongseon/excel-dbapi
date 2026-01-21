
# Documentation Index

- [Usage Guide](USAGE.md)
- [Development Guide](DEVELOPMENT.md)
- [Project Roadmap](ROADMAP.md)

## Limitations

- `PandasEngine` rewrites workbooks and may drop formatting, charts, and formulas.
- `OpenpyxlEngine` loads with `data_only=True`, so formulas are evaluated to values when reading.
- `rollback()` is only available when `autocommit=False`.
