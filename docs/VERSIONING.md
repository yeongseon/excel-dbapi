# Versioning Policy

Canonical version source: `VERSION`

The following surfaces must match this value:
- `pyproject.toml` (`project.version`)
- package runtime version (`excel_dbapi.__version__` reads `VERSION`)
- CHANGELOG release anchor (`<a name="vX.Y.Z">`)
- Git tag (`vX.Y.Z`) used by release workflow

Validate with:

```bash
python scripts/check_version_surfaces.py
```
