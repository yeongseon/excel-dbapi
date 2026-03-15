---
name: Release checklist
about: Track release prep
labels: release
---

## Version

- [ ] VERSION updated
- [ ] pyproject.toml version updated
- [ ] CHANGELOG updated
- [ ] Tag planned (`vX.Y.Z`)

## Validation

- [ ] CI green
- [ ] Coverage artifact generated (`coverage.xml`)
- [ ] `python -m build` passes

## Publish

- [ ] Trusted Publishing environment configured
- [ ] Release notes drafted
