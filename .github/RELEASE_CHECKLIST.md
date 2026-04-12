# Release Checklist

- [ ] Verify CI is green on `main`
- [ ] Run local quality gate with `make check-all`
- [ ] Build distribution artifacts with `make build`
- [ ] Update version in `pyproject.toml`
- [ ] Update `CHANGELOG.md` for the release
- [ ] Tag release `vX.Y.Z`
- [ ] Verify package publish succeeded
- [ ] Publish GitHub Release notes
