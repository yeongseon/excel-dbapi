# excel-dbapi — Maturity & Growth Roadmap (2026)

## 0) Goal
- Education-ready in 6 months
- Production-hardened in 12 months
- Clear boundary: **DB-API style Excel driver**, not full SQL engine

---

## 1) 30-Day Execution Plan (compressed)

### Week 1 — Release trust signals
- [ ] Align single source of truth for version (`pyproject`, git tag, changelog, docs)
- [ ] Fill PyPI long description from README
- [ ] Switch publish pipeline to Trusted Publishing (OIDC)
- [ ] Add release checklist template

**Definition of Done**
- One canonical version visible in repo/tag/package metadata
- Release notes generated from changelog

### Week 2 — CI/Test hygiene
- [ ] Fix coverage artifact mismatch in CI
- [ ] Add OS smoke matrix (ubuntu/windows/macos)
- [ ] Add parser negative tests for documented SQL non-goals
- [ ] Add atomic-save and rollback integration tests

**Definition of Done**
- CI green on matrix
- Coverage upload stable

### Week 3 — Education quickstart
- [ ] Add `10-minute quickstart` tutorial
- [ ] Add comparison table (pandas/openpyxl/sqlite vs excel-dbapi)
- [ ] Add “what it is / what it is not” section in README
- [ ] Add minimal classroom examples under `examples/education/`

**Definition of Done**
- New user reaches first successful query in <10 minutes

### Week 4 — Security & operations
- [ ] Explicit parameter-binding guidance in all SQL examples
- [ ] Concurrency limitation docs (single-writer guidance)
- [ ] File integrity docs (atomic replace semantics)
- [ ] Add issue/PR templates for reproducible bug reports

**Definition of Done**
- Security/ops notes are explicit and test-backed

---

## 2) Top-10 Priority Actions (issue-ready)

1. **Unify version surfaces**
   - Labels: `release`, `high-priority`
   - Output: single canonical version policy
2. **Trusted Publishing migration**
   - Labels: `security`, `release`
3. **PyPI metadata completion**
   - Labels: `docs`, `release`
4. **Coverage pipeline fix**
   - Labels: `ci`, `test`
5. **OS matrix smoke CI**
   - Labels: `ci`
6. **SQL boundary tests (unsupported grammar)**
   - Labels: `test`, `parser`
7. **Atomic save / rollback integration tests**
   - Labels: `test`, `reliability`
8. **Education quickstart + examples pack**
   - Labels: `docs`, `education`
9. **Concurrency and engine tradeoff docs hardening**
   - Labels: `docs`, `ops`
10. **Community scaffolding (issue/pr templates + roadmap page)**
   - Labels: `community`, `docs`

---

## 3) v1.0 / v1.1 Release Checklists

### v1.0 (Education-ready)
- [ ] Install works on clean env (3.10~3.13)
- [ ] README quickstart complete
- [ ] SQL subset + non-goals explicit
- [ ] CI green + release automation reliable
- [ ] Changelog + tags + package metadata consistent

### v1.1 (Production-hardened)
- [ ] Large-file read-mode guidance/option documented and tested
- [ ] Reliability tests for rollback/atomic persistence expanded
- [ ] Optional dependency split (keep classroom install light)
- [ ] Security and ops runbook finalized

---

## 4) PRD/TDD/ARCH Sync Tasks
- [ ] PRD: Positioning and user segments updated (education first)
- [ ] TDD: Non-goal test cases and integrity tests reflected
- [ ] ARCH: Concurrency boundary + engine tradeoff section expanded
- [ ] README: mirrors PRD one-liner and non-goal boundaries

---

## 5) Suggested Milestone Names
- `M1: Release Hygiene`
- `M2: Test & CI Reliability`
- `M3: Education Experience`
- `M4: Production Guardrails`

---

## 6) Success Metrics
- Time-to-first-success: <10 min
- CI pass rate: >95%
- Install-related issue ratio: down by 50%
- Release cadence: at least monthly patch/minor
