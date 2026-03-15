from pathlib import Path
import re

ROOT = Path(__file__).resolve().parents[1]


def main() -> int:
    version = (ROOT / "VERSION").read_text(encoding="utf-8").strip()
    pyproject_text = (ROOT / "pyproject.toml").read_text(encoding="utf-8")
    match = re.search(r'^version\s*=\s*"([^"]+)"', pyproject_text, flags=re.MULTILINE)
    pyproject_version = match.group(1) if match else None
    init_py = (ROOT / "src" / "excel_dbapi" / "__init__.py").read_text(encoding="utf-8")
    changelog = (ROOT / "CHANGELOG.md").read_text(encoding="utf-8")

    errors = []
    if pyproject_version != version:
        errors.append(f"pyproject version ({pyproject_version}) != VERSION ({version})")

    expected_loader = '(Path(__file__).resolve().parents[2] / "VERSION").read_text(encoding="utf-8").strip()'
    if expected_loader not in init_py:
        errors.append("__init__.py no longer derives __version__ from VERSION")

    if not re.search(rf'<a name="v{re.escape(version)}"', changelog):
        errors.append(f"CHANGELOG missing anchor for v{version}")

    if errors:
        print("Version surface check failed:")
        for error in errors:
            print(f"- {error}")
        return 1

    print(f"Version surfaces are aligned at {version}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
