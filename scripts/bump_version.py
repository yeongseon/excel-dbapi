import sys

import toml


def bump(version: str, bump_type: str) -> str:
    major, minor, patch = map(int, version.split("."))
    if bump_type == "major":
        major += 1
        minor = 0
        patch = 0
    elif bump_type == "minor":
        minor += 1
        patch = 0
    elif bump_type == "patch":
        patch += 1
    return f"{major}.{minor}.{patch}"


def main():
    if len(sys.argv) != 2 or sys.argv[1] not in ("major", "minor", "patch"):
        print("Usage: bump_version.py [major|minor|patch]")
        sys.exit(1)

    bump_type = sys.argv[1]
    pyproject = "pyproject.toml"
    data = toml.load(pyproject)

    # Changed: read from [project]
    current_version = data["project"]["version"]
    new_version = bump(current_version, bump_type)
    data["project"]["version"] = new_version

    with open(pyproject, "w") as f:
        toml.dump(data, f)

    print(f"Bumped version: {current_version} → {new_version}")


if __name__ == "__main__":
    main()