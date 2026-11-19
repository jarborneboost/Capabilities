"""The script installs its own libraries (inline `# /// script` header, read by `uv run`);
the tests install theirs from tests/pyproject.toml. Both lists must name the same libraries."""
import re
import tomllib
from pathlib import Path

REPO = Path(__file__).parents[3]
SCRIPT = REPO / "plugins" / "heistwright" / "skills" / "extract-requirements" / "scripts" / "annotate.py"
PYPROJECT = REPO / "tests" / "pyproject.toml"


def _script_dependencies() -> set[str]:
    header = re.search(r"^# /// script\n(.*?)^# ///$", SCRIPT.read_text(encoding="utf-8"), re.S | re.M)
    toml = "\n".join(line.removeprefix("#").removeprefix(" ") for line in header.group(1).splitlines())
    return set(tomllib.loads(toml)["dependencies"])


def test_script_and_test_dependencies_match():
    tested = set(tomllib.loads(PYPROJECT.read_text(encoding="utf-8"))["project"]["dependencies"])
    assert _script_dependencies() == {d for d in tested if not d.startswith("pytest")}
