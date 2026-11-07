#!/usr/bin/env bash
#
# Run every check for this skill.
#
#   ./verify.sh
#
# Exits non-zero if any check fails, so CI stops on the first broken thing.
# Requires: python3, libreoffice (soffice), poppler-utils (pdftoppm).

set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$REPO_ROOT"

echo "== environment =="
python3 --version
soffice --version
pdftoppm -v 2>&1 | head -1

echo
echo "== unit and integration tests =="
python3 .skills/govt-comment-resolution/tests/test_skill.py

echo
echo "== skill structure =="
python3 - <<'PY'
import re
import sys
from pathlib import Path

skill = Path(".skills/govt-comment-resolution")
text = (skill / "SKILL.md").read_text(encoding="utf-8")

match = re.match(r"^---\s*\n(.*?)\n---\s*\n", text, re.S)
if not match:
    sys.exit("SKILL.md has no frontmatter")
front = match.group(1)

name = re.search(r"^name:\s*(\S+)", front, re.M)
if not name:
    sys.exit("frontmatter has no name")
if name.group(1) != skill.name:
    sys.exit(f"frontmatter name {name.group(1)!r} does not match directory {skill.name!r}")
if not re.search(r"^description:", front, re.M):
    sys.exit("frontmatter has no description")

body_lines = len(text.splitlines())
if body_lines > 300:
    sys.exit(f"SKILL.md is {body_lines} lines, over the 300 line budget")

missing = []
for ref in sorted(set(re.findall(r"(?:references|scripts)/[\w.-]+", text))):
    if not (skill / ref).exists():
        missing.append(ref)
if missing:
    sys.exit(f"SKILL.md references files that do not exist: {', '.join(missing)}")

forbidden = [
    p for p in ("README.md", "CHANGELOG.md", "install.sh", ".env", ".env.example")
    if (skill / p).exists()
]
if forbidden:
    sys.exit(f"skill package contains forbidden files: {', '.join(forbidden)}")

print(f"SKILL.md ok: {body_lines} lines, {len(list(skill.rglob('*.py')))} scripts")
PY

echo
echo "== every tracked file survives the publishing path =="
python3 - <<'PY'
import subprocess
import sys
from pathlib import Path

tracked = subprocess.run(
    ["git", "ls-files"], capture_output=True, text=True
).stdout.split()

unpublishable = []
for name in tracked:
    path = Path(name)
    if not path.is_file():
        continue
    try:
        path.read_bytes().decode("utf-8")
    except UnicodeDecodeError as exc:
        unpublishable.append(f"{name} (byte {exc.start})")

if unpublishable:
    sys.exit(
        "these files carry bytes that the publishing path cannot transfer "
        "intact:\n  " + "\n  ".join(unpublishable)
    )
print(f"{len(tracked)} tracked files, all transferable")
PY

echo
echo "All checks passed."
