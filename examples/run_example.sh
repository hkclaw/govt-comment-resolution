#!/usr/bin/env bash
#
# Run one full review round against the sample document.
#
#   ./examples/run_example.sh
#
# Writes every artifact into .govt-review/, and prints the path of each one.
# This is the smoke test CI runs, and the way to regenerate the samples in
# docs/examples/.

set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$REPO_ROOT"

S=.skills/govt-comment-resolution/scripts
R=.govt-review

rm -rf "$R"
mkdir -p "$R"

echo "== build the sample document =="
python3 examples/make_fixture.py "$R/incoming.docx"

echo
echo "== identity roster =="
cat > "$R/identity.json" <<'EOF'
{
  "default_author": "jane.doe",
  "authors": [
    { "id": "jane.doe", "name": "Jane Doe", "org": "Planning Agency", "email": "jane.doe@agency.gov" },
    { "id": "alice", "name": "Alice Chen", "org": "Planning Agency", "email": "alice.chen@agency.gov" }
  ]
}
EOF

echo
echo "== extract =="
python3 "$S/extract_comments.py" --docx "$R/incoming.docx" --out "$R/comments" \
    --manifest "$R/manifest.json" --identity "$R/identity.json"

echo
echo "== triage =="
python3 - <<'PY'
from pathlib import Path

comments = Path(".govt-review/comments")


def set_fields(cid, status, owner, author, response=None, edit=None):
    path = comments / f"{cid}.md"
    text = path.read_text(encoding="utf-8")
    text = text.replace("status: open", f"status: {status}")
    text = text.replace('owner: ""', f"owner: {owner}\nreply_author: {author}")
    if response:
        text = text.replace(
            "<!-- Draft the reply here. One paragraph per comment. -->", response)
    if edit:
        text = text.replace(
            "<!-- Anything the next reviewer needs. -->", f"[edit] {edit}")
    path.write_text(text, encoding="utf-8")


set_fields("C-001", "approved", "jane.doe", "jane.doe",
           "The figure reflects the revised population projection published in March.",
           "grow by 4.2% -> grow by 4.1%")
set_fields("C-002", "needs_input", "alice", "alice")
set_fields("C-003", "approved", "alice", "alice",
           "A retention period of five years has been added to the clause.")
set_fields("C-004", "open", "", "")
set_fields("C-005", "rejected", "jane.doe", "jane.doe",
           "No change is needed; the appendix was updated separately.")
set_fields("C-006", "drafted", "jane.doe", "jane.doe",
           "The baseline column will be added once the 2024 figures are confirmed.")
set_fields("C-007", "approved", "jane.doe", "jane.doe",
           "The document has been read through and corrected.")
PY

echo
echo "== change plan =="
python3 "$S/build_plan.py" --comments "$R/comments" --out "$R/plans/change-plan.md" \
    --mode plan --manifest "$R/manifest.json" --version 1

echo
echo "== apply =="
python3 "$S/apply_replies.py" --docx "$R/incoming.docx" --comments "$R/comments" \
    --identity "$R/identity.json" --out "$R/working/document-v1.docx" \
    --manifest "$R/manifest.json"

echo
echo "== reply files =="
python3 "$S/build_replies.py" --comments "$R/comments" --identity "$R/identity.json" \
    --out "$R/replies" --to "planning.dept@example.gov" --date "2026-11-07" \
    --doc-title "Annual Planning Document"

echo
echo "== change report =="
python3 "$S/build_plan.py" --comments "$R/comments" --out "$R/plans/change-report.md" \
    --mode report --applied "$R/working/document-v1.docx" \
    --manifest "$R/manifest.json" --version 1

echo
echo "== PDFs =="
python3 "$S/render_pdfs.py" --docx "$R/working/document-v1.docx" --out "$R/output" --version 1
python3 "$S/compare_pages.py" --before "$R/incoming.docx" \
    --after "$R/working/document-v1.docx" --out "$R/output/document-1-review.pdf"

# The publishing path used for this repository transfers file content as
# text, so any binary byte in a PDF would arrive corrupted. Convert the
# sample PDFs to a byte-for-byte ASCII form that every reader accepts.
echo
echo "== making the sample PDFs ASCII-safe =="
for pdf in "$R/output"/document-1-clean.pdf "$R/output"/document-1-review.pdf; do
    if [ -f "$pdf" ]; then
        python3 "$S/pdf_to_ascii.py" "$pdf" "$pdf.tmp" && mv "$pdf.tmp" "$pdf"
    fi
done

# Copy the published samples into the repo so a reviewer can see them.
mkdir -p docs/examples
cp "$R/output/document-1-clean.pdf" docs/examples/document-1-clean.pdf
cp "$R/output/document-1-review.pdf" docs/examples/document-1-review.pdf
cp "$R/comments/index.md" docs/examples/comments-index.md
cp "$R/comments/C-001.md" docs/examples/comment-C-001.md
cp "$R/plans/change-plan.md" docs/examples/change-plan.md
cp "$R/plans/change-report.md" docs/examples/change-report.md
cp "$R/replies/C-001-reply.md" docs/examples/reply-C-001.md

echo
echo "== verify =="
python3 "$S/verify_round.py" --comments "$R/comments" --docx "$R/working/document-v1.docx" \
    --clean "$R/output/document-1-clean.pdf" \
    --review "$R/output/document-1-review.pdf" \
    --manifest "$R/manifest.json"

echo
echo "Artifacts:"
find "$R" -type f | sort | sed "s|^|  |"
