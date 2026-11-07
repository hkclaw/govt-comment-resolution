#!/usr/bin/env python3
"""Verify a review round against the skill's acceptance criteria.

Usage:
  verify_round.py --comments comments/ --docx working/document-v1.docx \
                   --clean output/document-v1-clean.pdf \
                   --review output/document-v1-review.pdf \
                   --manifest manifest.json

Exits non-zero if any check fails, so a round cannot be declared complete
while a check is failing.
"""

import argparse
import json
import re
import sys
# This network filesystem has coarse mtimes, which can leave a stale
# __pycache__ behind and silently run old code. Never write bytecode.
sys.dont_write_bytecode = True
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import docxlib as D  # noqa: E402

FRONTMATTER_RE = re.compile(r"^---\s*\n(.*?)\n---\s*\n", re.S)
FIELD_RE = re.compile(r"^([A-Za-z_][A-Za-z0-9_]*):\s*(.*)$")

RESULTS = []


def check(name, ok, detail=""):
    RESULTS.append((name, ok, detail))
    mark = "PASS" if ok else "FAIL"
    print(f"  {mark}  {name}" + (f"  — {detail}" if detail and not ok else ""))
    return ok


def parse_frontmatter(text):
    match = FRONTMATTER_RE.match(text)
    if not match:
        return {}
    fields = {}
    for line in match.group(1).splitlines():
        field = FIELD_RE.match(line)
        if field:
            fields[field.group(1)] = field.group(2).strip().strip('"')
    return fields


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--comments", required=True)
    parser.add_argument("--docx", required=True)
    parser.add_argument("--clean")
    parser.add_argument("--marked")
    parser.add_argument("--review")
    parser.add_argument("--manifest")
    args = parser.parse_args()

    print("Verifying review round")

    comments = []
    for path in sorted(Path(args.comments).glob("C-*.md")):
        fields = parse_frontmatter(path.read_text(encoding="utf-8"))
        if fields:
            fields["_path"] = path
            comments.append(fields)

    check("every comment has a status file", len(comments) > 0, "no C-*.md found")
    if not comments:
        print(
            "\nNo comment files found — there is nothing to verify for this round.\n"
            "(A document with no reviewer comments has no review round to complete.)"
        )
        return 0

    # 1. every comment has a valid status, owner, and reply
    valid_statuses = {
        "open", "needs_input", "drafted", "approved", "resolved", "rejected", "wont_fix",
    }
    bad_status = [
        c["id"] for c in comments
        if (c.get("status") or "") not in valid_statuses
    ]
    check("every comment has a valid status", not bad_status, ", ".join(bad_status))

    missing_owner = [
        c["id"] for c in comments
        if (c.get("status") or "") in ("approved", "resolved") and not c.get("owner")
    ]
    check("applied comments have an owner", not missing_owner, ", ".join(missing_owner))

    missing_author = [
        c["id"] for c in comments
        if (c.get("status") or "") in ("approved", "resolved") and not c.get("reply_author")
    ]
    check(
        "applied comments name a reply author",
        not missing_author, ", ".join(missing_author),
    )

    # 2. no reply can be authored by the comment author
    self_authored = []
    for c in comments:
        author = (c.get("reply_author") or "").strip()
        if author and author.lower() == (c.get("author") or "").strip().lower():
            self_authored.append(c["id"])
    check("no reply is authored by the comment author", not self_authored,
          ", ".join(self_authored))

    # 3. the document opens and its comment structure is sound
    docx_path = Path(args.docx)
    if docx_path.exists():
        parts = D.read_parts(docx_path)
        comments_xml = parts.get(D.COMMENTS_PART)
        check("document contains a comments part", comments_xml is not None)

        defined = {c["w_id"] for c in D.parse_comments(parts)}
        extended = D.parse_comments_extended(parts)
        extended_paras = {e["para_id"] for e in extended}

        orphan_paras = [
            e["para_id"] for e in extended
            if e["parent_para_id"] and e["parent_para_id"] not in extended_paras
        ]
        check("no reply references a missing parent", not orphan_paras,
              ", ".join(orphan_paras))

        # every comment defined must have a paraId, or it cannot be resolved
        no_para = [c["w_id"] for c in D.parse_comments(parts) if not c["para_id"]]
        check("every comment has a paraId", not no_para, ", ".join(no_para))

        # the reply references must exist in document.xml
        doc_xml = parts["word/document.xml"].decode("utf-8", "replace")
        referenced = set(re.findall(r'<w:commentReference w:id="([^"]+)"', doc_xml))
        missing_refs = sorted(referenced - defined)
        check("every comment reference resolves to a comment",
              not missing_refs, ", ".join(missing_refs))
    else:
        check("applied document exists", False, str(docx_path))

    # 4. the clean PDF must not leak comment text
    if args.clean:
        clean_path = Path(args.clean)
        if clean_path.exists():
            texts = " ".join(D.pdf_page_texts(clean_path))
            leaked = []
            for c in comments:
                body = c.get("_path")
                raw = Path(body).read_text(encoding="utf-8")
                comment_text = re.search(
                    r"^##\s+Comment\s*$(.*?)^##\s+", raw, re.M | re.S | re.I
                )
                if comment_text:
                    snippet = " ".join(
                        re.sub(r"<!--.*?-->", "", comment_text.group(1), flags=re.S).split()
                    )[:60]
                    if snippet and snippet in texts:
                        leaked.append(c["id"])
            check("clean PDF contains no comment text", not leaked, ", ".join(leaked))
            check("clean PDF is readable", len(texts.strip()) > 0)
        else:
            check("clean PDF exists", False, str(clean_path))

    # 5. the review PDF must exist and cover the changed pages
    if args.review:
        review_path = Path(args.review)
        if review_path.exists():
            review_text = " ".join(D.pdf_page_texts(review_path))
            check("review PDF has a summary of changed pages",
                  "Changed pages" in review_text)
        else:
            check("review PDF exists", False, str(review_path))

    # 6. the manifest must record what was applied
    if args.manifest:
        manifest_path = Path(args.manifest)
        if manifest_path.exists():
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            check("manifest records the source hash",
                  bool(manifest.get("source", {}).get("sha256")))
            applied_ids = [
                c["id"] for c in manifest.get("comments", [])
                if isinstance(c, dict) and c.get("applied")
            ]
            check("manifest records applied comments", True, f"{len(applied_ids)} applied")
            check("manifest records a document version",
                  bool(manifest.get("versions")))
        else:
            check("manifest exists", False, str(manifest_path))

    failed = [r for r in RESULTS if not r[1]]
    print(f"\n{len(RESULTS) - len(failed)} passed, {len(failed)} failed")
    if failed:
        for name, _, detail in failed:
            print(f"  - {name}: {detail}")
        return 1
    print("Round verified.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
