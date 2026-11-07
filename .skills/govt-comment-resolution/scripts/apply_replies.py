#!/usr/bin/env python3
"""Apply reviewed replies and tracked edits to a document.

Usage:
  apply_replies.py --docx incoming.docx --comments comments/ --identity identity.json \
                   --out working/document-v1.docx --manifest manifest.json \
                   [--dry-run]

Reads the per-comment Markdown files, and for each comment whose status is
`approved` or `resolved`:

  - writes a threaded reply authored by the resolved author
  - sets the resolve flag on the comment and its replies
  - applies the body edit named in the file as a tracked change

Authorship is validated: the author must exist in the identity roster and must
not be the comment's own author. A violation is a hard failure.
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

APPLY_STATUSES = {"approved", "resolved"}

FRONTMATTER_RE = re.compile(r"^---\s*\n(.*?)\n---\s*\n", re.S)
FIELD_RE = re.compile(r"^([A-Za-z_][A-Za-z0-9_]*):\s*(.*)$")


def parse_frontmatter(text):
    match = FRONTMATTER_RE.match(text)
    if not match:
        return {}
    fields = {}
    for line in match.group(1).splitlines():
        if not line.strip() or line.strip().startswith("#"):
            continue
        field = FIELD_RE.match(line)
        if not field:
            continue
        key, value = field.group(1), field.group(2).strip()
        if value.startswith('"') and value.endswith('"') and len(value) >= 2:
            value = value[1:-1].replace('\\"', '"').replace("\\\\", "\\")
        elif value.startswith("'") and value.endswith("'") and len(value) >= 2:
            value = value[1:-1]
        fields[key] = value
    return fields


def section_body(text, heading):
    """Return the body under '## heading', stripped of HTML comments."""
    pattern = re.compile(
        r"^##\s+" + re.escape(heading) + r"\s*$", re.M | re.I
    )
    match = pattern.search(text)
    if not match:
        return ""
    rest = text[match.end():]
    next_heading = re.search(r"^##\s+", rest, re.M)
    if next_heading:
        rest = rest[:next_heading.start()]
    rest = re.sub(r"<!--.*?-->", "", rest, flags=re.S)
    return rest.strip()


def load_comments(comments_dir):
    out = []
    for path in sorted(Path(comments_dir).glob("C-*.md")):
        text = path.read_text(encoding="utf-8")
        fields = parse_frontmatter(text)
        if not fields:
            continue
        fields["_path"] = str(path)
        fields["_response"] = section_body(text, "Our response")
        fields["_notes"] = section_body(text, "Notes")
        out.append(fields)
    return out


def resolve_author(fields, identity, comment_author):
    """Determine and validate the reply author. Raises on any ambiguity."""
    author_id = (fields.get("reply_author") or "").strip()
    if not author_id:
        raise ValueError(
            f"{fields['id']}: reply_author is empty — the reply must name a real "
            "person before it can be applied"
        )
    roster = {a["id"]: a for a in identity.get("authors", [])}
    if author_id not in roster:
        raise ValueError(
            f"{fields['id']}: reply_author '{author_id}' is not in the identity "
            f"roster (known: {', '.join(sorted(roster)) or 'none'})"
        )
    person = roster[author_id]
    if comment_author and person["name"] == comment_author:
        raise ValueError(
            f"{fields['id']}: reply_author '{person['name']}' is the same as the "
            "comment author — a reply must come from our side"
        )
    return person


def extract_edit(fields):
    """Parse the requested body edit from the comment file.

    Format in the 'Our response' or 'Notes' section:
        [edit] find -> replace
    """
    for section in (fields.get("_response", ""), fields.get("_notes", "")):
        for line in section.splitlines():
            line = line.strip()
            if line.startswith("[edit]"):
                payload = line[len("[edit]"):].strip()
                if "->" in payload:
                    find, replace = payload.split("->", 1)
                    return find.strip(), replace.strip()
    return None


def has_anchor(parts, w_id):
    """True when the comment has a range marker in the document body."""
    anchors, _ = D.extract_anchors(parts)
    return str(w_id) in anchors


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--docx", required=True)
    parser.add_argument("--comments", required=True)
    parser.add_argument("--identity", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--manifest", required=True)
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--date", default="1970-01-01T00:00:00Z")
    args = parser.parse_args()

    identity = json.loads(Path(args.identity).read_text(encoding="utf-8"))
    comments = load_comments(args.comments)
    if not comments:
        print("No comment files found; nothing to apply.")
        return 0

    manifest_path = Path(args.manifest)
    manifest = json.loads(manifest_path.read_text(encoding="utf-8")) if manifest_path.exists() else {}
    applied_index = {
        c.get("id"): c for c in manifest.get("comments", []) if isinstance(c, dict)
    }

    out_path = Path(args.out)

    # Continue from the existing output when there is one. Rebuilding from the
    # original source on a re-run would silently discard replies that were
    # applied in an earlier run, because those comments are skipped as
    # already-applied and so are never re-added.
    if out_path.exists():
        base_path = out_path
        base_note = "continuing from existing output"
    else:
        base_path = Path(args.docx)
        base_note = "starting from source document"

    parts = D.read_parts(base_path)
    D.ensure_comment_parts(parts)

    results = []
    failures = []
    warnings = []

    for fields in comments:
        cid = fields["id"]
        status = (fields.get("status") or "").strip()
        if status not in APPLY_STATUSES:
            results.append({"id": cid, "action": "skipped", "reason": f"status={status}"})
            continue

        prior = applied_index.get(cid, {})
        if prior.get("applied"):
            results.append({"id": cid, "action": "skipped", "reason": "already applied"})
            continue

        try:
            person = resolve_author(fields, identity, fields.get("author"))
        except ValueError as exc:
            failures.append({"id": cid, "error": str(exc)})
            continue
        author_id = (fields.get("reply_author") or "").strip()

        w_id = fields.get("w_id")
        # A comment that is not anchored in the body (an orphan) cannot carry
        # an in-document reply, but it still needs answering and still needs
        # its resolve flag set. The reply file is produced either way, so this
        # is a warning, not a failure.
        anchored = has_anchor(parts, w_id)
        reply = None
        if anchored:
            try:
                reply = D.add_reply(
                    parts, w_id, fields.get("_response") or "(no response drafted)",
                    person["name"], date=args.date,
                )
            except KeyError as exc:
                failures.append({"id": cid, "error": str(exc)})
                continue
        else:
            warnings.append({
                "id": cid,
                "warning": (
                    "comment is not anchored in the document body, so no in-document "
                    "reply could be written; the reply file was still produced and "
                    "the comment was still marked resolved"
                ),
            })

        resolved = D.resolve_comment(parts, w_id)

        edit = extract_edit(fields)
        edit_result = None
        if edit:
            find, replace = edit
            rev = D.apply_tracked_replace(parts, find, replace, person["name"], date=args.date)
            edit_result = {
                "find": find,
                "replace": replace,
                "applied": rev is not None,
                "revision_id": rev,
            }
            if rev is None:
                failures.append({
                    "id": cid,
                    "error": f"planned edit text not found in document: {find!r}",
                })

        results.append({
            "id": cid,
            "action": "applied",
            "author": person["name"],
            "owner": fields.get("owner") or "",
            "reply_author_id": author_id,
            "reply_w_id": reply["w_id"] if reply else None,
            "in_document_reply": reply is not None,
            "resolved_para_ids": resolved,
            "edit": edit_result,
        })

    if failures:
        print("FAILED — no document written:", file=sys.stderr)
        for failure in failures:
            print(f"  - {failure['id']}: {failure['error']}", file=sys.stderr)
        return 2

    if warnings:
        print("Warnings:")
        for warning in warnings:
            print(f"  - {warning['id']}: {warning['warning']}", file=sys.stderr)

    if args.dry_run:
        print(f"Dry run — validated ({base_note}). Planned actions:")
        for result in results:
            print(f"  {result['id']}: {result['action']} ({result.get('reason', '')})")
        return 0

    D.write_parts(parts, out_path)

    manifest.setdefault("comments", [])
    # Record the source hash if it is not already known. The extract step sets
    # it from the original incoming document; do not overwrite it on a later
    # round, where --docx may be the previous round's output.
    if not manifest.get("source", {}).get("sha256"):
        manifest["source"] = {
            "file": str(args.docx),
            "sha256": D.sha256_file(args.docx),
        }
    by_id = {c["id"]: c for c in manifest["comments"] if isinstance(c, dict)}
    for result in results:
        if result["action"] != "applied":
            continue
        entry = by_id.setdefault(result["id"], {"id": result["id"]})
        entry.update({
            "status": "resolved",
            "owner": result.get("owner") or entry.get("owner", ""),
            "reply_author": result.get("reply_author_id") or entry.get("reply_author", ""),
            "reply_author_name": result["author"],
            "applied": True,
            "applied_document": str(out_path),
            "in_document_reply": result["in_document_reply"],
        })
        if result.get("reply_w_id"):
            entry["reply_w_id"] = result["reply_w_id"]
        if result.get("edit"):
            entry["edit"] = result["edit"]
    manifest["comments"] = [by_id[c["id"]] for c in comments if c["id"] in by_id]
    manifest.setdefault("versions", []).append({
        "version": len(manifest.get("versions", [])) + 1,
        "document": str(out_path),
        "sha256": D.sha256_file(out_path),
        "source_sha256": D.sha256_file(args.docx),
        "applied": [r["id"] for r in results if r["action"] == "applied"],
    })
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    manifest_path.write_text(
        json.dumps(manifest, indent=2, ensure_ascii=False) + "\n", encoding="utf-8"
    )

    applied = [r for r in results if r["action"] == "applied"]
    print(f"Applied {len(applied)} of {len(comments)} comment(s) -> {out_path}")
    for result in applied:
        edit_note = ""
        if result.get("edit"):
            edit_note = (
                " [edit applied]" if result["edit"]["applied"] else " [EDIT NOT APPLIED]"
            )
        print(f"  {result['id']}: reply by {result['author']}{edit_note}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
