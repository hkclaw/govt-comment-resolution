#!/usr/bin/env python3
"""Generate the email-ready reply file for each answered comment.

Usage:
  build_replies.py --comments comments/ --identity identity.json \
                   --out replies/ [--department "Planning Department"] \
                   [--to dept@example.gov] [--date 2026-11-07]

One file per comment that carries a reply. The reply file is the reply of
record — this is what gets sent to the department by email.

Authorship is validated here too: the author must exist in the roster and must
not be the comment's own author. The same validation gates the in-document
reply, so the file and the document can never disagree about who is replying.
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

REPLY_STATUSES = {"approved", "resolved", "drafted", "rejected"}


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
        fields[key] = value
    return fields


def section_body(text, heading):
    pattern = re.compile(r"^##\s+" + re.escape(heading) + r"\s*$", re.M | re.I)
    match = pattern.search(text)
    if not match:
        return ""
    rest = text[match.end():]
    next_heading = re.search(r"^##\s+", rest, re.M)
    if next_heading:
        rest = rest[:next_heading.start()]
    rest = re.sub(r"<!--.*?-->", "", rest, flags=re.S)
    return rest.strip()


def strip_edit_markers(text):
    """Remove [edit] marker lines from the emailed body.

    The marker is an instruction to the tool, not something the department
    should read.
    """
    kept = []
    for line in (text or "").splitlines():
        if line.strip().startswith("[edit]"):
            continue
        kept.append(line)
    return "\n".join(kept).strip()


def quote(text):
    lines = [line for line in (text or "").splitlines() if line.strip()]
    if not lines:
        return ""
    return "\n".join(f"> {line}" for line in lines)


def build_reply_file(fields, person, department, to_address, date_str, doc_title):
    cid = fields.get("id")
    comment_author = fields.get("author") or "the reviewer"
    response = fields.get("_response") or ""
    anchor = fields.get("anchor_text") or ""

    subject = f"Response to comment {cid}"
    if doc_title:
        subject += f" — {doc_title}"

    lines = [
        "---",
        f"comment_id: {cid}",
        f"durable_id: {fields.get('durable_id')}",
        f"to: {to_address}",
        f"subject: {subject}",
        f"author: {person['name']}",
        f"author_org: {person['org']}",
        f"author_email: {person['email']}",
        f"date: {date_str}",
        "---",
        "",
        f"To: {to_address}",
        f"Subject: {subject}",
        "",
        "Dear Review Team,",
        "",
        f"Thank you for your comment on {comment_author}:",
        "",
        quote(fields.get("_comment_text") or ""),
        "",
    ]
    if anchor:
        lines += [
            "Your comment relates to the following passage:",
            "",
            quote(anchor),
            "",
        ]
    if response:
        lines += [strip_edit_markers(response), ""]
    else:
        lines += [
            "We note your comment and will respond separately.",
            "",
        ]
    lines += [
        "Yours sincerely,",
        "",
        person["name"],
        person["org"],
        person["email"],
        "",
    ]
    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--comments", required=True)
    parser.add_argument("--identity", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--department", default="the reviewing department")
    parser.add_argument("--to", default="")
    parser.add_argument("--date", default="")
    parser.add_argument("--doc-title", default="")
    args = parser.parse_args()

    identity = json.loads(Path(args.identity).read_text(encoding="utf-8"))
    roster = {a["id"]: a for a in identity.get("authors", [])}
    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)

    date_str = args.date or "1970-01-01"
    written = []
    skipped = []

    for path in sorted(Path(args.comments).glob("C-*.md")):
        text = path.read_text(encoding="utf-8")
        fields = parse_frontmatter(text)
        if not fields:
            continue
        status = (fields.get("status") or "").strip()
        if status not in REPLY_STATUSES:
            skipped.append((fields.get("id"), f"status={status}"))
            continue

        author_id = (fields.get("reply_author") or "").strip()
        if not author_id:
            skipped.append((fields.get("id"), "no reply_author set"))
            continue
        if author_id not in roster:
            skipped.append((
                fields.get("id"),
                f"reply_author '{author_id}' is not in the identity roster",
            ))
            continue
        person = roster[author_id]
        if fields.get("author") and person["name"] == fields["author"]:
            skipped.append((
                fields.get("id"),
                f"reply_author is the comment author ({person['name']})",
            ))
            continue

        fields["_response"] = section_body(text, "Our response")
        fields["_comment_text"] = section_body(text, "Comment")

        body = build_reply_file(
            fields, person, args.department, args.to, date_str, args.doc_title
        )
        target = out_dir / f"{fields.get('id')}-reply.md"
        target.write_text(body, encoding="utf-8")
        written.append((fields.get("id"), person["name"]))

    print(f"Wrote {len(written)} reply file(s) -> {out_dir}")
    for cid, author in written:
        print(f"  {cid}: reply by {author}")
    if skipped:
        print(f"Skipped {len(skipped)} comment(s):")
        for cid, reason in skipped:
            print(f"  - {cid}: {reason}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
