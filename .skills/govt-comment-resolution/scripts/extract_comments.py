#!/usr/bin/env python3
"""Extract comments from a .docx into per-comment Markdown files plus an index.

Usage:
  extract_comments.py --docx incoming.docx --out comments/ [--manifest manifest.json]

Each comment becomes comments/C-NNN.md with a machine-readable front-matter
block and human-editable body sections. comments/index.md summarises them all.

The ledger is keyed on the durable comment id (commentsIds.xml), because Word
renumbers w:id on every save. Where no durable id exists, a deterministic
synthetic id is derived and the comment is flagged stable_id: false.
"""

import argparse
import hashlib
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
    """Return the body under '## heading', stripped of HTML comments."""
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



def synthetic_id(author, date, text, anchor):
    digest = hashlib.sha1(
        "|".join([author or "", date or "", text or "", anchor or ""]).encode("utf-8")
    ).hexdigest()
    return f"SYN-{digest[:12].upper()}"


def comment_number(w_id):
    try:
        return f"C-{int(w_id):03d}"
    except (TypeError, ValueError):
        return f"C-{re.sub(r'[^0-9A-Za-z]+', '', str(w_id)) or 'X'}"


def yaml_quote(value):
    if value is None:
        return '""'
    text = str(value)
    if text == "":
        return '""'
    if re.search(r'[:#\n"\'{}[\],&*?|<>=!%@`]', text) or text.strip() != text:
        return '"' + text.replace("\\", "\\\\").replace('"', '\\"') + '"'
    return text


def build_comment_md(number, comment, durable_id, stable_id, anchor_text, status, owner):
    lines = [
        "---",
        f"id: {number}",
        f"durable_id: {durable_id}",
        f"stable_id: {'true' if stable_id else 'false'}",
        f"w_id: {yaml_quote(comment['w_id'])}",
        f"author: {yaml_quote(comment['author'])}",
        f"author_initials: {yaml_quote(comment['initials'])}",
        f"date: {yaml_quote(comment['date'])}",
        f"anchor_text: {yaml_quote(anchor_text)}",
        f"status: {status}",
        f"owner: {yaml_quote(owner)}",
        "---",
        "",
        "## Comment",
        "",
        comment["text"].strip() or "_(no comment text)_",
        "",
        "## Quoted text",
        "",
    ]
    if anchor_text:
        for line in anchor_text.splitlines():
            lines.append(f"> {line}")
    else:
        lines.append("_(no anchored text — this comment is not anchored in the body)_")
    lines += [
        "",
        "## Our response",
        "",
        "<!-- Draft the reply here. One paragraph per comment. -->",
        "",
        "## Notes",
        "",
        "<!-- Anything the next reviewer needs. -->",
        "",
    ]
    return "\n".join(lines)


def build_index(entries, problems, source_path):
    total = len(entries)
    lines = [
        "# Comment index",
        "",
        f"Source: `{source_path}`",
        f"Comments: {total}",
        "",
        "| ID | Author | Date | Status | Owner | Durable ID | File |",
        "|---|---|---|---|---|---|---|",
    ]
    for entry in entries:
        lines.append(
            f"| {entry['id']} | {entry['author'] or '—'} | {entry['date'] or '—'} "
            f"| {entry['status']} | {entry['owner'] or '—'} "
            f"| `{entry['durable_id']}` | [{entry['id']}]({entry['id']}.md) |"
        )
    lines += ["", "## Anchored text", ""]
    for entry in entries:
        anchor = entry["anchor_text"].replace("\n", " ").strip()
        if len(anchor) > 160:
            anchor = anchor[:157] + "..."
        lines.append(f"- **{entry['id']}** — {anchor or '_(not anchored)_'}")
    if problems:
        lines += ["", "## Problems detected", ""]
        for problem in problems:
            kind = problem["kind"]
            if kind == "not_anchored":
                lines.append(
                    f"- comment `{problem['w_id']}` is defined but has no anchor in the "
                    "document body — confirm what it refers to before responding"
                )
            elif kind == "dangling_range":
                lines.append(
                    f"- a range marker points at comment `{problem['w_id']}`, which is "
                    "not defined — the document may be damaged"
                )
            elif kind == "unknown_parent":
                lines.append(
                    f"- a reply references unknown parent `{problem['parent']}`"
                )
    lines.append("")
    return "\n".join(lines)


def merge_existing(existing, comment, durable_id, stable_id, anchor_text,
                   default_status, default_owner):
    """Refresh machine-derived fields while keeping the human's triage.

    The status, owner, reply_author, response body, and notes are the
    human's work — a re-extraction must not discard them. Everything else
    (author, date, quoted text, ids) is refreshed from the document.
    """
    fields = parse_frontmatter(existing)

    def field(name, fallback=""):
        value = fields.get(name)
        return value if value not in (None, "") else fallback

    status = field("status", default_status)
    owner = field("owner", default_owner)
    reply_author = fields.get("reply_author", "")

    response = section_body(existing, "Our response")
    notes = section_body(existing, "Notes")
    if not response:
        response = "<!-- Draft the reply here. One paragraph per comment. -->"
    if not notes:
        notes = "<!-- Anything the next reviewer needs. -->"

    lines = [
        "---",
        f"id: {field('id')}",
        f"durable_id: {durable_id}",
        f"stable_id: {'true' if stable_id else 'false'}",
        f"w_id: {yaml_quote(comment['w_id'])}",
        f"author: {yaml_quote(comment['author'])}",
        f"author_initials: {yaml_quote(comment['initials'])}",
        f"date: {yaml_quote(comment['date'])}",
        f"anchor_text: {yaml_quote(anchor_text)}",
        f"status: {status}",
        f"owner: {yaml_quote(owner)}",
    ]
    if reply_author:
        lines.append(f"reply_author: {yaml_quote(reply_author)}")
    lines += [
        "---",
        "",
        "## Comment",
        "",
        comment["text"].strip() or "_(no comment text)_",
        "",
        "## Quoted text",
        "",
    ]
    if anchor_text:
        for line in anchor_text.splitlines():
            lines.append(f"> {line}")
    else:
        lines.append("_(no anchored text — this comment is not anchored in the body)_")
    lines += ["", "## Our response", "", response, "", "## Notes", "", notes, ""]
    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--docx", required=True, help="source .docx")
    parser.add_argument("--out", required=True, help="comments output directory")
    parser.add_argument("--manifest", help="optional manifest.json to update")
    parser.add_argument(
        "--force", action="store_true",
        help="overwrite existing comment files (otherwise human edits are kept)",
    )
    args = parser.parse_args()

    docx_path = Path(args.docx)
    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)

    parts = D.read_parts(docx_path)
    comments = D.parse_comments(parts)
    extended = D.parse_comments_extended(parts)
    durable_map = {d["para_id"]: d["durable_id"] for d in D.parse_comments_ids(parts)}
    anchors, problems = D.extract_anchors(parts)

    done_by_para = {e["para_id"]: e["done"] for e in extended}
    parent_by_para = {e["para_id"]: e["parent_para_id"] for e in extended}

    # replies are not top-level comments; skip them so each file is one
    # department comment
    reply_paras = {e["para_id"] for e in extended if e["parent_para_id"]}

    # A reply legitimately has no anchor of its own (it shares the parent's
    # range), so only true top-level comments can be "not anchored".
    reply_w_ids = {
        c["w_id"] for c in comments if c["para_id"] in reply_paras
    }
    problems = [
        p for p in problems
        if not (p["kind"] == "not_anchored" and p["w_id"] in reply_w_ids)
    ]

    entries = []
    sequence = 0
    for comment in comments:
        para_id = comment["para_id"]
        if para_id and para_id in reply_paras:
            continue
        # Number by position among top-level comments so the sequence is
        # contiguous and stable regardless of how many replies are interleaved.
        sequence += 1
        number = f"C-{sequence:03d}"

        anchor_text = anchors.get(comment["w_id"], "")
        durable_id = durable_map.get(para_id)
        stable = durable_id is not None
        if not stable:
            durable_id = synthetic_id(
                comment["author"], comment["date"], comment["text"], anchor_text
            )

        status = "open"
        owner = ""
        if para_id and done_by_para.get(para_id):
            status = "resolved"

        target = out_dir / f"{number}.md"
        if target.exists() and not args.force:
            # Keep the human's triage. Re-extracting the same document must not
            # wipe the status, owner, reply, and notes that were set by hand.
            existing = target.read_text(encoding="utf-8")
            preserved = merge_existing(existing, comment, durable_id, stable,
                                       anchor_text, status, owner)
            target.write_text(preserved, encoding="utf-8")
        else:
            body = build_comment_md(
                number, comment, durable_id, stable, anchor_text, status, owner
            )
            target.write_text(body, encoding="utf-8")

        entries.append({
            "id": number,
            "durable_id": durable_id,
            "stable_id": stable,
            "w_id": comment["w_id"],
            "author": comment["author"],
            "date": comment["date"],
            "anchor_text": anchor_text,
            "status": status,
            "owner": owner,
        })

    # replies whose parent is unknown
    for entry in extended:
        parent = entry["parent_para_id"]
        if parent and parent not in {e["para_id"] for e in extended}:
            problems.append({"kind": "unknown_parent", "parent": parent})

    index_body = build_index(entries, problems, docx_path.name)
    (out_dir / "index.md").write_text(index_body, encoding="utf-8")

    if args.manifest:
        manifest_path = Path(args.manifest)
        manifest = json.loads(manifest_path.read_text(encoding="utf-8")) if manifest_path.exists() else {}
        manifest.setdefault("source", {})
        manifest["source"].update({
            "file": str(docx_path),
            "sha256": D.sha256_file(docx_path),
        })
        manifest["comments"] = [
            {
                "id": e["id"],
                "durable_id": e["durable_id"],
                "stable_id": e["stable_id"],
                "status": e["status"],
                "owner": e["owner"],
                "applied": False,
            }
            for e in entries
        ]
        manifest_path.parent.mkdir(parents=True, exist_ok=True)
        manifest_path.write_text(
            json.dumps(manifest, indent=2, ensure_ascii=False) + "\n", encoding="utf-8"
        )

    print(f"Extracted {len(entries)} comment(s) -> {out_dir}")
    if problems:
        print(f"Problems detected: {len(problems)} (see index.md)")
        for problem in problems:
            print(f"  - {problem['kind']}: {problem}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
