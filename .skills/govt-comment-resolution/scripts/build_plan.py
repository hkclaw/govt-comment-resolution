#!/usr/bin/env python3
"""Render the pre-apply change plan and the post-apply change report.

Usage:
  build_plan.py --comments comments/ --out plans/change-plan.md \
                --mode plan [--manifest manifest.json] [--version 1]

  build_plan.py --comments comments/ --out plans/change-report.md \
                --mode report --applied working/document-v1.docx \
                --manifest manifest.json --version 1

plan   — what WILL change, for approval before anything is modified.
report — what DID change, produced after the changes were applied.
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

STATUS_ORDER = [
    "open", "needs_input", "drafted", "approved", "resolved", "reopened",
    "rejected", "wont_fix",
]


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


def load_comments(comments_dir):
    out = []
    for path in sorted(Path(comments_dir).glob("C-*.md")):
        text = path.read_text(encoding="utf-8")
        fields = parse_frontmatter(text)
        if not fields:
            continue
        fields["_response"] = section_body(text, "Our response")
        fields["_notes"] = section_body(text, "Notes")
        out.append(fields)
    return out


def status_sort_key(fields):
    status = (fields.get("status") or "open").strip()
    try:
        return STATUS_ORDER.index(status)
    except ValueError:
        return len(STATUS_ORDER)


def truncate(text, limit=90):
    flat = " ".join((text or "").split())
    if len(flat) <= limit:
        return flat
    return flat[:limit - 1] + "…"


def build_plan(comments, version, source_hash):
    total = len(comments)
    will_apply = [c for c in comments if (c.get("status") or "").strip() in ("approved", "resolved")]
    needs_input = [c for c in comments if (c.get("status") or "").strip() == "needs_input"]

    lines = [
        "# Change plan",
        "",
        f"Version: {version}  ",
        f"Source document SHA-256: `{source_hash}`",
        "",
        f"**{len(will_apply)} of {total} comment(s) will be applied to the document.**",
        "",
        "> This plan describes what is about to change. Nothing has been modified yet.",
        "> Review it and confirm before the changes are applied.",
        "",
        "## Summary",
        "",
        "| ID | Status | Owner | Author of comment | Comment |",
        "|---|---|---|---|---|",
    ]
    for fields in sorted(comments, key=status_sort_key):
        lines.append(
            f"| {fields.get('id')} | {fields.get('status') or 'open'} "
            f"| {fields.get('owner') or '—'} | {fields.get('author') or '—'} "
            f"| {truncate(fields.get('_response') or fields.get('author'), 60)} |"
        )

    lines += ["", "## Detail", ""]
    for fields in sorted(comments, key=status_sort_key):
        status = (fields.get("status") or "open").strip()
        lines += [
            f"### {fields.get('id')} — {status}",
            "",
            f"- **Comment by:** {fields.get('author') or '—'} "
            f"({fields.get('date') or 'no date'})",
            f"- **Owner:** {fields.get('owner') or '—'}",
            f"- **Reply author:** {fields.get('reply_author') or '— (not set)'}",
            f"- **Durable id:** `{fields.get('durable_id')}`"
            + ("" if fields.get("stable_id", "true") == "true" else " (synthetic — not durable)"),
        ]
        anchor = fields.get("anchor_text") or ""
        if anchor:
            lines.append("- **Quoted text:**")
            lines.append("")
            lines.append(f"  > {truncate(anchor, 160)}")
        if fields.get("_response"):
            lines += ["- **Reply to be sent:**", ""]
            lines.append(f"  {truncate(fields['_response'], 200)}")
        edit = None
        for line in (fields.get("_response", "") + "\n" + fields.get("_notes", "")).splitlines():
            if line.strip().startswith("[edit]"):
                edit = line.strip()[len("[edit]"):].strip()
        if edit:
            lines.append(f"- **Body edit (tracked change):** `{edit}`")
        if fields.get("_notes"):
            lines += ["- **Notes:**", ""]
            lines.append(f"  {truncate(fields['_notes'], 200)}")
        lines.append("")

    if needs_input:
        lines += [
            "## Waiting on others",
            "",
            "These comments are assigned to someone else and are **not** included in this round:",
            "",
        ]
        for fields in needs_input:
            lines.append(f"- **{fields.get('id')}** — owner {fields.get('owner') or 'unassigned'}")
        lines.append("")

    return "\n".join(lines)


def build_report(comments, version, manifest, applied_path, source_hash):
    applied_index = {
        c.get("id"): c for c in manifest.get("comments", []) if isinstance(c, dict)
    }
    applied_hash = D.sha256_file(applied_path) if Path(applied_path).exists() else None

    lines = [
        "# Change report",
        "",
        f"Version: {version}  ",
        f"Source document SHA-256: `{source_hash}`  ",
        f"Applied document: `{applied_path}`" + (f"  \nSHA-256: `{applied_hash}`" if applied_hash else ""),
        "",
        "> This report describes what was actually applied, after the changes were made.",
        "",
        "## What was applied",
        "",
        "| ID | Status | Reply author | Reply written | Body edit |",
        "|---|---|---|---|---|",
    ]

    applied_count = 0
    missing = []
    # Applied comments first — this report is about what was applied.
    ordered = sorted(
        comments,
        key=lambda c: (0 if applied_index.get(c.get("id"), {}).get("applied") else 1,
                       status_sort_key(c)),
    )
    for fields in ordered:
        cid = fields.get("id")
        entry = applied_index.get(cid, {})
        was_applied = bool(entry.get("applied"))
        status = (fields.get("status") or "open").strip()
        if was_applied:
            applied_count += 1
        edit = entry.get("edit")
        edit_text = "—"
        if edit:
            edit_text = "applied" if edit.get("applied") else "**NOT APPLIED**"
        # An orphaned comment has a reply file but no in-document thread reply;
        # saying "yes" here would overstate what was written into the document.
        if was_applied:
            if entry.get("in_document_reply", True):
                reply_text = "yes (document + file)"
            else:
                reply_text = "file only (not anchored in document)"
        else:
            reply_text = "no"
        lines.append(
            f"| {cid} | {status} "
            f"| {entry.get('reply_author_name') or entry.get('reply_author') or fields.get('reply_author') or '—'} "
            f"| {reply_text} | {edit_text} |"
        )
        if status in ("approved", "resolved") and not was_applied:
            missing.append(fields)

    lines += [
        "",
        f"**{applied_count} comment(s) applied.**",
        "",
    ]

    if missing:
        lines += [
            "## Planned but not applied",
            "",
            "These comments were approved but no reply was written. They still need attention:",
            "",
        ]
        for fields in missing:
            lines.append(f"- **{fields.get('id')}** — status {fields.get('status')}")
        lines.append("")

    not_applied_edits = [
        (fields.get("id"), entry)
        for fields in comments
        for entry in [applied_index.get(fields.get("id"), {})]
        if entry.get("edit") and not entry["edit"].get("applied")
    ]
    if not_applied_edits:
        lines += [
            "## Edits that did not land",
            "",
            "The planned text was not found in the document, so no tracked change was made:",
            "",
        ]
        for cid, entry in not_applied_edits:
            lines.append(f"- **{cid}** — searched for `{entry['edit'].get('find')}`")
        lines.append("")

    unresolved = [
        c for c in comments
        if (c.get("status") or "").strip() in ("open", "drafted", "needs_input")
    ]
    if unresolved:
        lines += [
            "## Still outstanding",
            "",
            "| ID | Status | Owner | Comment |",
            "|---|---|---|---|",
        ]
        for fields in sorted(unresolved, key=status_sort_key):
            lines.append(
                f"| {fields.get('id')} | {fields.get('status')} "
                f"| {fields.get('owner') or '—'} "
                f"| {truncate(fields.get('anchor_text'), 60)} |"
            )
        lines.append("")

    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--comments", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--mode", choices=["plan", "report"], required=True)
    parser.add_argument("--manifest")
    parser.add_argument("--applied", help="applied document (report mode)")
    parser.add_argument("--version", default="1")
    args = parser.parse_args()

    comments = load_comments(args.comments)
    manifest = {}
    if args.manifest and Path(args.manifest).exists():
        manifest = json.loads(Path(args.manifest).read_text(encoding="utf-8"))
    source_hash = manifest.get("source", {}).get("sha256", "unknown")

    if args.mode == "plan":
        body = build_plan(comments, args.version, source_hash)
    else:
        if not args.applied:
            parser.error("--applied is required in report mode")
        body = build_report(comments, args.version, manifest, args.applied, source_hash)

    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(body, encoding="utf-8")
    print(f"{args.mode} -> {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
