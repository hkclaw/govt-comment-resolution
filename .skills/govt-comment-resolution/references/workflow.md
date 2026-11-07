# Workflow reference: file formats, status vocabulary, multi-round behaviour

---

## Review root layout

```
.govt-review/
├── manifest.json          # version record + idempotency ledger
├── identity.json          # roster of people who may author replies
├── incoming.docx          # original as received; never modified
├── comments/
│   ├── index.md           # regenerated summary of all comments
│   └── C-001.md ...       # one file per comment
├── replies/
│   └── C-001-reply.md     # one email-ready reply per comment
├── working/
│   └── document-v1.docx   # version N output, tracked changes intact
├── plans/
│   ├── change-plan.md     # PRE-apply
│   └── change-report.md   # POST-apply
└── output/
    ├── document-v1-clean.pdf
    ├── document-v1-marked.pdf
    └── document-v1-review.pdf
```

---

## Comment file format

Front-matter is the machine-readable half; the body is the human half.

```markdown
---
id: C-001
durable_id: 10A0B0C01
stable_id: true
w_id: "1"
author: R. Alvarez (Planning Dept)
author_initials: RA
date: "2026-11-06T09:00:00Z"
anchor_text: "The regional allocation is projected to grow by 4.2%..."
status: needs_input
owner: alice
reply_author: alice
---

## Comment

Please justify the 4.2% growth figure in paragraph 1.

## Quoted text

> The regional allocation is projected to grow by 4.2%...

## Our response

<!-- Draft the reply here. One paragraph per comment. -->

## Notes

<!-- Anything the next reviewer needs. -->
```

`durable_id` is the key. `w_id` is only meaningful within one file version.

`stable_id: false` means the document had no `commentsIds.xml`, so a synthetic
id was derived. If the document is re-extracted, that synthetic id may change.

### Re-extraction preserves triage

Running `extract_comments.py` again on the same document refreshes the
machine-derived fields (author, date, quoted text, ids) but **keeps** the
human's work: status, owner, reply_author, the drafted response, and the
notes. Use `--force` to regenerate the files from scratch and discard the
triage.

### Edit requests

A body edit is requested with a marker line in the response or notes section:

```
[edit] old text -> new text
```

The edit is applied as a tracked change authored by `reply_author`. If the
old text is not found, the change is reported as planned-but-not-applied rather
than dropped silently.

---

## Status vocabulary

| Status | Meaning | Applies this round? |
|---|---|---|
| `open` | not yet assessed | no |
| `needs_input` | needs another person's answer first | no, owner assigned |
| `drafted` | reply written, not yet reviewed | no |
| `approved` | reply approved, ready to apply | **yes** |
| `resolved` | already applied | **yes** (skipped as already done) |
| `rejected` | comment declined, no change | no |
| `wont_fix` | deferred to a later round | no |

Only `approved` and `resolved` are applied to the document.

---

## Reply file format

```markdown
---
comment_id: C-001
durable_id: 10A0B0C01
to: planning.dept@example.gov
subject: Response to comment C-001 — Annual Planning Document
author: Jane Doe
author_org: Planning Agency
author_email: jane.doe@agency.gov
date: 2026-11-07
---

To: planning.dept@example.gov
Subject: Response to comment C-001 — Annual Planning Document

Dear Review Team,

Thank you for your comment on R. Alvarez (Planning Dept):

> Please justify the 4.2% growth figure in paragraph 1.

Your comment relates to the following passage:

> The regional allocation is projected to grow by 4.2%...

The 4.2% figure reflects the revised population projection published in March.

Yours sincerely,

Jane Doe
Planning Agency
jane.doe@agency.gov
```

The reply file is the reply of record. This is what goes to the department.
The same text and the same author also appear as an in-document thread reply,
generated from the same source, so the two cannot disagree.

---

## Identity roster

```json
{
  "default_author": "jane.doe",
  "authors": [
    { "id": "jane.doe", "name": "Jane Doe", "org": "Planning Agency", "email": "jane.doe@agency.gov" },
    { "id": "alice", "name": "Alice Chen", "org": "Planning Agency", "email": "alice.chen@agency.gov" }
  ]
}
```

Validation, applied to both the reply file and the in-document reply:

1. `reply_author` must be present.
2. `reply_author` must be in the roster.
3. `reply_author` must not be the comment's author.

A violation is a hard failure with no document written. This is what makes
"I am sure I am the one replying" structural rather than aspirational.

---

## Manifest

```json
{
  "round": 1,
  "source": { "file": "incoming.docx", "sha256": "…", "received": "…" },
  "versions": [
    { "version": 1, "document": "working/document-v1.docx", "sha256": "…",
      "applied": ["C-001", "C-003"] }
  ],
  "comments": [
    { "id": "C-001", "durable_id": "10A0B0C01", "stable_id": true,
      "status": "resolved", "owner": "jane.doe",
      "reply_author": "jane.doe", "reply_author_name": "Jane Doe",
      "applied": true, "in_document_reply": true,
      "reply_w_id": "9",
      "edit": { "find": "grow by 4.2%", "replace": "grow by 4.1%", "applied": true } }
  ],
  "outputs": [
    { "kind": "clean", "file": "output/document-v1-clean.pdf", "sha256": "…" }
  ]
}
```

`applied: true` is the idempotency ledger. A re-run skips that comment. The
`sha256` values are what let a later round prove which input a plan described.

`owner` and `reply_author` hold the roster **id**; `reply_author_name` holds
the display name. Keeping both means a later round can match on either form.

`in_document_reply` is `false` for an orphaned comment: it was answered by
reply file only, because it has no anchor to carry a threaded reply. The change
report distinguishes the two cases rather than claiming every applied comment
was also written into the document.

---

## Multi-round behaviour

A second review round works like this:

1. The department returns the document with new comments.
2. Extract into the same review root; the manifest's `source.sha256` records
   which file this round started from.
3. Comment identity carries forward through `durable_id`. A comment that was
   resolved in round 1 and reappears in round 2 is the same comment. Match on
   `durable_id`, never on `w:id`, because Word renumbers `w:id` on every save.
4. `apply_replies.py` continues from the existing `working/` document when one
   exists, so round 1's replies and tracked changes are preserved rather than
   discarded.

**Not yet specified:** automatic re-opening of resolved comments when a
department raises the same point again, and cross-round aggregation of a single
comment into several rounds' reply files. For now, treat each round's comment
set as its own ledger and match across rounds by `durable_id` by hand.

---

## PDF outputs

| File | Built from | Contains | Goes to department? |
|---|---|---|---|
| `clean` | revisions accepted, comments stripped | body text only | **yes** |
| `marked` | as-is | body text + visible revisions | no |
| `review` | before vs after, aligned | changed pages side by side | no |

The clean PDF is verified to contain none of the comment text. That check is
part of `verify_round.py`.

The review PDF locates changed pages by **aligning the page sequences**, not by
comparing page numbers, because an edit that changes text length shifts content
across page boundaries. Pages that merely shifted are excluded; only pages
whose content genuinely differs appear.
