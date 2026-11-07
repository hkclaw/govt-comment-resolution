---
name: govt-comment-resolution
description: |
  Resolve review comments left by a government department on a Word document, and
  produce the reply materials that go back to them. Use when the user asks to
  "respond to the department's comments", "resolve the comments in this
  document", "reply to the review comments", "handle the comments from
  <department>", or has a .docx with comments from a reviewer and wants each
  one answered, tracked as a change plan, and returned as reply text plus a
  final PDF. Covers per-comment status tracking, replies that need someone
  else's input, pre-apply change plans, post-apply change reports, tracked
  changes, and clean/marked/review PDFs. Do NOT use for general Word editing
  with no comments involved (use the docx skill directly), for writing a new
  document from scratch, or for reading a PDF's comments (PDF is an output
  here, not an input).
---

# Government Comment Resolution

Run one review round over a Word document that a government department has
commented on: extract the comments, triage them, reply to them, apply the
changes as tracked revisions, and deliver the PDFs.

## When this applies

The input is a `.docx` containing reviewer comments (Word comments), and the
output is a set of replies plus a revised document. Judgment stays with you:
reading a comment, deciding whether it is actionable, drafting the answer.
Everything mechanical lives in `scripts/`.

## Inputs to collect

Before starting, confirm these. Ask only for what is genuinely missing; do not
interrogate the user for things you can infer.

1. **The incoming document.** The `.docx` with the department's comments.
2. **Who we are.** The roster of people who may author replies. Needed to
   enforce "I am the one replying". If the user gives a name, build
   `identity.json` from it. Never invent an author.
3. **The department.** Who the reply goes to, and their email address.
4. **The document title.** Used in reply subject lines.
5. **Which colleagues are answering which comments.** Only needed when some
   comments need someone else's input.

Ask for the identity roster first if it is missing: it is the one input you
cannot work around, and every reply depends on it.

## Setup

Create the review root and its identity file:

```bash
mkdir -p .govt-review/{comments,replies,working,plans,output}
cat > .govt-review/identity.json <<'EOF'
{
  "default_author": "jane.doe",
  "authors": [
    { "id": "jane.doe", "name": "Jane Doe", "org": "Planning Agency", "email": "jane.doe@agency.gov" }
  ]
}
EOF
```

`id` is the short handle used in comment files; `name` is what appears as the
reply author and as the tracked-change author.

## Procedure

Set `S=.skills/govt-comment-resolution/scripts` and `R=.govt-review`.

### 1. Extract

```bash
python3 $S/extract_comments.py --docx <incoming.docx> --out $R/comments \
    --manifest $R/manifest.json
```

Produces one Markdown file per comment plus `comments/index.md`. Each file
carries front-matter (id, durable_id, status, owner) and a body with the
comment, the quoted text, and empty sections for your response and notes.

Re-running this refreshes the machine-derived fields but **keeps** the human's
triage. Use `--force` to wipe it deliberately.

Read `index.md` and act on its **Problems detected** section before continuing.
Orphaned comments (defined but not anchored) need a human decision. A reply
can still be written for them, but no in-document reply is possible.

### 2. Triage

For each comment file, read it and set the front-matter:

- `status`. One of `open`, `needs_input`, `drafted`, `approved`, `resolved`,
  `rejected`, `wont_fix`.
- `owner`. Who is responsible.
- `reply_author`. Who signs the reply. Required before a reply is applied.

Then write the response under `## Our response`.

Use `needs_input` when a comment cannot be answered without someone else's
input. Assign the owner and leave the response empty; that comment is
deferred to the next round and is listed separately in the plan.

To request a body edit, add a line to the response or notes:

```
[edit] old text -> new text
```

It is applied as a tracked change authored by the same person as the reply.

### 3. Build the change plan, then stop

```bash
python3 $S/build_plan.py --comments $R/comments --out $R/plans/change-plan.md \
    --mode plan --manifest $R/manifest.json --version 1
```

**Present the plan and wait for approval.** Do not modify the document before
the user has confirmed it. This is the point of the plan. It is a review
gate, not a formality.

### 4. Apply

```bash
python3 $S/apply_replies.py --docx <incoming.docx> --comments $R/comments \
    --identity $R/identity.json --out $R/working/document-v1.docx \
    --manifest $R/manifest.json
```

This writes a threaded reply on each comment, sets the resolve flag, and
applies body edits as tracked changes. Use `--dry-run` first to validate.

Re-running is safe: already-applied comments are skipped and the existing
output is continued from, so a crash cannot double-apply a reply.

### 5. Generate the reply files

```bash
python3 $S/build_replies.py --comments $R/comments --identity $R/identity.json \
    --out $R/replies --to "<dept-email>" --date <YYYY-MM-DD> \
    --doc-title "<document title>"
```

One email-ready file per comment. These are what get emailed to the department.

### 6. Build the change report

```bash
python3 $S/build_plan.py --comments $R/comments --out $R/plans/change-report.md \
    --mode report --applied $R/working/document-v1.docx \
    --manifest $R/manifest.json --version 1
```

States what was actually applied, what was planned but did not land, and what
is still outstanding.

### 7. Render the PDFs

```bash
python3 $S/render_pdfs.py --docx $R/working/document-v1.docx --out $R/output --version 1
python3 $S/compare_pages.py --before <incoming.docx> \
    --after $R/working/document-v1.docx --out $R/output/document-1-review.pdf
```

Three PDFs:

| File | Contents | Sent to department? |
|---|---|---|
| `document-v1-clean.pdf` | revisions accepted, all comments removed | **Yes** |
| `document-v1-marked.pdf` | revisions visible | No |
| `document-v1-review.pdf` | before/after of every changed page | No |

The review PDF pairs each changed page side by side, with the tracked changes
shown on the after side and a word-level diff underneath. It is the artifact
for reviewing what changed once the changes are made.

### 8. Verify

```bash
python3 $S/verify_round.py --comments $R/comments --docx $R/working/document-v1.docx \
    --clean $R/output/document-v1-clean.pdf --review $R/output/document-1-review.pdf \
    --manifest $R/manifest.json
```

Exits non-zero if any check fails. A round is not complete while this fails.

## Output contract

Deliver, and state where each file is:

- `comments/`. One status file per comment plus `index.md`.
- `replies/`. One email-ready reply file per answered comment.
- `plans/change-plan.md`. The pre-apply plan, for review.
- `plans/change-report.md`. The post-apply report.
- `output/document-vN-clean.pdf`. **The only file sent to the department.**
- `output/document-vN-marked.pdf`. Internal review.
- `output/document-vN-review.pdf`. Before/after changed-page review.
- `working/document-vN.docx`. The document with tracked changes.
- `manifest.json`. The version record.

## Failure handling

- **Author missing or equal to the comment author.** Hard failure, no
  document written. This is the mechanism that guarantees we know who is
  replying. Fix the `reply_author` field rather than bypassing it.
- **Orphaned comment.** No in-document reply is possible. The reply file is
  still produced and the comment is still marked resolved; a warning is
  reported. Say so in your summary, because the department will not find a
  reply in the document for it.
- **Planned edit text not found.** No tracked change is made. It is reported
  in the change report as "planned but not applied". Do not silently drop it.
- **LibreOffice missing.** The PDF steps cannot run. Report the missing
  dependency; do not hand over a partial PDF set.
- **Document will not open after edits.** The apply step writes nothing on
  failure, so the previous document is intact. Investigate before retrying.
- **Many comments shift pages.** The review PDF handles this by aligning page
  sequences. Pages that merely shifted are correctly excluded.

## Notes

- The ledger is keyed on `durable_id`, not Word's `w:id`, because Word
  renumbers `w:id` on every save.
- `references/ooxml-comments.md`. The OOXML comment rules and the upstream
  bugs this skill works around. Read it before changing `scripts/docxlib.py`.
- `references/workflow.md`. File formats, status vocabulary, multi-round
  behaviour.
- `scripts/docxlib.py` is the only place that writes OOXML. Keep it that way.
- Regression tests: `python3 .skills/govt-comment-resolution/tests/test_skill.py`
