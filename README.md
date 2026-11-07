# Resolve government comments on a Word document

A Mavis skill that reads the review comments a government department left on a
Word document, tracks what needs answering, replies on your behalf, and produces
the PDFs you send back.

## What it does

One review round, end to end:

1. Reads every comment out of the `.docx`, with the text each one is anchored to.
2. Gives each comment a status and an owner, so a comment waiting on a colleague
   stays visible instead of disappearing.
3. Drafts a reply for each comment, and writes it as a file you can email.
4. Produces a change plan for review **before** anything is modified.
5. Applies the approved edits as tracked changes and marks the comments resolved.
6. Produces a change report for review **after** the changes are applied.
7. Emits three PDFs: a clean copy to send, a marked copy, and a before/after
   review of every changed page.

The reply is attributed to a named person from a roster you configure. A reply
with no author, an author outside the roster, or an author matching the
comment's own author stops the run before any document is written.

## Install

Copy the skill into your workspace:

```bash
git clone https://github.com/<owner>/<repo>.git
cp -r <repo>/.skills/govt-comment-resolution .skills/
```

The skill needs Python 3, LibreOffice, and poppler-utils on the path. The core
workflow uses the standard library alone. `requirements.txt` covers the
before/after PDF and the verification script.

## Use it

Ask for the work in your own words, for example "respond to the Planning
Department's comments on this document". The skill loads itself.

The full procedure lives in `.skills/govt-comment-resolution/SKILL.md`. It runs
these steps:

```bash
S=.skills/govt-comment-resolution/scripts
R=.govt-review

# 1. Extract the comments into a per-comment ledger
python3 $S/extract_comments.py --docx incoming.docx --out $R/comments \
    --manifest $R/manifest.json

# 2. Triage: set status, owner, and reply_author in each comments/C-*.md

# 3. Build the change plan and review it
python3 $S/build_plan.py --comments $R/comments --out $R/plans/change-plan.md \
    --mode plan --manifest $R/manifest.json --version 1

# 4. Apply
python3 $S/apply_replies.py --docx incoming.docx --comments $R/comments \
    --identity $R/identity.json --out $R/working/document-v1.docx \
    --manifest $R/manifest.json

# 5. Write the email-ready reply files
python3 $S/build_replies.py --comments $R/comments --identity $R/identity.json \
    --out $R/replies --to dept@example.gov --date 2026-11-07

# 6. Change report
python3 $S/build_plan.py --comments $R/comments --out $R/plans/change-report.md \
    --mode report --applied $R/working/document-v1.docx \
    --manifest $R/manifest.json --version 1

# 7. PDFs
python3 $S/render_pdfs.py --docx $R/working/document-v1.docx --out $R/output --version 1
python3 $S/compare_pages.py --before incoming.docx \
    --after $R/working/document-v1.docx --out $R/output/document-1-review.pdf

# 8. Verify
python3 $S/verify_round.py --comments $R/comments --docx $R/working/document-v1.docx \
    --clean $R/output/document-1-clean.pdf --review $R/output/document-1-review.pdf \
    --manifest $R/manifest.json
```

## The three PDFs

| File | Contains | Send to the department |
|---|---|---|
| `document-vN-clean.pdf` | Revisions accepted, all comments removed | Yes |
| `document-vN-marked.pdf` | Revisions visible | No |
| `document-vN-review.pdf` | Every changed page, before and after, side by side | No |

The review PDF lists the changed pages on a summary sheet, then gives each one a
landscape sheet with the before rendering on the left, the after rendering with
the tracked changes showing on the right, and a word-level diff underneath.

## Who may sign a reply

`identity.json` holds the roster:

```json
{
  "default_author": "jane.doe",
  "authors": [
    { "id": "jane.doe", "name": "Jane Doe", "org": "Planning Agency", "email": "jane.doe@agency.gov" }
  ]
}
```

`id` is the short handle used in comment files. `name` is what appears as the
reply author and as the tracked-change author.

## The ledger

One Markdown file per comment, plus `comments/index.md`. The front-matter is
what the tools read; the body is what you edit.

```markdown
---
id: C-001
durable_id: 10A0B0C01
w_id: "1"
author: R. Alvarez (Planning Dept)
status: needs_input
owner: alice
reply_author: alice
---

## Comment

Please justify the 4.2% growth figure.

## Our response

<!-- Draft the reply here. -->
```

To request a change to the document text, add a line to the response or notes:

```
[edit] old text -> new text
```

It is applied as a tracked change by the same person who signs the reply.

The ledger is keyed on `durable_id`, not on Word's `w:id`, because Word
renumbers `w:id` on every save.

### Status values

| Status | Meaning | Applied this round |
|---|---|---|
| `open` | Not assessed yet | No |
| `needs_input` | Waiting on a colleague | No |
| `drafted` | Reply written, not reviewed | No |
| `approved` | Ready to apply | Yes |
| `resolved` | Already applied | Skipped as done |
| `rejected` | Declined, no change | No |
| `wont_fix` | Deferred to a later round | No |

## Version control

`manifest.json` records the source hash, one entry per document version, and
the state of every comment. `applied: true` makes the apply step safe to
re-run: a second run skips the comment and continues from the existing output
instead of rebuilding from the original.

## Verify

```bash
./verify.sh
```

Runs the test suite and checks the skill structure. CI runs the same script,
then runs one full review round against a sample document so the pipeline is
exercised end to end.

```bash
./examples/run_example.sh
```

That builds a sample document with comments, runs the whole workflow over it,
and prints every artifact it produced. Sample outputs are in
`docs/examples/`.

### Continuous integration

The workflow file lives at `ci/verify.yml`. GitHub only runs workflows that
sit in `.github/workflows/`, so move it there to enable CI:

```bash
mkdir -p .github/workflows && cp ci/verify.yml .github/workflows/verify.yml
```

It is kept outside `.github/workflows/` because the API this repository is
published through refuses to write to that directory. The guard stops pushed
code from turning on continuous integration for itself. Copying the file is a
deliberate, reviewable step.

## Try it before you use it

The sample in `docs/examples/` shows one round end to end: the ledger entry, the
pre-apply plan, the post-apply report, an email-ready reply, the clean PDF, and
the before/after review PDF.

## Publish

`scripts/publish.sh` creates the GitHub repo and pushes `main`. It waits for
`gh auth login` to finish first, so you can start it before you authorise.

```bash
gh auth login --hostname github.com --git-protocol https --web
./scripts/publish.sh
```

Set `REPO_NAME`, `REPO_VISIBILITY`, `GIT_AUTHOR_NAME`, and
`GIT_AUTHOR_EMAIL` to override the defaults.

## Why this ships its own OOXML code

The `docx` skill's comment helpers are broken in ways that produce wrong
documents without any error. Each of these was confirmed against LibreOffice
and the OpenXML SDK:

- `ResolveComment` writes the resolve flag to the wrong part, and can mark the
  wrong comment resolved.
- `AddReply` keys reply threading on `w:id` instead of the comment paragraph's
  `w14:paraId`, so replies come out orphaned.
- A reply without its own `commentReference` is silently dropped by Word and
  LibreOffice. The thread looks like it holds only the original comment.

`.skills/govt-comment-resolution/references/ooxml-comments.md` records all
thirteen rules the implementation follows, each with the failure that
produced it.

## Not done

- Comments in a PDF are not read. The PDF is an output.
- Multi-round behaviour is partial. Comment identity carries across rounds
  through `durable_id`, but a resolved comment is not reopened automatically
  when the department raises the same point again.
- A colleague's answer is merged by hand into the comment file.

## Licence

MIT. See `LICENSE`.
