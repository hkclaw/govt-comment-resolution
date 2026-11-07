# Design: Government Comment Resolution Skill

**Date:** 2026-11-06
**Skill name:** `govt-comment-resolution`
**Status:** Approved for implementation

---

## 1. Purpose

A reusable skill for handling review comments that a government department has left
on a Word document, and for producing the reply materials that go back to them.

The workflow covers one review round end to end:

1. Read the comments out of the incoming `.docx`.
2. Give every comment a **resolve status** and a named owner.
3. Draft the reply to each comment.
4. Produce a **change plan** for review *before* anything is modified.
5. After approval, apply the edits as **tracked changes** and mark comments resolved.
6. Produce a **change report** for review *after* the changes are applied.
7. Emit the final deliverables: a clean PDF, a marked review PDF, and per-comment
   reply files suitable for sending by email.

### 1.1 Success criteria

A round is complete when:

- Every comment in the incoming document has exactly one status file, and every
  status file carries a valid status, an owner, and a reply.
- No comment is silently dropped. Comments that cannot be processed (orphaned
  anchors, duplicate IDs, unparseable parts) are reported, not ignored.
- The reply to every comment names a real human as its author. A reply with a
  missing or inherited author is a hard failure, not a warning.
- The pre-apply change plan is written to disk and reviewed before any edit is
  made to the document.
- The post-apply change report lists every edit that was actually applied, and
  any edit that was planned but not applied is called out.
- The clean PDF contains no comment text and no revision markup.
- The marked PDF contains the revisions and is never the file sent to the
  department.

---

## 2. Decisions (confirmed with the user)

| Decision | Choice | Consequence |
|---|---|---|
| Status storage | One Markdown file per comment, plus an index file | Colleague handoff = send one file; no concurrent-edit collisions |
| Reply storage | Separate files, one per comment, for sending by email | The **email-ready file** is the reply of record; authorship is a validated file field |
| Document standard | English, no special official-document standard | No GB/T 9704 geometry; standard professional formatting only |
| Versioning | Versioned folders + `manifest.json`, no git | Round history is directory-based and manifest-driven |
| Ledger shape | One file per comment + index file | Index regenerates from the per-comment files |
| Edit application | Tracked changes, awaiting acceptance | Department can accept/reject each change |
| Final output | Two PDFs: clean copy + marked review copy | Only the clean copy is sent onward |
| Reply authorship | Validated against `identity.json` roster | A reply with a missing or inherited author is a hard failure |

### 2.1 "Both before and after" — interpretation

The requirement was read as two distinct review artifacts:

- **Pre-apply change plan** (`change-plan.md`) — what *will* change, for approval
  before any modification. Contains no claims about outcomes.
- **Post-apply change report** (`change-report.md`) — what *did* change, produced
  after edits are applied, including any planned-but-not-applied divergence.

This interpretation was stated to the user and not contested.

---

## 3. Architecture

The skill is a **procedure plus deterministic scripts**. Judgment work (reading a
comment, drafting a reply, deciding whether a comment is actionable) stays in the
model. Mechanical work (parsing OOXML, writing replies, flipping resolve flags,
building PDFs) moves into scripts so it is repeatable and verifiable.

This split follows the principle that a script is for deterministic local
computation with no LLM judgment in the loop. Every mechanical step here meets
that bar.

### 3.1 Components

```
.govt-review/                     # working root for one review round
├── manifest.json                 # round metadata, file inventory, state
├── incoming.docx                 # original as received, never modified
├── comments/                     # one file per comment + index
│   ├── index.md
│   └── C-001.md ... C-0NN.md
├── replies/                      # one email-ready reply file per comment
│   └── C-001-reply.md ...
├── working/
│   ├── document-v1.docx          # after round 1 (tracked changes)
│   └── document-v2.docx          # after round 2
├── plans/
│   ├── change-plan.md            # PRE-apply, for review
│   └── change-report.md          # POST-apply, for review
└── output/
    ├── document-v1-clean.pdf     # send to department
    ├── document-v1-marked.pdf    # internal review only
    └── document-v1-review.pdf    # before/after changed-page review
```

### 3.2 Scripts

| Script | Responsibility | Judgment involved? |
|---|---|---|
| `scripts/extract_comments.py` | Parse `.docx` comment parts → per-comment Markdown + index | No |
| `scripts/apply_replies.py` | Write reply comments, set resolve flags, apply tracked edits | No |
| `scripts/build_replies.py` | Generate the email-ready reply file per comment | No |
| `scripts/build_plan.py` | Render pre-apply plan / post-apply report from comment files | No |
| `scripts/render_pdfs.py` | Produce clean + marked PDFs | No |
| `scripts/compare_pages.py` | Align page sequences, produce before/after changed-page PDF | No |
| `scripts/verify_round.py` | Assert the acceptance checks in §1.1 | No |

All scripts are Python 3 stdlib only (`zipfile`, `xml.etree`, `json`, `re`) plus
`subprocess` calls to `soffice` and `pdftoppm`. No third-party packages.

Every script sets `sys.dont_write_bytecode = True`: this filesystem has coarse
mtimes, and a surviving `__pycache__` entry silently runs stale code.

### 3.3 Regression tests

`tests/test_skill.py` — 55 checks covering the failure modes found during
development and evaluation. Each one corresponds to a defect that shipped a
silently wrong document at least once:

- a reply without a `commentReference` is dropped by Office consumers
- `ResolveComment` targets the wrong comment
- reply threading keyed on `w:id` instead of `w14:paraId`
- a tracked change that scrambles the sentence reading order
- stripping comments deletes real body text
- content types serialised with an `ns0:` prefix
- re-running apply discards previously applied replies
- change detection confused by tracked-change markup
- page pairing failing on heavily rewritten pages
- `[edit]` markers leaking into the emailed reply
- re-extraction wiping the human's triage
- the manifest recording one comment's owner on another

---

## 4. Data model

### 4.1 Comment identity — `durableId` is the key

Word renumbers `w:id` on every save. A ledger keyed on `w:id` breaks the moment
the file is opened and re-saved. `commentsIds.xml` carries `w16cid:durableId`,
which survives save cycles and copy/paste between documents.

**The ledger is keyed on `durableId`.** `w:id` is recorded alongside it for
cross-referencing within a single file version, but is never the key.

`durableId` is absent from documents that predate Word 2016 or were produced by
third-party tools. The extractor falls back to a deterministic synthetic ID
(`SYN-<sha1 of author+date+text+anchor, first 12 hex>`) and marks the comment
`stable_id: false`, so downstream stages know the identity is not durable and a
re-extraction could produce a different key.

### 4.2 Per-comment status file

`comments/C-001.md` — one file per comment, human-editable, machine-readable
front-matter block:

```markdown
---
id: C-001
durable_id: 1A2B3C4D
stable_id: true
w_id: "1"
author: DeptOfficer
author_initials: D
date: "2026-11-06T09:00:00Z"
anchor_text: "Second paragraph text"
status: needs_input
owner: alice
reply_author: Jane Doe
reply_author_org: Planning Agency
---

## Comment
Government comment A

## Quoted text
> Second paragraph text

## Our response
<!-- drafted reply body -->

## Notes
<!-- anything the next reviewer needs -->
```

Front-matter is the machine-readable half; the Markdown body is the human half.
A parser reads the front-matter; a human edits the body.

### 4.3 Status vocabulary

| Status | Meaning | Next action |
|---|---|---|
| `open` | Not yet assessed | Triage |
| `needs_input` | Requires a colleague's answer before we can respond | Assign owner, wait |
| `drafted` | Reply written, not yet reviewed | Review and approve reply |
| `approved` | Reply approved, ready to apply | Apply to document |
| `resolved` | Applied to the document and closed | Verify, then finalise |
| `reopened` | The department replied to our answer | Answer again |
| `rejected` | Comment declined; no change made | Justify in notes |
| `wont_fix` | Deferred to a later round | Record deferral reason |

`needs_input` is the state that supports the requirement that some comments must
be answered by someone else. It carries an `owner`, and the reply cannot be
drafted until that owner supplies the substance.

`reopened` is how a second review round reaches the ledger. The department's
reply to one of our answers carries no anchor of its own, so without this the
re-raise is invisible and the comment stays wrongly resolved. Extraction takes
`--identity` so replies we wrote are not mistaken for the department coming
back, and the department's words are recorded in the comment file.

### 4.4 Reply file

`replies/C-001-reply.md` — one per comment, standalone, email-ready:

```markdown
---
comment_id: C-001
durable_id: 1A2B3C4D
to: planning.dept@example.gov
subject: Response to comment C-001 — Planning Document Review
author: Jane Doe
author_org: Planning Agency
author_email: jane.doe@agency.gov
date: "2026-11-07"
---

To: planning.dept@example.gov
Subject: Response to comment C-001 — Planning Document Review

Dear Review Team,

Thank you for your comment:

> Government comment A

We appreciate your comment regarding the following text:

> Second paragraph text

Jane Doe
Planning Agency
```

**Authorship is a required, validated field.** The reply generator refuses to
emit a reply whose author is empty, whose author equals the comment's author
(a reply from the department to itself is always a bug), or whose author is
missing from the configured roster (§4.5). This is the mechanism that enforces
"I am sure I am the one replying" — it is structural, not advisory.

The same validation gates the in-document thread reply, so the file and the
document can never disagree about who is replying.

### 4.5 Identity configuration

`identity.json` at the review root holds the roster of people who may author
replies and the tracked changes applied to the document:

```json
{
  "default_author": "jane.doe",
  "authors": [
    { "id": "jane.doe", "name": "Jane Doe", "org": "Planning Agency", "email": "jane.doe@agency.gov" },
    { "id": "alice", "name": "Alice Chen", "org": "Planning Agency", "email": "alice.chen@agency.gov" }
  ]
}
```

Tracked changes are authored with the resolved author's display name, so the
department sees a real person in the revision marks.

### 4.6 Manifest

`manifest.json` is the round's version-control record. With no git, it is the
only thing that ties a document version to the comments and replies that
produced it.

```json
{
  "round": 1,
  "department": "Planning Department",
  "document_title": "Annual Planning Document",
  "created": "2026-11-07T00:00:00Z",
  "source": {
    "file": "incoming.docx",
    "sha256": "…",
    "received": "2026-11-06T09:00:00Z"
  },
  "versions": [
    {
      "version": 1,
      "document": "working/document-v1.docx",
      "sha256": "…",
      "created": "2026-11-07T00:10:00Z",
      "plan_sha256": "…",
      "report_sha256": "…"
    }
  ],
  "comments": [
    {
      "id": "C-001",
      "durable_id": "1A2B3C4D",
      "stable_id": true,
      "status": "resolved",
      "owner": "jane.doe",
      "reply_author": "jane.doe",
      "applied": true,
      "applied_in_version": 1,
      "reply_file": "replies/C-001-reply.md"
    }
  ],
  "outputs": [
    { "kind": "clean", "file": "output/document-v1-clean.pdf", "sha256": "…" },
    { "kind": "marked", "file": "output/document-v1-marked.pdf", "sha256": "…" }
  ]
}
```

The `applied` flag on each comment is the idempotency ledger (§5, Phase 4) and
the `sha256` values are what let a later round prove which input a given plan
described.

---

## 5. Workflow

### Phase 1 — Extract

Run `extract_comments.py` against `incoming.docx`.

Reads all four comment parts (`comments.xml`, `commentsExtended.xml`,
`commentsIds.xml`, `commentsExtensible.xml`), walks `document.xml` to capture
each comment's **anchored text** and its thread position, and emits one
Markdown file per comment plus `index.md`.

Extraction reports, and does not swallow:
- comments defined but not anchored in the body (`orphaned`)
- `commentRangeStart`/`End` IDs with no matching comment definition (`dangling`)
- comments whose reply chain references an unknown parent
- documents missing the extended comment parts entirely (older Word format) —
  these still extract, but resolve status cannot be written back, and the
  extractor says so explicitly

### Phase 2 — Triage

The model reads each `comments/C-*.md`, assigns a status and owner, and drafts
the reply body. `needs_input` comments are routed to a colleague, who edits that
one file and returns it.

### Phase 3 — Pre-apply plan

`build_plan.py` renders `plans/change-plan.md` from the comment files: a table
of every comment with its status, owner, quoted text, and the edit that will be
made. **This is a stop-and-review gate.** The model presents it and waits for
approval. No document is modified before approval is recorded.

The plan records the document SHA-256 so the report can prove which input the
plan described.

### Phase 4 — Apply

`apply_replies.py` applies approved work:

- Writes the reply as a **Word comment reply** on the target comment, with the
  author's identity set explicitly. The reply exists in two places by design:
  the **standalone file** (§4.4) is the reply of record and the artifact that
  gets emailed; the **in-document thread reply** keeps the comment thread
  coherent for the next review round, so a reader opening the document can see
  what was answered without opening the email. Both carry the same body and the
  same author, generated from one source, so they cannot diverge.
- Sets `w15:done="1"` on the target comment **and all of its replies**, so the
  thread resolves as a unit in Word.
- Applies body-text edits as **tracked changes** (`w:ins` / `w:del`) authored by
  the resolved author.

Revision IDs are allocated above the current document maximum and are unique
across the document. Deletions use `w:delText`, never `w:t` — using `w:t`
inside a deletion silently corrupts the file.

**Idempotency.** Each applied comment records its `durableId` in the manifest.
Re-running the apply step against an already-applied round is a no-op, so a
crash or a re-run cannot double-apply a reply or a tracked change.

### Phase 5 — Post-apply report

`build_plan.py` renders `plans/change-report.md`: what was planned, what was
applied, and any divergence — planned edits that were skipped, comments left
unresolved, replies whose author could not be resolved. The report states
explicitly when a planned edit did not land, rather than quietly omitting it.

### Phase 6 — Deliver

`render_pdfs.py` produces two PDFs from the same source:

- **Clean** (`document-vN-clean.pdf`) — all tracked changes accepted, all
  comment parts and anchors removed. Contains no comment text and no revision
  markup. This is the file sent to the department.
- **Marked** (`document-vN-marked.pdf`) — revisions preserved and visible, for
  internal review. Never sent onward.

### Phase 7 — Before/after review PDF

`compare_pages.py` produces a third artifact:

- **Review** (`document-vN-review.pdf`) — one landscape page per **changed**
  page, with the before rendering on the left and the after rendering on the
  right, labelled. A summary page up front lists every changed page, its kind
  (modified / added / deleted), and a word-level diff of that page's text.

This is the "review the changes once they are made" artifact. It is internal:
it is never sent to the department, because the before side contains the
original text of a document under review.

**Changed pages are found by aligning page sequences, not by index.** An edit
that changes text length shifts content across page boundaries, so "page 3
changed" does not mean "before page 3 → after page 3". The comparison:

1. Renders both documents to PDF and extracts per-page text.
2. Aligns the two page sequences with a monotonic dynamic-programming match
   using text similarity (a page is paired only if similarity clears a
   threshold).
3. Reports a pair as changed when the normalised text differs; an unmatched
   before page is *deleted*, an unmatched after page is *added*.

Pages that only *shifted* are correctly reported as unchanged — only pages whose
content genuinely differs appear in the review PDF.

The after side is rendered **with tracked changes visible**, so the reviewer
sees the insertion/deletion markup in context, not just a silently different
page.

`verify_round.py` then asserts the acceptance checks from §1.1, including that
the clean PDF contains none of the comment text, and that the review PDF covers
every changed page.

---

## 6. Correctness requirements for OOXML comment handling

These were established empirically against LibreOffice 7.4.7 and the OpenXML
SDK, and are the reason this skill ships its own implementation rather than
reusing the upstream `docx` skill's comment helpers.

1. **A reply needs its own `commentReference` in `document.xml`**, inside the
   parent's `commentRangeStart`/`commentRangeEnd`. Without it, Word and
   LibreOffice silently drop the reply — the comment thread looks like it
   contains only the original comment.
2. **Threading keys on `w14:paraId`, not `w:id`.** `commentsExtended.xml`
   carries `w15:commentEx` whose `w15:paraId` matches the comment paragraph's
   `w14:paraId`, and `w15:paraIdParent` points at the *parent's paraId*. The
   upstream implementation sets `paraIdParent` from the numeric comment ID,
   which produces orphaned replies.
3. **`w15:done` lives in `commentsExtended.xml`**, on the `w15:commentEx`
   element — not in `commentsExtensible.xml`. The upstream implementation looks
   in the wrong part and selects the first matching element rather than the
   target comment, so it can resolve the wrong comment.
4. **Every new comment needs all four parts updated** — `comments.xml`,
   `commentsExtended.xml`, `commentsIds.xml`, `commentsExtensible.xml` — plus
   the content-type overrides and relationships.
5. **`commentsExtensible.xml` uses the `application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml`**
   content type. The `ms-word` variant in some guides is rejected by the
   OpenXML SDK.
6. **Deleting content uses `w:delText`, never `w:t`.**
7. **Content types and relationship parts must serialise with a default
   namespace** (`xmlns="..."`), not an `ns0:` prefix. The prefixed form is
   rejected by OPC readers and the document will not open.
8. **When stripping comments for the clean copy**, remove only the run that
   carries the `commentReference`; runs containing real text must survive.
9. **Change detection compares accepted text, not marked text.** The marked
   rendering contains both the deleted and the inserted text, so an edited page
   looks far less similar to its original than it really is. Comparing against
   the accepted text keeps edited pages paired; the images still come from the
   marked rendering so the reviewer sees the markup.
10. **Pages are paired by shared vocabulary, not sequence ratio.** A page with
    a long insertion scored 0.50 against its own original under a sequence
    ratio — below the pairing threshold — so it was reported as deleted plus
    added. Maximum bidirectional word overlap scored 0.909 for the same pair,
    and 0.000 for two genuinely different pages.
11. **The manifest records `owner` and `reply_author` per comment.** A single
    stale loop variable once applied the last comment's owner to every entry.

---

## 7. Failure handling

| Failure | Behaviour |
|---|---|
| Comment parts missing entirely | Extract what is available; report that resolve status cannot be written back. Do not fabricate. |
| Orphaned comment (defined, not anchored) | Emit the comment file with `anchor_text: ""` and flag it in the index. A human decides. |
| Dangling range marker | Report in the index; do not create a comment file for a non-existent comment. |
| Reply author missing / matches comment author | **Hard fail.** No reply file is written. |
| Identity roster missing author | **Hard fail** with the list of known author IDs. |
| Tracked change target text not found | Record as planned-but-not-applied in the change report; never silently skip. |
| `soffice` or `pdftoppm` missing | Report the missing dependency; do not emit a partial PDF set. |
| Document fails to open after edits | Roll back to the pre-apply copy, report the failing step. |
| Re-run on an applied round | No-op via manifest record. |
| Rendering fails (cross-device temp move) | Write rendered pages to a temp dir on the same device as the output. |

---

## 8. Out of scope

- Reading comments from PDF documents (the PDF is an output, not an input).
- Automated merging of a colleague's reply file back into the ledger — a human
  copies the substance into the comment file. Merging is trivial but error-prone
  to automate, and the manual step keeps attribution unambiguous.
- Generating the department-facing covering letter beyond the per-comment
  replies.
- Any change to the upstream `docx` skill's own broken comment helpers. This
  skill is self-contained and does not modify a shared plugin.

---

## 9. Implementation notes

- The skill lives at `.skills/govt-comment-resolution/` and is picked up by the
  syncer automatically.
- `SKILL.md` carries the trigger description, the procedure, the output
  contract, and failure handling. Long OOXML detail goes in
  `references/ooxml-comments.md`; the status vocabulary and file formats go in
  `references/workflow.md`.
- Scripts are Python 3 stdlib only, so they run on any environment with Python 3
  and LibreOffice, with no install step.
- The skill is validated against a real prompt before delivery: a producer
  (skill loaded) versus a baseline (no skill), compared on completeness,
  adherence to the output contract, and absence of unnecessary work.
