# Example outputs

These files are the output of one review round against the sample document.
Regenerate them with:

```bash
./examples/run_example.sh
```

The round starts from `examples/sample-incoming.docx`, which
`examples/make_fixture.py` builds. The fixture carries comments from three
reviewers across five pages, a reply inside a comment thread, and an orphaned
comment that no passage is anchored to.

| File | What it is |
|---|---|
| `comment-C-001.md` | One comment file from the ledger |
| `comments-index.md` | The ledger index |
| `change-plan.md` | The pre-apply plan, before anything was modified |
| `change-report.md` | The post-apply report, after the changes landed |
| `reply-C-001.md` | One email-ready reply file |
| `document-1-clean.pdf` | The clean copy, the only file sent to the department |
| `document-1-review.pdf` | The before/after review of every changed page |

The triage step in `examples/run_example.sh` shows the status, owner, and
`reply_author` for each comment, including one left for a colleague
(`C-002`, `needs_input`) and one declined (`C-005`, `rejected`).

## Why the PDFs differ slightly from a fresh export

The repository is published through an API that transfers file content as
text. A PDF normally carries binary bytes (embedded fonts, compressed
streams, image data), which cannot survive that transfer intact.

The sample PDFs are therefore converted to a byte-for-byte ASCII form by
`scripts/pdf_to_ascii.py`: the embedded font program is dropped and replaced
with a base-14 font, image data is re-encoded as ASCII85, and the
cross-reference table is rebuilt. Every reader accepts the result, and the
text and layout are unchanged. `examples/run_example.sh` runs this step
automatically.

The untouched originals are written to `.govt-review/output/` before the
conversion runs.
