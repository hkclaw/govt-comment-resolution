# OOXML comment handling: rules and known upstream bugs

Everything here was established empirically against LibreOffice 7.4.7 and the
OpenXML SDK (net9.0). Each rule corresponds to a failure that was observed, not
to a specification reading alone.

Read this before changing `scripts/docxlib.py`.

---

## The four-file comment architecture

A Word comment spans four parts plus two references in `document.xml`:

| Part | Holds | Required |
|---|---|---|
| `word/comments.xml` | the comment bodies | yes |
| `word/commentsExtended.xml` | `w15:commentEx`, holding paraId, parent paraId, `done` | yes for resolve + threading |
| `word/commentsIds.xml` | `w16cid:commentId`, holding the durable id | yes for stable identity |
| `word/commentsExtensible.xml` | `w16cex:commentExtensible`, holding dateUtc | recommended |

`document.xml` carries the anchors:

```xml
<w:p>
  <w:commentRangeStart w:id="1"/>
  <w:r><w:t>Text under review</w:t></w:r>
  <w:commentRangeEnd w:id="1"/>
  <w:r><w:commentReference w:id="1"/></w:r>
</w:p>
```

The `w:id` on all three markers must match the `w:id` in `comments.xml`.

---

## Rule 1: a reply needs its own `commentReference`

**Failure observed:** a reply added to `comments.xml` with a correct
`w15:paraIdParent` but no `commentReference` in `document.xml` was silently
dropped by both LibreOffice and Word. The thread appeared to contain only the
original comment. No error, no warning.

The reply's `commentReference` goes inside the parent's range:

```xml
<w:p>
  <w:commentRangeStart w:id="1"/>
  <w:r><w:t>Text under review</w:t></w:r>
  <w:commentRangeEnd w:id="1"/>
  <w:r><w:commentReference w:id="1"/></w:r>
  <w:r><w:commentReference w:id="9"/></w:r>   <!-- the reply -->
</w:p>
```

`docxlib.add_reply()` handles this.

---

## Rule 2: threading keys on `w14:paraId`, never `w:id`

`w15:commentEx/@w15:paraId` matches the `w14:paraId` of the comment's first
paragraph in `comments.xml`. `w15:paraIdParent` points at the *parent's
paraId*.

```xml
<w15:commentEx w15:paraId="00000009" w15:paraIdParent="1A2B3C4D" w15:done="0"/>
```

**Upstream bug:** the `docx` skill's `CommentSynchronizer.AddReply` sets
`paraIdParent` from the numeric comment ID (`parentCommentId.ToString("X8")`),
producing `paraIdParent="00000001"`, which matches no paragraph. Replies come
out orphaned. It also never sets `w14:paraId` on the reply's own paragraph.

---

## Rule 3: `w15:done` lives in `commentsExtended.xml`

It is an attribute of `w15:commentEx`, and it must be set on the **target**
comment.

**Upstream bug:** `CommentSynchronizer.ResolveComment` looks in
`commentsExtensible.xml` for a `comment` element (that part holds
`commentExtensible` elements keyed by durableId. It has no `done` flag at
all), and it selects `FirstOrDefault` over *any* matching element rather than
the requested comment. In practice it resolves the wrong comment, or none.

---

## Rule 4: all four parts update together

A new comment needs an entry in `comments.xml`, `commentsExtended.xml`,
`commentsIds.xml`, and `commentsExtensible.xml`, plus the content-type
overrides and relationships. Missing `commentsIds.xml` means no durable id, so
the comment's identity cannot be tracked across rounds.

---

## Rule 5: content types

`commentsExtensible.xml` uses:

```
application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml
```

The `ms-word` variant (`application/vnd.ms-word.commentsExtensible+xml`)
appears in some guides and is **rejected by the OpenXML SDK**. The document
fails to open with "invalid part with an unexpected content type".

The other three:

```
application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml
application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml
application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml
```

---

## Rule 6: deletions use `w:delText`, never `w:t`

Inside `<w:del>`, text must be `<w:delText>`. Using `<w:t>` inside a deletion
is invalid and silently corrupts the file.

```xml
<w:del w:id="10" w:author="Jane Doe" w:date="2026-11-07T00:00:00Z">
  <w:r><w:delText xml:space="preserve">old text</w:delText></w:r>
</w:del>
```

---

## Rule 7: content types and relationships need a default namespace

`[Content_Types].xml` and `*.rels` must serialise as:

```xml
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
```

not `<ns0:Types xmlns:ns0="...">`. The prefixed form is rejected by OPC
readers and the document will not open at all.

`xml.etree.ElementTree` uses a prefix by default. `docxlib._serialize()` takes
a `default_ns` argument for exactly this reason. Always pass it for these two
parts.

---

## Rule 8: stripping comments must only remove the reference run

When removing `w:commentReference`, remove **the run that contains it**, not
every run in the paragraph.

**Bug observed during development:** the first implementation removed all runs
from the paragraph, which silently deleted real body text from the outgoing
PDF. The document still opened; the content was just gone. Only a text
comparison against the source caught it.

---

## Rule 9: tracked changes must preserve reading order

A replacement must produce:

```xml
<w:r><w:t>head</w:t></w:r>
<w:del ...><w:r><w:delText>old</w:delText></w:r></w:del>
<w:ins ...><w:r><w:t>new</w:t></w:r></w:ins>
<w:r><w:t>tail</w:t></w:r>
```

**Bug observed:** merging head and tail into one run placed *both* before the
deletion, so the sentence tail jumped ahead of the edit:

> "projected to  over the coming year.grow by 4.1%"

The document opened fine and the PDF looked plausible, but the sentence was
scrambled in the copy that would have gone to the department. `docxlib`
emits head, then deletion, then insertion, then tail.

---

## Rule 10: revision ids are unique and monotonic

Every `w:ins`, `w:del`, `w:rPrChange`, `w:pPrChange` needs a `w:id` that is
unique across the document. `docxlib._next_revision_id()` scans the document
for the current maximum and continues from there.

---

## Rule 11: page pairing must compare accepted text

The before/after review PDF finds changed pages by aligning the two page
sequences. The comparison must run against the **accepted** text of the after
document (revisions applied), not the marked rendering.

**Failure observed during evaluation:** the marked rendering contains both the
deleted and the inserted text, so an edited page looks far less similar to its
original than it really is. With a similarity threshold in place, edited pages
stopped pairing and were reported as *deleted* plus *added*. The review PDF
showed blank panels where the before/after pair should have been, and reported
twice as many changed pages as actually existed.

The images still come from the **marked** rendering, so the reviewer sees the
tracked-change markup. The accepted and marked renderings can paginate
differently (accepting a deletion removes text the marked rendering still
shows), so accepted page numbers are mapped onto marked page numbers before
the images are selected.

---

## Rule 12: pair pages by shared vocabulary, not sequence ratio

Page similarity uses the maximum bidirectional word overlap:

```
shared / len(smaller)   vs   shared / len(larger), take the larger
```

A raw `difflib.SequenceMatcher` ratio falls apart on heavy edits. A page with
a long insertion scored 0.50 against its own original, below the pairing
threshold, so it was reported as deleted + added. Word overlap scored 0.909 for
the same pair, while two genuinely different pages scored 0.000.

Measured on representative pairs:

| Pair | Sequence ratio | Word overlap |
|---|---|---|
| heavy edit of the same page | 0.500 | **0.909** |
| slight edit of the same page | 0.991 | 0.952 |
| identical page, shifted | 1.000 | 1.000 |
| two different pages | 0.296 | **0.000** |

---

## Rule 13: never leave stale bytecode behind

This filesystem has coarse mtimes. A `__pycache__` entry can survive an edit
and silently run the old code. It cost real debugging time during development,
producing a fix that appeared not to work.

Every CLI script sets `sys.dont_write_bytecode = True` before importing
`docxlib`. Keep that line.

---

## Rendered verification

Structural validity is not enough. A document can satisfy every rule above
and still read wrong. Verify in a real renderer:

```bash
soffice --headless --convert-to fodt --outdir /tmp/check out.docx
```

Then check the flat ODF for the expected annotations and their state:

```xml
<office:annotation office:name="..." loext:resolved="true">
  <dc:creator>Jane Doe</dc:creator>
  <text:p>Our reply</text:p>
</office:annotation>
```

`loext:resolved="true"` confirms the resolve flag survived the round trip, and
`<dc:creator>` confirms the reply is attributed to the right person.

---

## Environment notes

- LibreOffice renders tracked changes as visible markup in PDF export:
  deletions struck through, insertions underlined and coloured by author.
- `soffice` writes its output next to the input; `docxlib.soffice_to_pdf()`
  moves it to the requested path.
- The upstream `docx` skill's `render_docx_pages.py` moves rendered PNGs from
  `/tmp` to the output directory, which fails across filesystems. Render to a
  temp dir on the same device as the output instead.
