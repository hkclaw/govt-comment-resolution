#!/usr/bin/env python3
"""Build a before/after review PDF showing every changed page side by side.

Usage:
  compare_pages.py --before incoming.docx --after working/document-v1.docx \
                   --out output/document-v1-review.pdf

The after side is rendered with tracked changes visible, so the reviewer sees
the markup in context.

Changed pages are found by ALIGNING the two page sequences, not by comparing
page numbers: an edit that changes text length shifts content across page
boundaries, so "page 3 changed" does not mean "before page 3 -> after page 3".
Pages that only shifted are correctly reported as unchanged.
"""

import argparse
import difflib
import io
import shutil
import sys
# This network filesystem has coarse mtimes, which can leave a stale
# __pycache__ behind and silently run old code. Never write bytecode.
sys.dont_write_bytecode = True
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import docxlib as D  # noqa: E402

SIMILARITY_THRESHOLD = 0.55
DPI = 150


def normalize(text):
    return " ".join((text or "").split())


def similarity(a, b):
    """How likely two pages are to be the same page, before vs after.

    Uses the maximum bidirectional word overlap rather than a raw sequence
    ratio: a page that was substantially rewritten still shares most of its
    vocabulary, while two genuinely different pages share almost none. A
    sequence ratio falls apart on heavy edits (a page with a long insertion
    can score below 0.5 against its own original).
    """
    left = set((a or "").lower().split())
    right = set((b or "").lower().split())
    if not left and not right:
        return 1.0
    if not left or not right:
        return 0.0
    shared = len(left & right)
    return max(shared / len(left), shared / len(right))


def align_pages(before_texts, after_texts, threshold=SIMILARITY_THRESHOLD):
    """Monotonic alignment of two page sequences via dynamic programming.

    Returns a list of (before_idx, after_idx, similarity). Either index may be
    None, meaning the page was deleted or added.
    """
    nb, na = len(before_texts), len(after_texts)
    score = [[0.0] * (na + 1) for _ in range(nb + 1)]
    back = [[None] * (na + 1) for _ in range(nb + 1)]

    for i in range(1, nb + 1):
        for j in range(1, na + 1):
            sim = similarity(before_texts[i - 1], after_texts[j - 1])
            matched = sim if sim >= threshold else 0.0
            option_match = score[i - 1][j - 1] + matched
            option_delete = score[i - 1][j]
            option_add = score[i][j - 1]
            best = max(option_match, option_delete, option_add)
            score[i][j] = best
            if matched > 0 and best == option_match:
                back[i][j] = "M"
            elif best == option_delete:
                back[i][j] = "D"
            else:
                back[i][j] = "A"

    ops = []
    i, j = nb, na
    while i > 0 or j > 0:
        move = back[i][j] if i > 0 and j > 0 else None
        if move == "M":
            ops.append((i - 1, j - 1, similarity(before_texts[i - 1], after_texts[j - 1])))
            i -= 1
            j -= 1
        elif move == "D":
            ops.append((i - 1, None, 0.0))
            i -= 1
        elif move == "A":
            ops.append((None, j - 1, 0.0))
            j -= 1
        elif i > 0 and j > 0:
            ops.append((i - 1, j - 1, similarity(before_texts[i - 1], after_texts[j - 1])))
            i -= 1
            j -= 1
        elif i > 0:
            ops.append((i - 1, None, 0.0))
            i -= 1
        else:
            ops.append((None, j - 1, 0.0))
            j -= 1
    ops.reverse()
    return ops


def map_pages(source_texts, target_texts, threshold=SIMILARITY_THRESHOLD):
    """Map each source page index to the best matching target page index.

    Used to line up the accepted rendering of a document with its marked
    rendering, which can differ in pagination because accepting revisions
    removes text that the marked rendering still shows.
    """
    mapping = {}
    for src_idx, tgt_idx, _sim in align_pages(source_texts, target_texts, threshold):
        if src_idx is not None and tgt_idx is not None:
            mapping[src_idx] = tgt_idx
    return mapping


def find_changed(before_texts, after_texts):
    """Pages whose content differs, using the pairings from align_pages."""
    changed = []
    for bi, aj, sim in align_pages(before_texts, after_texts):
        if bi is not None and aj is not None:
            if normalize(before_texts[bi]) != normalize(after_texts[aj]):
                changed.append({
                    "before_page": bi + 1,
                    "after_page": aj + 1,
                    "kind": "modified",
                    "similarity": sim,
                })
        elif bi is not None:
            changed.append({
                "before_page": bi + 1, "after_page": None,
                "kind": "deleted", "similarity": 0.0,
            })
        else:
            changed.append({
                "before_page": None, "after_page": aj + 1,
                "kind": "added", "similarity": 0.0,
            })
    return changed


def word_diff(before_text, after_text):
    """Word-level diff summary of what changed on a page."""
    before_words = normalize(before_text).split()
    after_words = normalize(after_text).split()
    matcher = difflib.SequenceMatcher(None, before_words, after_words)
    chunks = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            continue
        old = " ".join(before_words[i1:i2])
        new = " ".join(after_words[j1:j2])
        if tag == "delete":
            chunks.append(f"- {old}")
        elif tag == "insert":
            chunks.append(f"+ {new}")
        else:
            chunks.append(f"- {old}")
            chunks.append(f"+ {new}")
    return chunks


def render_pair_sheet(before_png, after_png, title, subtitle, chunks):
    """Compose one landscape sheet: before on the left, after on the right."""
    from PIL import Image, ImageDraw, ImageFont
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen import canvas

    left = Image.open(before_png) if before_png else None
    right = Image.open(after_png) if after_png else None

    if left is None:
        left = Image.new("RGB", (800, 1131), "white")
    if right is None:
        right = Image.new("RGB", (800, 1131), "white")

    page_w, page_h = landscape(A4)
    margin = 28
    gap = 16
    header = 62
    panel_w = (page_w - margin * 2 - gap) / 2
    panel_h = page_h - margin * 2 - header - 74

    def fit(image):
        ratio = min(panel_w / image.width, panel_h / image.height)
        return image.resize(
            (max(1, int(image.width * ratio)), max(1, int(image.height * ratio))),
            Image.LANCZOS,
        )

    left_fit = fit(left)
    right_fit = fit(right)

    buf = io.BytesIO()
    pdf = canvas.Canvas(buf, pagesize=(page_w, page_h))
    pdf.setTitle(title)

    pdf.setFont("Helvetica-Bold", 15)
    pdf.drawString(margin, page_h - margin - 16, title)
    pdf.setFont("Helvetica", 9.5)
    pdf.drawString(margin, page_h - margin - 33, subtitle)

    top = page_h - margin - header
    for label, image, x in (
        ("BEFORE", left_fit, margin),
        ("AFTER", right_fit, margin + panel_w + gap),
    ):
        box_h = panel_h
        box_w = panel_w
        pdf.setStrokeColorRGB(0.72, 0.74, 0.78)
        pdf.setLineWidth(0.7)
        pdf.rect(x, top - box_h, box_w, box_h, stroke=1, fill=0)
        pdf.setFont("Helvetica-Bold", 10)
        pdf.setFillColorRGB(0.25, 0.27, 0.31)
        pdf.drawString(x + 7, top - 14, label)
        # Embed as PNG. Raw RGB pixels would be three bytes per pixel, which
        # makes the review PDF enormous.
        pdf.drawImage(
            ImageReader(io.BytesIO(_png_bytes(image))),
            x + (box_w - image.width) / 2,
            top - 22 - image.height,
            width=image.width,
            height=image.height,
        )

    y = top - panel_h - 20
    if chunks:
        pdf.setFont("Helvetica-Bold", 9.5)
        pdf.setFillColorRGB(0.15, 0.15, 0.18)
        pdf.drawString(margin, y, "Text changes on this page:")
        y -= 13
        pdf.setFont("Courier", 8.4)
        for chunk in chunks[:14]:
            if y < margin + 12:
                pdf.showPage()
                pdf.setFont("Courier", 8.4)
                y = page_h - margin - 14
            colour = (0.72, 0.18, 0.18) if chunk.startswith("-") else (0.13, 0.45, 0.20)
            pdf.setFillColorRGB(*colour)
            pdf.drawString(margin + 4, y, chunk[:150])
            y -= 11
    else:
        pdf.setFont("Helvetica-Oblique", 9.5)
        pdf.setFillColorRGB(0.4, 0.4, 0.45)
        pdf.drawString(margin, y, "Layout changed; wording is identical on this page.")
    pdf.showPage()
    pdf.save()
    return buf.getvalue()


def _png_bytes(image):
    out = io.BytesIO()
    image.save(out, format="PNG")
    out.seek(0)
    return out.getvalue()


def build_summary_page(changed, before_name, after_name):
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    import io

    buf = io.BytesIO()
    pdf = canvas.Canvas(buf, pagesize=A4)
    width, height = A4
    margin = 48

    pdf.setTitle("Changed pages review")
    pdf.setFont("Helvetica-Bold", 16)
    pdf.drawString(margin, height - margin - 10, "Changed pages — before / after")

    pdf.setFont("Helvetica", 9.5)
    y = height - margin - 34
    pdf.setFillColorRGB(0.25, 0.27, 0.31)
    pdf.drawString(margin, y, f"BEFORE: {before_name}")
    y -= 13
    pdf.drawString(margin, y, f"AFTER:  {after_name}")
    y -= 24

    if not changed:
        pdf.setFont("Helvetica-Oblique", 11)
        pdf.drawString(margin, y, "No pages changed.")
        pdf.save()
        return buf.getvalue()

    pdf.setFont("Helvetica-Bold", 10)
    pdf.drawString(margin, y, f"{len(changed)} changed page(s):")
    y -= 18

    pdf.setFont("Helvetica", 9.5)
    for entry in changed:
        if entry["kind"] == "modified":
            label = (f"Before page {entry['before_page']}  ->  "
                     f"After page {entry['after_page']}")
        elif entry["kind"] == "deleted":
            label = f"Before page {entry['before_page']}  ->  deleted"
        else:
            label = f"New page  ->  After page {entry['after_page']}"
        pdf.setFillColorRGB(0.12, 0.12, 0.15)
        pdf.drawString(margin + 4, y, label)
        y -= 14
        if y < margin + 30:
            pdf.showPage()
            pdf.setFont("Helvetica", 9.5)
            y = height - margin - 20

    y -= 8
    pdf.setFont("Helvetica-Oblique", 9)
    pdf.setFillColorRGB(0.4, 0.4, 0.45)
    pdf.drawString(margin, y, "Internal review artifact — not for onward transmission.")
    pdf.save()
    return buf.getvalue()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--before", required=True, help="original .docx")
    parser.add_argument("--after", required=True, help="revised .docx")
    parser.add_argument("--out", required=True, help="output review .pdf")
    args = parser.parse_args()

    before_docx = Path(args.before)
    after_docx = Path(args.after)
    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    work = Path(tempfile.mkdtemp(prefix="compare-"))
    try:
        before_pdf = work / "before.pdf"
        after_pdf = work / "after.pdf"
        D.soffice_to_pdf(before_docx, before_pdf)
        D.soffice_to_pdf(after_docx, after_pdf)

        before_texts = D.pdf_page_texts(before_pdf)
        marked_texts = D.pdf_page_texts(after_pdf)

        # Change detection must run against the ACCEPTED text of the after
        # document. The marked rendering contains BOTH the deleted and the
        # inserted text, so an edited page looks far less similar to its
        # original than it really is — with a similarity threshold in place,
        # edited pages stop pairing and get reported as deleted + added.
        #
        # The images still come from the marked rendering, so the reviewer
        # sees the tracked-change markup.
        accepted_parts = D.read_parts(after_docx)
        D.accept_all_revisions(accepted_parts)
        clean_docx = work / "after-clean.docx"
        D.write_parts(accepted_parts, clean_docx)
        clean_pdf = work / "after-clean.pdf"
        D.soffice_to_pdf(clean_docx, clean_pdf)
        accepted_texts = D.pdf_page_texts(clean_pdf)

        # The accepted and marked renderings can paginate differently, so
        # map accepted page numbers onto marked page numbers for the images.
        accepted_to_marked = map_pages(accepted_texts, marked_texts)

        changed = find_changed(before_texts, accepted_texts)
        print(f"Pages: before={len(before_texts)} after={len(accepted_texts)} "
              f"changed={len(changed)}")

        if not changed:
            summary = build_summary_page([], before_docx.name, after_docx.name)
            out_path.write_bytes(summary)
            print(f"no changed pages -> {out_path}")
            return 0

        before_pngs = D.pdf_page_images(before_pdf, work / "before_png", dpi=DPI)
        after_pngs = D.pdf_page_images(after_pdf, work / "after_png", dpi=DPI)

        from pypdf import PdfReader, PdfWriter

        writer = PdfWriter()
        summary_reader = PdfReader(io.BytesIO(
            build_summary_page(changed, before_docx.name, after_docx.name)
        ))
        for page in summary_reader.pages:
            writer.add_page(page)

        for entry in changed:
            before_png = None
            after_png = None
            if entry["before_page"]:
                idx = entry["before_page"] - 1
                if 0 <= idx < len(before_pngs):
                    before_png = before_pngs[idx]
            if entry["after_page"]:
                accepted_idx = entry["after_page"] - 1
                marked_idx = accepted_to_marked.get(accepted_idx, accepted_idx)
                if 0 <= marked_idx < len(after_pngs):
                    after_png = after_pngs[marked_idx]

            before_text = before_texts[entry["before_page"] - 1] if entry["before_page"] else ""
            after_text = ""
            if entry["after_page"]:
                after_idx = entry["after_page"] - 1
                if after_idx < len(accepted_texts):
                    after_text = accepted_texts[after_idx]
            chunks = word_diff(before_text, after_text)

            if entry["kind"] == "modified":
                title = f"Page {entry['before_page']} / {entry['after_page']} — modified"
            elif entry["kind"] == "deleted":
                title = f"Page {entry['before_page']} — deleted"
            else:
                title = f"New page {entry['after_page']} — added"

            subtitle = (
                "Left: before review round.  Right: after review round, "
                "with tracked changes shown."
            )
            sheet = render_pair_sheet(
                before_png, after_png, title, subtitle, chunks
            )
            for page in PdfReader(io.BytesIO(sheet)).pages:
                writer.add_page(page)

        with open(out_path, "wb") as handle:
            writer.write(handle)
        print(f"review -> {out_path} ({len(changed)} changed page(s))")
    finally:
        import shutil
        shutil.rmtree(work, ignore_errors=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
