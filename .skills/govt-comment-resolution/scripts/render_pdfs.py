#!/usr/bin/env python3
"""Render the deliverable PDFs for a review round.

Usage:
  render_pdfs.py --docx working/document-v1.docx --out output --version 1

Produces:
  document-vN-clean.pdf   tracked changes accepted, all comments removed
  document-vN-marked.pdf  revisions preserved and visible (internal only)
"""

import argparse
import shutil
import sys
# This network filesystem has coarse mtimes, which can leave a stale
# __pycache__ behind and silently run old code. Never write bytecode.
sys.dont_write_bytecode = True
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import docxlib as D  # noqa: E402


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--docx", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--version", required=True)
    args = parser.parse_args()

    docx_path = Path(args.docx)
    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    version = args.version

    # Marked copy: revisions intact, comments still present.
    marked = out_dir / f"document-{version}-marked.pdf"
    D.soffice_to_pdf(docx_path, marked)
    print(f"marked -> {marked}")

    # Clean copy: accept revisions, strip every comment.
    parts = D.read_parts(docx_path)
    D.accept_all_revisions(parts)
    D.strip_comments(parts)

    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        clean_docx = tmp_dir / "clean.docx"
        D.write_parts(parts, clean_docx)
        clean = out_dir / f"document-{version}-clean.pdf"
        D.soffice_to_pdf(clean_docx, clean)
        shutil.move(str(clean), str(clean))
    print(f"clean  -> {out_dir / f'document-{version}-clean.pdf'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
