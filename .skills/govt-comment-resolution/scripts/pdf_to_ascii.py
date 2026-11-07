#!/usr/bin/env python3
"""Convert a PDF to a byte-for-byte ASCII-safe PDF.

The GitHub publishing path used by this skill transfers file content as text.
A PDF normally contains binary bytes (embedded font programs, compressed
streams), so an unmodified PDF cannot be published intact.

This script rewrites a PDF so every byte is ASCII:

1. qpdf expands the file into QDF form: object streams are expanded and
   streams are left uncompressed.
2. The embedded font program is removed and the font is re-pointed at a
   base-14 font, which every PDF reader has.
3. The cross-reference table is rebuilt, because removing an object shifts
   every offset after it.
4. The binary marker comment after the PDF header is replaced with spaces.
   It is advisory, and readers accept its absence.

The result opens in LibreOffice, Word, and any PDF reader, and survives the
text-only publishing path unchanged.

Usage:
  python3 scripts/pdf_to_ascii.py input.pdf output.pdf
"""

import base64
import re
import subprocess
import zlib
import sys
import tempfile
from pathlib import Path


def qdf_expand(pdf_path, out_path):
    """Expand a PDF into QDF form: no object streams, streams left compressed.

    Streams stay compressed. Inflating them would make the ASCII85 encoding
    of an embedded image roughly thirty times larger than the original.
    """
    result = subprocess.run(
        ["qpdf", "--qdf", "--object-streams=disable", "--deterministic-id",
         str(pdf_path), str(out_path)],
        capture_output=True, text=True, timeout=300,
    )
    if not Path(out_path).exists():
        raise RuntimeError(f"qpdf failed: {result.stderr.strip()}")
    return out_path


def strip_embedded_font(data):
    """Remove the embedded font program and re-point the font at base-14.

    Returns the rewritten bytes.
    """
    # The font descriptor references the font program as "/FontFile2 N 0 R".
    descriptor = re.search(
        rb"(\d+) 0 obj\s*<<(.*?/FontFile2\s+(\d+)\s+0\s+R.*?)>>",
        data, re.S,
    )
    if not descriptor:
        return data

    desc_num = descriptor.group(1).decode()
    program_num = descriptor.group(3).decode()

    # Drop the "/FontFile2 N 0 R" entry from the descriptor.
    data = re.sub(
        rb"(" + desc_num.encode() + rb" 0 obj\s*<<.*?)/FontFile2\s+\d+\s+0\s+R\s*",
        rb"\1", data, flags=re.S,
    )

    # Re-point the font object at a base-14 font and drop the descriptor link.
    for match in re.finditer(rb"\n(\d+) 0 obj\s*<<(.*?)>>\s*endobj", data, re.S):
        body = match.group(2)
        if b"/FontDescriptor" not in body or desc_num.encode() not in body:
            continue
        new_body = re.sub(rb"/BaseFont\s*/[\w+\-]+", b"/BaseFont /Helvetica", body)
        new_body = re.sub(
            rb"/FontDescriptor\s+" + desc_num.encode() + rb"\s+0\s+R", b"", new_body
        )
        new_body = re.sub(
            rb"/Encoding\s*/Identity-H", b"/Encoding /WinAnsiEncoding", new_body
        )
        data = data[: match.start(2)] + new_body + data[match.end(2):]
        break

    # Remove the font program object itself.
    data = re.sub(
        rb"\n" + program_num.encode() + rb" 0 obj\s*<<.*?>>\s*stream\r?\n"
        rb".*?\r?\nendstream\s*endobj",
        b"", data, flags=re.S,
    )
    return data


def strip_qdf_artifacts(data):
    """Remove the QDF-only comments."""
    return data.replace(b"%QDF: ignore_newline\n", b"")


def strip_binary_marker(data):
    """Replace the 4-byte binary marker after the PDF header with spaces."""
    index = data.find(b"\n%")
    if index < 0:
        return data
    end = data.find(b"\n", index + 1)
    if end < 0:
        return data
    return data[: index + 1] + b"    " + data[end:]


def resolve_indirect_lengths(data):
    """Turn QDF's indirect "/Length N 0 R" into a direct integer.

    QDF stores each stream length in its own object. A normal PDF writes the
    number inline. Resolving it keeps the output readable by strict parsers.
    """
    lengths = {}
    for match in re.finditer(rb"\n(\d+) 0 obj\s*\n?(\d+)\s*\n?endobj", data):
        lengths[match.group(1).decode()] = int(match.group(2))

    def replace(match):
        number = match.group(1).decode()
        if number not in lengths:
            return match.group(0)
        return b"/Length " + str(lengths[number]).encode()

    data = re.sub(rb"/Length\s+(\d+)\s+0\s+R", replace, data)

    # Drop the now-unused length objects.
    data = re.sub(rb"\n\d+ 0 obj\s*\n?\d+\s*\n?endobj", b"", data)
    return data


def ascii85_encode_streams(data):
    """Re-encode binary stream data as ASCII85.

    A PDF image XObject holds pixel bytes, which are not ASCII. PDF permits
    ASCII85Decode as a stream filter, and it composes with FlateDecode: the
    reader decodes ASCII85 first, then inflates.

    qpdf expands streams on the way in, so an image arrives here as raw RGB.
    The pixels are recompressed with zlib before being ASCII85 encoded;
    carrying the raw data would inflate the file by roughly thirty times.
    """
    def convert(match):
        dictionary, payload = match.group(1), match.group(2)
        if b"/Subtype /Image" not in dictionary:
            return match.group(0)

        raw = payload
        try:
            raw = zlib.decompress(payload)
        except Exception:
            pass

        try:
            compressed = zlib.compress(raw, 9)
        except Exception:
            compressed = raw

        encoded = base64.a85encode(compressed, adobe=False, pad=False)
        lines = b"\n".join(
            encoded[i:i + 120] for i in range(0, len(encoded), 120)
        )

        new_dictionary = re.sub(
            rb"/Filter\s*(?:\[[^\]]*\]|/\w+)",
            b"/Filter [/ASCII85Decode /FlateDecode]",
            dictionary,
        )
        new_dictionary = re.sub(rb"/Length\s+\d+\s+0\s+R", b"", new_dictionary)
        new_dictionary = new_dictionary.rstrip()
        if new_dictionary.endswith(b">>"):
            new_dictionary = new_dictionary[:-2]
        return (
            new_dictionary
            + b"/Length " + str(len(lines)).encode()
            + b">>\nstream\n" + lines + b"\nendstream"
        )

    pattern = re.compile(
        rb"(\d+ 0 obj\s*<<[^<>]*?/Subtype /Image.*?>>)"
        rb"\s*stream\r?\n(.*?)\r?\nendstream",
        re.S,
    )
    return pattern.sub(convert, data)


def strip_creation_date(data):
    """Remove the creation date, the derived checksum, and fix the document ID.

    LibreOffice stamps the time of conversion into the PDF, and qpdf derives a
    document checksum and a random ID from it. Left in place, two runs of the
    same workflow produce different bytes, so the samples never match what is
    published and every run shows as a change.

    The ID is replaced with a fixed value rather than a derived one, because
    the content it would be derived from is the very thing that changes
    between runs.
    """
    data = re.sub(rb"/CreationDate\s*\(D:[^)]*\)", b"", data)
    data = re.sub(rb"/ModDate\s*\(D:[^)]*\)", b"", data)
    data = re.sub(rb"/DocChecksum\s*/[0-9A-Fa-f]+", b"", data)
    data = re.sub(
        rb"/ID\s*\[[^\]]*\]",
        b"/ID [<00000000000000000000000000000000>"
        b"<00000000000000000000000000000000>]",
        data,
    )
    return data


def rebuild_xref(data):
    """Rebuild the cross-reference table after objects moved or vanished."""
    objects = {}
    for match in re.finditer(rb"(?m)^(\d+) 0 obj", data):
        objects[int(match.group(1))] = match.start()

    if not objects:
        raise RuntimeError("no objects found; is this a PDF?")

    xref_at = data.find(b"\nxref\n")
    trailer_at = data.find(b"\ntrailer ", xref_at)
    if xref_at < 0 or trailer_at < 0:
        raise RuntimeError("could not locate the cross-reference table")

    head = data[:xref_at]
    tail = data[trailer_at + 1:]

    highest = max(objects)
    entries = b"0000000000 65535 f \n"
    for number in range(1, highest + 1):
        if number in objects:
            entries += b"%010d 00000 n \n" % objects[number]
        else:
            entries += b"0000000000 65535 f \n"

    xref = b"xref\n0 " + str(highest + 1).encode() + b"\n" + entries
    rebuilt = head + xref + tail

    offset = rebuilt.find(b"xref\n")
    rebuilt = re.sub(rb"startxref\s+\d+", b"startxref " + str(offset).encode(), rebuilt)
    rebuilt = re.sub(rb"/Size\s+\d+", b"/Size " + str(highest + 1).encode(), rebuilt)
    return rebuilt


def to_ascii(pdf_path, out_path):
    with tempfile.TemporaryDirectory() as tmp:
        qdf_path = Path(tmp) / "expanded.pdf"
        qdf_expand(pdf_path, qdf_path)
        data = qdf_path.read_bytes()

    # Strip the creation date before qpdf computes the document ID, otherwise
    # the ID still changes between runs.
    data = strip_creation_date(data)
    data = strip_embedded_font(data)
    data = ascii85_encode_streams(data)
    data = strip_qdf_artifacts(data)
    data = resolve_indirect_lengths(data)
    data = strip_binary_marker(data)
    data = rebuild_xref(data)

    non_ascii = sum(1 for byte in data if byte > 127)
    if non_ascii:
        raise RuntimeError(f"{non_ascii} non-ASCII byte(s) remain after conversion")

    Path(out_path).write_bytes(data)
    return out_path


def main():
    if len(sys.argv) != 3:
        print(__doc__)
        return 2
    source = Path(sys.argv[1])
    target = Path(sys.argv[2])
    to_ascii(source, target)
    print(f"{source} -> {target} ({target.stat().st_size} bytes, pure ASCII)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
