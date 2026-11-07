"""docxlib — OOXML comment / revision primitives for the govt comment resolution skill.

Python 3 stdlib only. Every function that mutates a document takes a dict of
part-name -> bytes and returns a new dict, so callers control when and where
the package is written.

Correctness notes (established empirically against LibreOffice 7.4 and the
OpenXML SDK — see references/ooxml-comments.md):

1. A reply needs its own w:commentReference inside the parent's
   w:commentRangeStart/End, or Word and LibreOffice silently drop it.
2. Reply threading keys on w14:paraId, never w:id.
3. w15:done lives in commentsExtended.xml on w15:commentEx, and it must be set
   on the TARGET comment, not the first match.
4. Every new comment updates all four comment parts.
5. commentsExtensible.xml uses the openxmlformats content type; the ms-word
   variant is rejected by the SDK.
6. Deletions use w:delText, never w:t.
7. Content-types and relationships parts must serialise with a default
   namespace, not an ns0: prefix.
"""

import copy
import hashlib
import json
import re
import shutil
import subprocess
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
W16CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
W16CEX = "http://schemas.microsoft.com/office/word/2018/wordml/cex"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

CT_COMMENTS = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
CT_COMMENTS_EXTENDED = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"
CT_COMMENTS_IDS = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml"
CT_COMMENTS_EXTENSIBLE = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml"

REL_COMMENTS = f"{R}/comments"
REL_COMMENTS_EXTENDED = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"
REL_COMMENTS_IDS = "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds"
REL_COMMENTS_EXTENSIBLE = "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible"

COMMENTS_PART = "word/comments.xml"
COMMENTS_EXTENDED_PART = "word/commentsExtended.xml"
COMMENTS_IDS_PART = "word/commentsIds.xml"
COMMENTS_EXTENSIBLE_PART = "word/commentsExtensible.xml"

COMMENT_PARTS = (
    COMMENTS_PART,
    COMMENTS_EXTENDED_PART,
    COMMENTS_IDS_PART,
    COMMENTS_EXTENSIBLE_PART,
)


def q(ns, tag):
    return f"{{{ns}}}{tag}"


def _register(prefix_map):
    for prefix, uri in prefix_map.items():
        try:
            ET.register_namespace(prefix, uri)
        except ValueError:
            pass


_register({
    "w": W, "w14": W14, "w15": W15, "w16cid": W16CID, "w16cex": W16CEX, "r": R,
})


# --------------------------------------------------------------------------
# package IO
# --------------------------------------------------------------------------

def read_parts(docx_path):
    """Read every part of a .docx into a dict of name -> bytes."""
    with zipfile.ZipFile(docx_path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def write_parts(parts, out_path):
    """Write a parts dict back out as a .docx package."""
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    tmp = out_path.with_name(out_path.name + ".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as archive:
        for name in sorted(parts):
            archive.writestr(name, parts[name])
    shutil.move(str(tmp), str(out_path))
    return out_path


def sha256_file(path):
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(65536), b""):
            digest.update(chunk)
    return digest.hexdigest()


def sha256_bytes(data):
    return hashlib.sha256(data).hexdigest()


# --------------------------------------------------------------------------
# XML helpers
# --------------------------------------------------------------------------

def _serialize(root, default_ns=None):
    """Serialise an ElementTree root, optionally forcing a default namespace.

    OPC readers reject parts whose root is prefixed (ns0:Types) instead of
    using a default namespace, so content-types and relationships parts must be
    serialised with the default namespace registered.
    """
    if default_ns:
        ET.register_namespace("", default_ns)
    return ET.tostring(root, encoding="UTF-8", xml_declaration=True)


def _find_parent(root, target):
    for candidate in root.iter():
        if target in list(candidate):
            return candidate
    return None


def _text_of(element):
    return "".join(node.text or "" for node in element.iter(q(W, "t")))


# --------------------------------------------------------------------------
# comment parsing
# --------------------------------------------------------------------------

def parse_comments(parts):
    """Parse comments.xml into a list of comment dicts.

    Each dict carries the w:id, author, date, body text, and the w14:paraId of
    the comment's first paragraph (the key reply threading depends on).
    """
    raw = parts.get(COMMENTS_PART)
    if raw is None:
        return []
    root = ET.fromstring(raw)
    out = []
    for comment in root.findall(q(W, "comment")):
        paragraph = comment.find(q(W, "p"))
        para_id = paragraph.get(q(W14, "paraId")) if paragraph is not None else None
        out.append({
            "w_id": comment.get(q(W, "id")),
            "author": comment.get(q(W, "author")) or "",
            "initials": comment.get(q(W, "initials")) or "",
            "date": comment.get(q(W, "date")) or "",
            "text": _text_of(comment),
            "para_id": para_id,
        })
    return out


def parse_comments_extended(parts):
    """Parse commentsExtended.xml: paraId, parent paraId, and the done flag."""
    raw = parts.get(COMMENTS_EXTENDED_PART)
    if raw is None:
        return []
    root = ET.fromstring(raw)
    out = []
    for entry in root.findall(q(W15, "commentEx")):
        out.append({
            "para_id": entry.get(q(W15, "paraId")),
            "parent_para_id": entry.get(q(W15, "paraIdParent")),
            "done": (entry.get(q(W15, "done")) or "0") == "1",
        })
    return out


def parse_comments_ids(parts):
    """Parse commentsIds.xml: the durable (survives save) comment identity."""
    raw = parts.get(COMMENTS_IDS_PART)
    if raw is None:
        return []
    root = ET.fromstring(raw)
    out = []
    for entry in root.findall(q(W16CID, "commentId")):
        out.append({
            "para_id": entry.get(q(W16CID, "paraId")),
            "durable_id": entry.get(q(W16CID, "durableId")),
        })
    return out


def extract_anchors(parts):
    """Map w:id -> quoted text, by walking commentRangeStart/End in the body.

    Returns (anchors, problems) where problems records markers with no
    matching comment definition and comments with no markers.
    """
    raw = parts.get("word/document.xml")
    anchors = {}
    if raw is None:
        return anchors, []
    root = ET.fromstring(raw)
    body = root.find(q(W, "body"))
    if body is None:
        return anchors, []

    open_ids = []
    buffers = {}
    for element in body.iter():
        if element.tag == q(W, "commentRangeStart"):
            cid = element.get(q(W, "id"))
            buffers[cid] = []
            open_ids.append(cid)
        elif element.tag == q(W, "commentRangeEnd"):
            cid = element.get(q(W, "id"))
            if cid in open_ids:
                open_ids.remove(cid)
        elif element.tag == q(W, "t"):
            for cid in open_ids:
                buffers[cid].append(element.text or "")

    for cid, chunks in buffers.items():
        anchors[cid] = "".join(chunks).strip()

    problems = []
    defined = {c["w_id"] for c in parse_comments(parts)}
    for cid in anchors:
        if cid not in defined:
            problems.append({"kind": "dangling_range", "w_id": cid})
    for cid in defined:
        if cid not in anchors:
            problems.append({"kind": "not_anchored", "w_id": cid})
    return anchors, problems


def thread_structure(parts):
    """Build parent -> children paraId mapping from commentsExtended.xml."""
    children = {}
    for entry in parse_comments_extended(parts):
        parent = entry["parent_para_id"]
        if parent:
            children.setdefault(parent, []).append(entry["para_id"])
    return children


def descendants(parts, para_id):
    """All paraIds in the reply subtree rooted at para_id (excluding itself)."""
    children = thread_structure(parts)
    out = []
    stack = list(children.get(para_id, []))
    while stack:
        current = stack.pop()
        out.append(current)
        stack.extend(children.get(current, []))
    return out


# --------------------------------------------------------------------------
# comment writing
# --------------------------------------------------------------------------

def _ensure_content_type(parts, part_name, content_type):
    root = ET.fromstring(parts["[Content_Types].xml"])
    for override in root:
        if override.get("PartName") == part_name:
            return
    ET.SubElement(root, f"{{{CT_NS}}}Override", {
        "PartName": part_name,
        "ContentType": content_type,
    })
    parts["[Content_Types].xml"] = _serialize(root, CT_NS)


def _ensure_relationship(parts, rel_type, target):
    root = ET.fromstring(parts["word/_rels/document.xml.rels"])
    for rel in root:
        if (rel.get("Type") or "") == rel_type:
            return
    used = {rel.get("Id") for rel in root}
    index = 1
    while f"rId{index}" in used:
        index += 1
    ET.SubElement(root, f"{{{RELS_NS}}}Relationship", {
        "Id": f"rId{index}",
        "Type": rel_type,
        "Target": target,
    })
    parts["word/_rels/document.xml.rels"] = _serialize(root, RELS_NS)


def _next_numeric_id(existing):
    best = 0
    for value in existing:
        try:
            best = max(best, int(value))
        except (TypeError, ValueError):
            continue
    return best + 1


def _next_revision_id(parts):
    raw = parts.get("word/document.xml", b"")
    ids = [int(m) for m in re.findall(rb'w:id="(\d+)"', raw)]
    return max(ids) + 1 if ids else 1


def _make_para_id(n):
    return f"{n:08X}"


def add_reply(parts, parent_w_id, text, author, initials=None, date=None):
    """Add a threaded reply to a comment.

    Implements all four correctness requirements: w14:paraId on the reply
    paragraph, w15:paraIdParent pointing at the parent's paraId, a
    w:commentReference inside the parent's range in document.xml, and updates
    to all four comment parts.
    """
    comments = parse_comments(parts)
    parent = next((c for c in comments if c["w_id"] == str(parent_w_id)), None)
    if parent is None:
        raise KeyError(f"comment {parent_w_id} not found")
    if not parent["para_id"]:
        raise KeyError(
            f"comment {parent_w_id} has no w14:paraId; cannot thread a reply to it"
        )

    next_id = _next_numeric_id([c["w_id"] for c in comments])
    new_para_id = _make_para_id(next_id)
    durable_id = _make_para_id(next_id)
    initials = initials or (author[:1] or "A").upper()
    date = date or "1970-01-01T00:00:00Z"

    # 1) comments.xml
    root = ET.fromstring(parts[COMMENTS_PART])
    comment = ET.SubElement(root, q(W, "comment"), {
        q(W, "id"): str(next_id),
        q(W, "author"): author,
        q(W, "initials"): initials,
        q(W, "date"): date,
    })
    paragraph = ET.SubElement(comment, q(W, "p"), {q(W14, "paraId"): new_para_id})
    run = ET.SubElement(paragraph, q(W, "r"))
    node = ET.SubElement(run, q(W, "t"))
    node.text = text
    node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    parts[COMMENTS_PART] = _serialize(root)

    # 2) commentsExtended.xml — paraId + paraIdParent
    root = ET.fromstring(parts[COMMENTS_EXTENDED_PART])
    ET.SubElement(root, q(W15, "commentEx"), {
        q(W15, "paraId"): new_para_id,
        q(W15, "paraIdParent"): parent["para_id"],
        q(W15, "done"): "0",
    })
    parts[COMMENTS_EXTENDED_PART] = _serialize(root)

    # 3) commentsIds.xml — durable identity
    root = ET.fromstring(parts[COMMENTS_IDS_PART])
    ET.SubElement(root, q(W16CID, "commentId"), {
        q(W16CID, "paraId"): new_para_id,
        q(W16CID, "durableId"): durable_id,
    })
    parts[COMMENTS_IDS_PART] = _serialize(root)

    # 4) commentsExtensible.xml — dateUtc
    root = ET.fromstring(parts[COMMENTS_EXTENSIBLE_PART])
    ET.SubElement(root, q(W16CEX, "commentExtensible"), {
        q(W16CEX, "durableId"): durable_id,
        q(W16CEX, "dateUtc"): date,
    })
    parts[COMMENTS_EXTENSIBLE_PART] = _serialize(root)

    # 5) document.xml — the reply's own commentReference inside the parent's
    #    range, otherwise the reply is silently dropped by Office consumers.
    root = ET.fromstring(parts["word/document.xml"])
    body = root.find(q(W, "body"))
    placed = False
    for paragraph in body.iter(q(W, "p")):
        starts = {
            node.get(q(W, "id")) for node in paragraph.findall(q(W, "commentRangeStart"))
        }
        if str(parent_w_id) not in starts:
            continue
        reference_run = ET.Element(q(W, "r"))
        ET.SubElement(reference_run, q(W, "commentReference"), {q(W, "id"): str(next_id)})
        paragraph.append(reference_run)
        placed = True
        break
    if not placed:
        raise KeyError(
            f"no commentRangeStart for comment {parent_w_id} in document.xml; "
            "a reply cannot be anchored"
        )
    parts["word/document.xml"] = _serialize(root)

    return {
        "w_id": str(next_id),
        "para_id": new_para_id,
        "durable_id": durable_id,
        "parent_para_id": parent["para_id"],
    }


def resolve_comment(parts, w_id, include_replies=True):
    """Set w15:done="1" on a comment (and its replies) in commentsExtended.xml.

    Operates on the TARGET comment's paraId, not the first matching element.
    """
    comments = parse_comments(parts)
    target = next((c for c in comments if c["w_id"] == str(w_id)), None)
    if target is None:
        raise KeyError(f"comment {w_id} not found")
    if not target["para_id"]:
        raise KeyError(f"comment {w_id} has no w14:paraId")

    affected = {target["para_id"]}
    if include_replies:
        affected.update(descendants(parts, target["para_id"]))

    root = ET.fromstring(parts[COMMENTS_EXTENDED_PART])
    touched = False
    for entry in root.findall(q(W15, "commentEx")):
        if entry.get(q(W15, "paraId")) in affected:
            entry.set(q(W15, "done"), "1")
            touched = True
    if not touched:
        raise KeyError(f"no commentEx entry for comment {w_id}")
    parts[COMMENTS_EXTENDED_PART] = _serialize(root)
    return sorted(affected)


def ensure_comment_parts(parts):
    """Create any missing comment parts so the writer can operate."""
    if COMMENTS_PART not in parts:
        parts[COMMENTS_PART] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:comments xmlns:w="{W}"/>'
        ).encode("utf-8")
    if COMMENTS_EXTENDED_PART not in parts:
        parts[COMMENTS_EXTENDED_PART] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w15:commentsEx xmlns:w15="{W15}"/>'
        ).encode("utf-8")
    if COMMENTS_IDS_PART not in parts:
        parts[COMMENTS_IDS_PART] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w16cid:commentsIds xmlns:w16cid="{W16CID}"/>'
        ).encode("utf-8")
    if COMMENTS_EXTENSIBLE_PART not in parts:
        parts[COMMENTS_EXTENSIBLE_PART] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w16cex:commentsExtensible xmlns:w16cex="{W16CEX}"/>'
        ).encode("utf-8")
    _ensure_content_type(parts, "/word/comments.xml", CT_COMMENTS)
    _ensure_content_type(parts, "/word/commentsExtended.xml", CT_COMMENTS_EXTENDED)
    _ensure_content_type(parts, "/word/commentsIds.xml", CT_COMMENTS_IDS)
    _ensure_content_type(parts, "/word/commentsExtensible.xml", CT_COMMENTS_EXTENSIBLE)
    _ensure_relationship(parts, REL_COMMENTS, "comments.xml")
    _ensure_relationship(parts, REL_COMMENTS_EXTENDED, "commentsExtended.xml")
    _ensure_relationship(parts, REL_COMMENTS_IDS, "commentsIds.xml")
    _ensure_relationship(parts, REL_COMMENTS_EXTENSIBLE, "commentsExtensible.xml")
    return parts


# --------------------------------------------------------------------------
# tracked changes
# --------------------------------------------------------------------------

def _make_text_run(template_run, text):
    """A run carrying `text`, reusing the template run's formatting."""
    run = ET.Element(q(W, "r"))
    rpr = template_run.find(q(W, "rPr")) if template_run is not None else None
    if rpr is not None:
        run.append(copy.deepcopy(rpr))
    node = ET.SubElement(run, q(W, "t"))
    node.text = text
    node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return run


def apply_tracked_replace(parts, find_text, replace_text, author, date=None):
    """Replace text in the body as a tracked change (w:del + w:ins).

    The surviving text keeps its original reading order:

        <w:r>head</w:r> <w:del>find</w:del> <w:ins>replace</w:ins> <w:r>tail</w:r>

    Returns the revision id used, or None when find_text was not present.
    """
    date = date or "1970-01-01T00:00:00Z"
    root = ET.fromstring(parts["word/document.xml"])
    body = root.find(q(W, "body"))

    target_text = None
    for node in body.iter(q(W, "t")):
        if node.text and find_text in node.text:
            target_text = node
            break
    if target_text is None:
        return None

    run = _find_parent(body, target_text)
    if run is None:
        return None
    paragraph = _find_parent(body, run)
    if paragraph is None:
        return None

    original = target_text.text
    head, _, tail = original.partition(find_text)
    index = list(paragraph).index(run)
    rev = _next_revision_id(parts)

    def insert_at(position, element):
        paragraph.insert(position, element)
        return position + 1

    # head (text before the match) stays first
    if head:
        index = insert_at(index, _make_text_run(run, head))

    # the deletion
    deletion = ET.Element(q(W, "del"), {
        q(W, "id"): str(rev),
        q(W, "author"): author,
        q(W, "date"): date,
    })
    del_run = ET.SubElement(deletion, q(W, "r"))
    rpr = run.find(q(W, "rPr"))
    if rpr is not None:
        del_run.append(copy.deepcopy(rpr))
    del_text = ET.SubElement(del_run, q(W, "delText"))
    del_text.text = find_text
    del_text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    index = insert_at(index, deletion)

    # the insertion
    if replace_text:
        insertion = ET.Element(q(W, "ins"), {
            q(W, "id"): str(rev + 1),
            q(W, "author"): author,
            q(W, "date"): date,
        })
        ins_run = ET.SubElement(insertion, q(W, "r"))
        if rpr is not None:
            ins_run.append(copy.deepcopy(rpr))
        ins_text = ET.SubElement(ins_run, q(W, "t"))
        ins_text.text = replace_text
        ins_text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        index = insert_at(index, insertion)

    # tail (text after the match) must come LAST, or the sentence order breaks
    if tail:
        insert_at(index, _make_text_run(run, tail))

    paragraph.remove(run)
    parts["word/document.xml"] = _serialize(root)
    return rev


def accept_all_revisions(parts):
    """Accept every tracked change: drop w:del, unwrap w:ins."""
    root = ET.fromstring(parts["word/document.xml"])
    body = root.find(q(W, "body"))

    for deletion in list(body.iter(q(W, "del"))):
        parent = _find_parent(body, deletion)
        if parent is not None:
            parent.remove(deletion)

    for insertion in list(body.iter(q(W, "ins"))):
        parent = _find_parent(body, insertion)
        if parent is None:
            continue
        position = list(parent).index(insertion)
        for offset, child in enumerate(list(insertion)):
            parent.insert(position + offset, child)
        parent.remove(insertion)

    parts["word/document.xml"] = _serialize(root)
    return parts


def strip_comments(parts):
    """Remove every comment: parts, anchors, references, content types, rels.

    Only the run carrying a w:commentReference is removed; runs holding real
    text survive.
    """
    root = ET.fromstring(parts["word/document.xml"])
    body = root.find(q(W, "body"))

    for reference in list(body.iter(q(W, "commentReference"))):
        run = _find_parent(body, reference)
        if run is None:
            continue
        parent = _find_parent(body, run)
        if parent is not None:
            parent.remove(run)

    for tag in ("commentRangeStart", "commentRangeEnd"):
        for element in list(body.iter(q(W, tag))):
            parent = _find_parent(body, element)
            if parent is not None:
                parent.remove(element)

    parts["word/document.xml"] = _serialize(root)

    for part in COMMENT_PARTS:
        parts.pop(part, None)

    ct_root = ET.fromstring(parts["[Content_Types].xml"])
    for override in list(ct_root):
        name = override.get("PartName") or ""
        if any(key in name for key in (
            "comments.xml", "commentsExtended", "commentsIds", "commentsExtensible"
        )):
            ct_root.remove(override)
    parts["[Content_Types].xml"] = _serialize(ct_root, CT_NS)

    rels_root = ET.fromstring(parts["word/_rels/document.xml.rels"])
    for rel in list(rels_root):
        target = rel.get("Target") or ""
        rel_type = rel.get("Type") or ""
        if "comments" in target or "comments" in rel_type:
            rels_root.remove(rel)
    parts["word/_rels/document.xml.rels"] = _serialize(rels_root, RELS_NS)
    return parts


# --------------------------------------------------------------------------
# rendering
# --------------------------------------------------------------------------

def soffice_to_pdf(docx_path, out_pdf):
    """Convert a .docx to PDF with LibreOffice."""
    docx_path = Path(docx_path)
    out_pdf = Path(out_pdf)
    out_pdf.parent.mkdir(parents=True, exist_ok=True)
    result = subprocess.run(
        ["soffice", "--headless", "--convert-to", "pdf", "--outdir",
         str(out_pdf.parent), str(docx_path)],
        capture_output=True, text=True, timeout=600,
    )
    produced = out_pdf.parent / (docx_path.stem + ".pdf")
    if not produced.exists():
        raise RuntimeError(
            f"LibreOffice produced no PDF for {docx_path}: {result.stderr.strip()}"
        )
    if produced != out_pdf:
        shutil.move(str(produced), str(out_pdf))
    return out_pdf


def pdf_page_texts(pdf_path):
    """Per-page text of a PDF, via pypdf."""
    from pypdf import PdfReader

    reader = PdfReader(str(pdf_path))
    return [(page.extract_text() or "") for page in reader.pages]


def pdf_page_images(pdf_path, out_dir, dpi=110):
    """Render PDF pages to PNG via pdftoppm. Returns list of file paths."""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    prefix = out_dir / "page"
    subprocess.run(
        ["pdftoppm", "-r", str(dpi), "-png", str(pdf_path), str(prefix)],
        capture_output=True, text=True, timeout=600, check=True,
    )
    return sorted(out_dir.glob("page-*.png"))
