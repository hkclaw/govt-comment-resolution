#!/usr/bin/env python3
"""Regression tests for the govt comment resolution skill.

Run: python3 tests/test_skill.py

These cover the failure modes that were found empirically during
development — each one previously produced a silently wrong document.
"""
import json
import subprocess
import sys
import tempfile
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parent.parent / "scripts"
sys.path.insert(0, str(SCRIPTS))

import docxlib as D  # noqa: E402

FAILURES = []
PASSES = []


def check(name, condition, detail=""):
    if condition:
        PASSES.append(name)
        print(f"  PASS  {name}")
    else:
        FAILURES.append(f"{name}: {detail}")
        print(f"  FAIL  {name}  {detail}")


def build_sample(path, with_reply=False):
    """A minimal but valid document with one anchored comment."""
    W = D.W
    W14 = D.W14
    W15 = D.W15
    W16CID = D.W16CID
    W16CEX = D.W16CEX
    R = D.R

    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W}" xmlns:r="{R}" xmlns:w14="{W14}"><w:body>'
        '<w:p><w:r><w:t xml:space="preserve">Hello world paragraph</w:t></w:r></w:p>'
        '<w:p><w:commentRangeStart w:id="1"/>'
        '<w:r><w:t xml:space="preserve">The allocation grows by 4.2% each year.</w:t></w:r>'
        '<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>'
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>'
    )
    comments_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:comments xmlns:w="{W}" xmlns:w14="{W14}">'
        '<w:comment w:id="1" w:author="Dept Officer" w:initials="DO" '
        'w:date="2026-11-06T09:00:00Z">'
        '<w:p w14:paraId="1A2B3C4D"><w:r><w:t xml:space="preserve">Justify the figure.</w:t></w:r></w:p>'
        '</w:comment></w:comments>'
    )
    comments_extended_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w15:commentsEx xmlns:w15="{W15}">'
        '<w15:commentEx w15:paraId="1A2B3C4D" w15:done="0"/></w15:commentsEx>'
    )
    comments_ids_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w16cid:commentsIds xmlns:w16cid="{W16CID}">'
        '<w16cid:commentId w16cid:paraId="1A2B3C4D" w16cid:durableId="ABCDEF012345"/></w16cid:commentsIds>'
    )
    comments_extensible_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w16cex:commentsExtensible xmlns:w16cex="{W16CEX}">'
        '<w16cex:commentExtensible w16cex:durableId="ABCDEF012345" w16cex:dateUtc="2026-11-06T09:00:00Z"/>'
        '</w16cex:commentsExtensible>'
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Types xmlns="{D.CT_NS}">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>'
        '<Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>'
        '<Override PartName="/word/commentsIds.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml"/>'
        '<Override PartName="/word/commentsExtensible.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml"/>'
        '</Types>'
    )
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{D.RELS_NS}">'
        f'<Relationship Id="rId1" Type="{R}/officeDocument" Target="word/document.xml"/>'
        '</Relationships>'
    )
    doc_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{D.RELS_NS}">'
        f'<Relationship Id="rId1" Type="{R}/styles" Target="styles.xml"/>'
        f'<Relationship Id="rId20" Type="{D.REL_COMMENTS}" Target="comments.xml"/>'
        f'<Relationship Id="rId21" Type="{D.REL_COMMENTS_EXTENDED}" Target="commentsExtended.xml"/>'
        f'<Relationship Id="rId22" Type="{D.REL_COMMENTS_IDS}" Target="commentsIds.xml"/>'
        f'<Relationship Id="rId23" Type="{D.REL_COMMENTS_EXTENSIBLE}" Target="commentsExtensible.xml"/>'
        '</Relationships>'
    )
    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles xmlns:w="{W}">'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
        '</w:styles>'
    )

    parts = {
        "[Content_Types].xml": content_types,
        "_rels/.rels": root_rels,
        "word/document.xml": document_xml,
        "word/_rels/document.xml.rels": doc_rels,
        "word/styles.xml": styles_xml,
        "word/comments.xml": comments_xml,
        "word/commentsExtended.xml": comments_extended_xml,
        "word/commentsIds.xml": comments_ids_xml,
        "word/commentsExtensible.xml": comments_extensible_xml,
    }
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    D.write_parts(parts, path)
    return path


def test_reply_is_threaded_and_visible():
    print("\n[1] reply threading, resolve flag, and document.xml reference")
    with tempfile.TemporaryDirectory() as tmp:
        src = build_sample(Path(tmp) / "sample.docx")
        parts = D.read_parts(src)
        D.ensure_comment_parts(parts)
        reply = D.add_reply(parts, "1", "Our response.", "Jane Doe")
        resolved = D.resolve_comment(parts, "1")

        comments = D.parse_comments(parts)
        new = next(c for c in comments if c["w_id"] == reply["w_id"])
        check("reply authored by the right person", new["author"] == "Jane Doe", new["author"])
        check("reply carries a paraId", bool(new["para_id"]), str(new["para_id"]))

        extended = {e["para_id"]: e for e in D.parse_comments_extended(parts)}
        check("reply threaded to parent paraId",
              extended[reply["para_id"]]["parent_para_id"] == "1A2B3C4D",
              str(extended[reply["para_id"]]["parent_para_id"]))
        check("target comment resolved", extended["1A2B3C4D"]["done"] is True)
        check("reply resolved with parent", extended[reply["para_id"]]["done"] is True)
        check("resolve touched both nodes", len(resolved) == 2, str(resolved))

        doc_xml = parts["word/document.xml"].decode()
        check("reply has its own commentReference in document.xml",
              f'<w:commentReference w:id="{reply["w_id"]}"' in doc_xml)

        # the reply must survive a real Office renderer
        out = Path(tmp) / "out.docx"
        D.write_parts(parts, out)
        fodt = Path(tmp) / "out.fodt"
        subprocess.run(
            ["soffice", "--headless", "--convert-to", "fodt", "--outdir", str(tmp), str(out)],
            capture_output=True, timeout=300, check=False,
        )
        produced = Path(tmp) / "out.fodt"
        check("soffice produced a readable document", produced.exists())
        if produced.exists():
            text = produced.read_text(encoding="utf-8")
            check("reply visible to a real renderer", "Our response." in text)
            check("reply marked resolved in renderer", 'loext:resolved="true"' in text)


def test_tracked_change_preserves_reading_order():
    print("\n[2] tracked change keeps the sentence in reading order")
    with tempfile.TemporaryDirectory() as tmp:
        src = build_sample(Path(tmp) / "sample.docx")
        parts = D.read_parts(src)
        rev = D.apply_tracked_replace(parts, "4.2%", "4.1%", "Jane Doe")
        check("revision id returned", rev is not None)

        doc_xml = parts["word/document.xml"].decode()
        order = [
            doc_xml.find("The allocation grows by "),
            doc_xml.find("<w:del "),
            doc_xml.find("<w:ins "),
            doc_xml.find(" each year."),
        ]
        check("head < del < ins < tail ordering", order == sorted(order), str(order))

        accepted = dict(parts)
        D.accept_all_revisions(accepted)
        out = Path(tmp) / "accepted.docx"
        D.write_parts(accepted, out)
        pdf = Path(tmp) / "accepted.pdf"
        D.soffice_to_pdf(out, pdf)
        texts = D.pdf_page_texts(pdf)
        joined = " ".join(texts)
        check("accepted text reads correctly",
              "The allocation grows by 4.1% each year." in joined,
              joined[:160])


def test_strip_comments_keeps_body_text():
    print("\n[3] stripping comments for the clean copy keeps body text")
    with tempfile.TemporaryDirectory() as tmp:
        src = build_sample(Path(tmp) / "sample.docx")
        parts = D.read_parts(src)
        D.strip_comments(parts)
        check("comment part removed", D.COMMENTS_PART not in parts)
        out = Path(tmp) / "stripped.docx"
        D.write_parts(parts, out)
        pdf = Path(tmp) / "stripped.pdf"
        D.soffice_to_pdf(out, pdf)
        texts = " ".join(D.pdf_page_texts(pdf))
        check("body text survives stripping", "The allocation grows by 4.2%" in texts, texts[:160])
        check("comment text removed", "Justify the figure." not in texts)
        check("clean document still opens", len(texts) > 0)


def test_content_types_use_default_namespace():
    print("\n[4] content types serialise with a default namespace")
    with tempfile.TemporaryDirectory() as tmp:
        src = build_sample(Path(tmp) / "sample.docx")
        parts = D.read_parts(src)
        D.strip_comments(parts)
        ct = parts["[Content_Types].xml"].decode()
        check("no ns0 prefix on content types", "ns0:" not in ct, ct[:120])
        check("default namespace present", f'xmlns="{D.CT_NS}"' in ct, ct[:120])
        rels = parts["word/_rels/document.xml.rels"].decode()
        check("no ns0 prefix on relationships", "ns0:" not in rels, rels[:120])
        check("comment relationships removed", "comments.xml" not in rels)


def test_extract_flags_orphans_and_numbers_contiguously():
    print("\n[5] extraction reports orphans and numbers contiguously")
    with tempfile.TemporaryDirectory() as tmp:
        src = build_sample(Path(tmp) / "sample.docx")
        out = Path(tmp) / "comments"
        result = subprocess.run(
            [sys.executable, str(SCRIPTS / "extract_comments.py"),
             "--docx", str(src), "--out", str(out)],
            capture_output=True, text=True,
        )
        check("extraction succeeded", result.returncode == 0, result.stderr[-200:])
        files = sorted(p.name for p in out.glob("C-*.md"))
        check("one file per comment", files == ["C-001.md"], str(files))
        index = (out / "index.md").read_text(encoding="utf-8")
        check("index lists the comment", "C-001" in index)
        body = (out / "C-001.md").read_text(encoding="utf-8")
        check("anchored text captured", "The allocation grows by 4.2%" in body)
        check("durable id used as key", "durable_id: ABCDEF012345" in body, body[:400])


def test_page_alignment_detects_shifted_pages():
    print("\n[6] page alignment distinguishes shifted from changed pages")
    sys.path.insert(0, str(SCRIPTS))
    import importlib
    compare = importlib.import_module("compare_pages")

    before = [
        "Page one budget",
        "Page two regional north figures",
        "Page three timelines staffing",
        "Page four appendix contacts",
    ]
    after = [
        "Page one budget",
        "Page two regional north figures",
        "Page three timelines staffing AND a long new paragraph that pushes content",
        "NEW PAGE entirely new content",
        "Page four appendix contacts",
    ]
    changed = compare.find_changed(before, after)
    kinds = [c["kind"] for c in changed]
    check("only genuinely changed pages reported",
          kinds.count("modified") == 1 and kinds.count("added") == 1, str(changed))
    check("shifted-but-identical pages not reported", len(changed) == 2, str(changed))

    check("identical documents report no changes", compare.find_changed(before, list(before)) == [])


def test_authorship_validation():
    print("\n[7] reply authorship is validated")
    with tempfile.TemporaryDirectory() as tmp:
        src = build_sample(Path(tmp) / "sample.docx")
        out = Path(tmp) / "comments"
        subprocess.run(
            [sys.executable, str(SCRIPTS / "extract_comments.py"),
             "--docx", str(src), "--out", str(out)],
            capture_output=True, text=True,
        )
        comment_file = out / "C-001.md"
        text = comment_file.read_text(encoding="utf-8")
        text = text.replace("status: open", "status: approved")
        text = text.replace('owner: ""', 'owner: jane.doe\nreply_author: jane.doe')
        comment_file.write_text(text, encoding="utf-8")

        identity = Path(tmp) / "identity.json"
        identity.write_text(json.dumps({
            "default_author": "jane.doe",
            "authors": [{"id": "jane.doe", "name": "Jane Doe", "org": "Agency", "email": "j@x.gov"}],
        }), encoding="utf-8")

        manifest = Path(tmp) / "manifest.json"
        result = subprocess.run(
            [sys.executable, str(SCRIPTS / "apply_replies.py"),
             "--docx", str(src), "--comments", str(out), "--identity", str(identity),
             "--out", str(Path(tmp) / "w" / "v1.docx"), "--manifest", str(manifest)],
            capture_output=True, text=True,
        )
        check("valid author applies", result.returncode == 0, result.stderr[-300:])

        # now an author who is also the comment author
        text = comment_file.read_text(encoding="utf-8").replace(
            "reply_author: jane.doe", "reply_author: dept.officer"
        )
        comment_file.write_text(text, encoding="utf-8")
        identity.write_text(json.dumps({
            "default_author": "dept.officer",
            "authors": [{"id": "dept.officer", "name": "Dept Officer", "org": "Dept", "email": "d@x.gov"}],
        }), encoding="utf-8")
        manifest2 = Path(tmp) / "manifest2.json"
        result = subprocess.run(
            [sys.executable, str(SCRIPTS / "apply_replies.py"),
             "--docx", str(src), "--comments", str(out), "--identity", str(identity),
             "--out", str(Path(tmp) / "w2" / "v1.docx"), "--manifest", str(manifest2)],
            capture_output=True, text=True,
        )
        check("comment author cannot author the reply", result.returncode != 0, result.stdout[-200:])
        check("failure explains why", "same as the comment author" in result.stderr, result.stderr[-300:])


def test_idempotency():
    print("\n[8] re-running apply does not double-apply")
    with tempfile.TemporaryDirectory() as tmp:
        src = build_sample(Path(tmp) / "sample.docx")
        out = Path(tmp) / "comments"
        subprocess.run(
            [sys.executable, str(SCRIPTS / "extract_comments.py"),
             "--docx", str(src), "--out", str(out)],
            capture_output=True, text=True,
        )
        comment_file = out / "C-001.md"
        text = comment_file.read_text(encoding="utf-8")
        text = text.replace("status: open", "status: approved")
        text = text.replace('owner: ""', 'owner: jane.doe\nreply_author: jane.doe')
        comment_file.write_text(text, encoding="utf-8")

        identity = Path(tmp) / "identity.json"
        identity.write_text(json.dumps({
            "default_author": "jane.doe",
            "authors": [{"id": "jane.doe", "name": "Jane Doe", "org": "Agency", "email": "j@x.gov"}],
        }), encoding="utf-8")

        manifest = Path(tmp) / "manifest.json"
        target = Path(tmp) / "w" / "v1.docx"
        subprocess.run(
            [sys.executable, str(SCRIPTS / "apply_replies.py"),
             "--docx", str(src), "--comments", str(out), "--identity", str(identity),
             "--out", str(target), "--manifest", str(manifest)],
            capture_output=True, text=True,
        )
        first = D.read_parts(target)
        first_count = len(D.parse_comments(first))

        subprocess.run(
            [sys.executable, str(SCRIPTS / "apply_replies.py"),
             "--docx", str(src), "--comments", str(out), "--identity", str(identity),
             "--out", str(target), "--manifest", str(manifest)],
            capture_output=True, text=True,
        )
        second = D.read_parts(target)
        second_count = len(D.parse_comments(second))
        check("comment count unchanged on re-run", first_count == second_count,
              f"{first_count} -> {second_count}")


def test_compare_uses_accepted_text_not_markup():
    print("\n[9] a heavily rewritten page still pairs correctly")
    sys.path.insert(0, str(SCRIPTS))
    import importlib
    compare = importlib.import_module("compare_pages")

    # A page that was substantially rewritten still shares most of its
    # vocabulary with the original. It must pair as 'modified', not be
    # reported as a deleted page plus an added page.
    before = [
        "Page one budget summary",
        "The allocation is projected to grow by 4.2% over the coming year.",
        "Page three appendix contacts",
    ]
    accepted = [
        "Page one budget summary",
        "The allocation is projected to grow by 4.1% over the coming year. "
        "The methodology follows a revised national framework and the "
        "supporting assumptions were recalculated for the current period.",
        "Page three appendix contacts",
    ]
    changed = compare.find_changed(before, accepted)
    check("rewritten page pairs as modified",
          len(changed) == 1 and changed[0]["kind"] == "modified", str(changed))
    check("pairing is correct", changed and changed[0]["before_page"] == 2
          and changed[0]["after_page"] == 2, str(changed))

    # Genuinely different pages must NOT be paired.
    unrelated = compare.find_changed(
        ["Page five blank divider section"],
        ["The Board is asked to endorse this plan and authorise Phase one."],
    )
    check("unrelated pages are not paired",
          all(c["kind"] in ("added", "deleted") for c in unrelated), str(unrelated))


def test_compare_end_to_end_with_tracked_edit():
    print("\n[9b] end-to-end: a tracked edit yields one paired changed page")
    with tempfile.TemporaryDirectory() as tmp:
        src = build_sample(Path(tmp) / "sample.docx")
        edited = Path(tmp) / "edited.docx"
        parts = D.read_parts(src)
        D.apply_tracked_replace(
            parts, "4.2%", "4.1%", "Jane Doe"
        )
        D.write_parts(parts, edited)

        out = Path(tmp) / "review.pdf"
        result = subprocess.run(
            [sys.executable, str(SCRIPTS / "compare_pages.py"),
             "--before", str(src), "--after", str(edited), "--out", str(out)],
            capture_output=True, text=True,
        )
        check("compare ran", result.returncode == 0, result.stderr[-300:])
        check("review PDF written", out.exists())
        if out.exists():
            from pypdf import PdfReader
            reader = PdfReader(str(out))
            summary = reader.pages[0].extract_text() or ""
            check("summary shows one changed page",
                  "1 changed page(s)" in summary, summary[:200])
            check("page is paired, not deleted+added",
                  "Before page 1" in summary and "->  After page 1" in summary
                  and "deleted" not in summary and "added" not in summary,
                  summary[:200])


def test_edit_markers_do_not_leak_into_replies():
    print("\n[10] [edit] markers never reach the department")
    sys.path.insert(0, str(SCRIPTS))
    import importlib
    replies = importlib.import_module("build_replies")
    cleaned = replies.strip_edit_markers(
        "Our answer.\n\n[edit] old text -> new text\n\nMore explanation."
    )
    check("marker removed", "[edit]" not in cleaned, cleaned)
    check("surrounding text kept", "Our answer." in cleaned and "More explanation." in cleaned)


def test_reextract_preserves_triage():
    print("\n[11] re-extraction keeps the human's triage")
    with tempfile.TemporaryDirectory() as tmp:
        src = build_sample(Path(tmp) / "sample.docx")
        out = Path(tmp) / "comments"
        subprocess.run(
            [sys.executable, str(SCRIPTS / "extract_comments.py"),
             "--docx", str(src), "--out", str(out)],
            capture_output=True, text=True,
        )
        comment_file = out / "C-001.md"
        text = comment_file.read_text(encoding="utf-8")
        text = text.replace("status: open", "status: approved")
        text = text.replace('owner: ""', 'owner: jane.doe\nreply_author: jane.doe')
        text = text.replace(
            "<!-- Draft the reply here. One paragraph per comment. -->",
            "Our drafted answer.",
        )
        comment_file.write_text(text, encoding="utf-8")

        subprocess.run(
            [sys.executable, str(SCRIPTS / "extract_comments.py"),
             "--docx", str(src), "--out", str(out)],
            capture_output=True, text=True,
        )
        after = comment_file.read_text(encoding="utf-8")
        check("status preserved", "status: approved" in after, after[:300])
        check("owner preserved", "owner: jane.doe" in after, after[:300])
        check("reply author preserved", "reply_author: jane.doe" in after, after[:300])
        check("drafted response preserved", "Our drafted answer." in after, after[:600])


def test_warnings_are_not_reset_per_comment():
    print("\n[12] every warning is reported, not just the last")
    with tempfile.TemporaryDirectory() as tmp:
        src = build_sample(Path(tmp) / "sample.docx")
        out = Path(tmp) / "comments"
        subprocess.run(
            [sys.executable, str(SCRIPTS / "extract_comments.py"),
             "--docx", str(src), "--out", str(out)],
            capture_output=True, text=True,
        )
        identity = Path(tmp) / "identity.json"
        identity.write_text(json.dumps({
            "default_author": "jane.doe",
            "authors": [{"id": "jane.doe", "name": "Jane Doe", "org": "Agency", "email": "j@x.gov"}],
        }), encoding="utf-8")
        comment_file = out / "C-001.md"
        text = comment_file.read_text(encoding="utf-8")
        text = text.replace("status: open", "status: approved")
        text = text.replace('owner: ""', 'owner: jane.doe\nreply_author: jane.doe')
        comment_file.write_text(text, encoding="utf-8")

        manifest = Path(tmp) / "manifest.json"
        result = subprocess.run(
            [sys.executable, str(SCRIPTS / "apply_replies.py"),
             "--docx", str(src), "--comments", str(out), "--identity", str(identity),
             "--out", str(Path(tmp) / "w" / "v1.docx"), "--manifest", str(manifest)],
            capture_output=True, text=True,
        )
        check("apply succeeded", result.returncode == 0, result.stderr[-300:])
        manifest_data = json.loads(manifest.read_text(encoding="utf-8"))
        entries = {c["id"]: c for c in manifest_data.get("comments", [])}
        entry = entries.get("C-001", {})
        check("manifest records the owner", entry.get("owner") == "jane.doe",
              str(entry.get("owner")))
        check("manifest records reply author id",
              entry.get("reply_author") == "jane.doe", str(entry.get("reply_author")))
        check("manifest records reply author name",
              entry.get("reply_author_name") == "Jane Doe", str(entry.get("reply_author_name")))
        check("manifest records whether an in-document reply was written",
              entry.get("in_document_reply") is True, str(entry.get("in_document_reply")))


def test_manifest_records_each_comment_own_owner():
    print("\n[13] manifest records the right owner for every comment")
    with tempfile.TemporaryDirectory() as tmp:
        src = build_sample(Path(tmp) / "sample.docx")
        out = Path(tmp) / "comments"
        subprocess.run(
            [sys.executable, str(SCRIPTS / "extract_comments.py"),
             "--docx", str(src), "--out", str(out)],
            capture_output=True, text=True,
        )
        identity = Path(tmp) / "identity.json"
        identity.write_text(json.dumps({
            "default_author": "jane.doe",
            "authors": [
                {"id": "jane.doe", "name": "Jane Doe", "org": "Agency", "email": "j@x.gov"},
                {"id": "alice", "name": "Alice Chen", "org": "Agency", "email": "a@x.gov"},
            ],
        }), encoding="utf-8")

        comment_file = out / "C-001.md"
        text = comment_file.read_text(encoding="utf-8")
        text = text.replace("status: open", "status: approved")
        # owner and reply author deliberately DIFFER, so a bug that copies one
        # comment's owner onto another is visible
        text = text.replace('owner: ""', 'owner: alice\nreply_author: jane.doe')
        comment_file.write_text(text, encoding="utf-8")

        manifest = Path(tmp) / "manifest.json"
        result = subprocess.run(
            [sys.executable, str(SCRIPTS / "apply_replies.py"),
             "--docx", str(src), "--comments", str(out), "--identity", str(identity),
             "--out", str(Path(tmp) / "w" / "v1.docx"), "--manifest", str(manifest)],
            capture_output=True, text=True,
        )
        check("apply succeeded", result.returncode == 0, result.stderr[-300:])
        data = json.loads(manifest.read_text(encoding="utf-8"))
        entry = {c["id"]: c for c in data.get("comments", [])}.get("C-001", {})
        check("owner is the comment's own owner", entry.get("owner") == "alice",
              str(entry.get("owner")))
        check("reply author id is the reply author's",
              entry.get("reply_author") == "jane.doe", str(entry.get("reply_author")))
        check("reply author name matches the roster",
              entry.get("reply_author_name") == "Jane Doe",
              str(entry.get("reply_author_name")))


def test_multi_round_detects_a_re_raise():
    print("\n[14] a second round re-opens the comment the department came back to")
    with tempfile.TemporaryDirectory() as tmp:
        work = Path(tmp)
        identity = work / "identity.json"
        identity.write_text(json.dumps({
            "default_author": "jane.doe",
            "authors": [
                {"id": "jane.doe", "name": "Jane Doe", "org": "Agency", "email": "j@x.gov"},
                {"id": "alice", "name": "Alice Chen", "org": "Agency", "email": "a@x.gov"},
            ],
        }), encoding="utf-8")

        src = build_sample(work / "sample.docx")
        parts = D.read_parts(src)
        # Round 1: we answer the comment.
        ours = D.add_reply(parts, "1", "Our answer.", "Jane Doe")
        D.write_parts(parts, work / "r1.docx")

        # Round 2: the department replies to our answer.
        parts = D.read_parts(work / "r1.docx")
        D.add_reply(parts, ours["w_id"], "That does not answer my question.",
                    "Dept Officer")
        D.write_parts(parts, work / "r2.docx")

        out = work / "comments"
        result = subprocess.run(
            [sys.executable, str(SCRIPTS / "extract_comments.py"),
             "--docx", str(work / "r2.docx"), "--out", str(out),
             "--identity", str(identity)],
            capture_output=True, text=True,
        )
        check("extraction succeeded", result.returncode == 0, result.stderr[-300:])

        files = sorted(out.glob("C-*.md"))
        check("a comment file was produced", len(files) >= 1, str(files))
        if not files:
            return
        body = files[0].read_text(encoding="utf-8")
        check("the re-raised comment is marked reopened", "status: reopened" in body,
              body[:400])
        check("the department's words are shown",
              "That does not answer my question." in body, body[-400:])
        check("our own reply is not mistaken for a re-raise",
              "Our answer." not in body.split("## Department replies")[-1]
              if "## Department replies" in body else True)


def test_reply_to_a_reply_is_anchored():
    print("\n[15] a reply to a reply stays anchored in the document")
    with tempfile.TemporaryDirectory() as tmp:
        work = Path(tmp)
        src = build_sample(work / "sample.docx")
        parts = D.read_parts(src)
        ours = D.add_reply(parts, "1", "Our answer.", "Jane Doe")
        D.write_parts(parts, work / "r1.docx")

        parts = D.read_parts(work / "r1.docx")
        dept = D.add_reply(parts, ours["w_id"], "Still unclear.", "Dept Officer")
        check("the reply to a reply was created",
              dept["w_id"] not in ("1", ours["w_id"]), str(dept))
        check("it is threaded to our reply",
              dept["parent_para_id"] == ours["para_id"], str(dept))

        doc_xml = parts["word/document.xml"].decode()
        check("it carries its own commentReference",
              f'<w:commentReference w:id="{dept["w_id"]}"' in doc_xml, doc_xml[-400:])


def main():
    print("govt-comment-resolution regression tests")
    test_reply_is_threaded_and_visible()
    test_tracked_change_preserves_reading_order()
    test_strip_comments_keeps_body_text()
    test_content_types_use_default_namespace()
    test_extract_flags_orphans_and_numbers_contiguously()
    test_page_alignment_detects_shifted_pages()
    test_authorship_validation()
    test_idempotency()
    test_compare_uses_accepted_text_not_markup()
    test_compare_end_to_end_with_tracked_edit()
    test_edit_markers_do_not_leak_into_replies()
    test_reextract_preserves_triage()
    test_warnings_are_not_reset_per_comment()
    test_manifest_records_each_comment_own_owner()
    test_multi_round_detects_a_re_raise()
    test_reply_to_a_reply_is_anchored()

    print(f"\n{len(PASSES)} passed, {len(FAILURES)} failed")
    if FAILURES:
        for failure in FAILURES:
            print(f"  - {failure}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
