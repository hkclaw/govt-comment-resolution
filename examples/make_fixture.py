#!/usr/bin/env python3
"""Build a sample Word document that carries reviewer comments.

The fixture exercises the cases the skill has to survive:

- comments spread over several pages
- comments from three different reviewers
- a reply inside a comment thread
- an orphaned comment (defined, but anchored to no passage)

Usage:
  python3 examples/make_fixture.py examples/sample-incoming.docx
"""

import sys
import zipfile
from pathlib import Path

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
W16CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
W16CEX = "http://schemas.microsoft.com/office/word/2018/wordml/cex"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
RELS = "http://schemas.openxmlformats.org/package/2006/relationships"

PAGE_BREAK = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

# w:id, author, paraId, comment text, body text
COMMENTS = [
    ("1", "R. Alvarez (Planning Dept)", "1A2B3C01",
     "Please justify the 4.2% growth figure in paragraph 1.",
     "The regional allocation is projected to grow by 4.2% over the coming year."),
    ("2", "R. Alvarez (Planning Dept)", "1A2B3C02",
     "This citation is out of date. Use the 2025 edition.",
     "The methodology follows the established national framework for assessment."),
    ("3", "M. Okafor (Legal)", "1A2B3C03",
     "Clause 7 needs a data retention period.",
     "Personal data shall be processed in accordance with applicable regulation."),
    ("4", "M. Okafor (Legal)", "1A2B3C04",
     "What is the total cost of this phase?",
     "Phase one covers the three priority regions identified in the annex."),
    ("5", "S. Lindqvist (Finance)", "1A2B3C05",
     "Figures do not match the budget appendix.",
     "Projected expenditure for the period is summarised in the table below."),
    ("6", "S. Lindqvist (Finance)", "1A2B3C06",
     "Missing the 2024 baseline column.",
     "A breakdown of expenditure by category is provided in the appendix."),
]


def paragraph(text, comment_id=None, para_id=None):
    pid = f' w14:paraId="{para_id}"' if para_id else ""
    if comment_id is None:
        return f'<w:p{pid}><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
    return (
        f'<w:p{pid}><w:commentRangeStart w:id="{comment_id}"/>'
        f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'
        f'<w:commentRangeEnd w:id="{comment_id}"/>'
        f'<w:r><w:commentReference w:id="{comment_id}"/></w:r></w:p>'
    )


def build(out_path):
    body = [
        paragraph("Annual Planning Document", para_id="0A0A0001"),
        paragraph("Prepared for the Regional Development Board", para_id="0A0A0002"),
        PAGE_BREAK,
        paragraph("1. Regional Context", para_id="0A0A0003"),
        paragraph(COMMENTS[0][4], comment_id="1", para_id="0A0A0004"),
        paragraph(COMMENTS[1][4], comment_id="2", para_id="0A0A0005"),
        PAGE_BREAK,
        paragraph("2. Legal Basis", para_id="0A0A0006"),
        paragraph(COMMENTS[2][4], comment_id="3", para_id="0A0A0007"),
        paragraph(COMMENTS[3][4], comment_id="4", para_id="0A0A0008"),
        PAGE_BREAK,
        paragraph("3. Financial Summary", para_id="0A0A0009"),
        paragraph(COMMENTS[4][4], comment_id="5", para_id="0A0A000A"),
        paragraph(COMMENTS[5][4], comment_id="6", para_id="0A0A000B"),
        PAGE_BREAK,
        paragraph("4. Conclusion and Next Steps", para_id="0A0A000C"),
        paragraph(
            "The Board is asked to endorse this plan and authorise the "
            "commencement of Phase one.", para_id="0A0A000D"),
        paragraph("Appendix A - Regional Breakdown", para_id="0A0A000E"),
        paragraph("End of document.", para_id="0A0A000F"),
    ]

    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W}" xmlns:r="{R}" xmlns:w14="{W14}"><w:body>'
        + "".join(body)
        + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
          '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>'
          '</w:sectPr></w:body></w:document>'
    )

    comment_blocks = []
    for cid, author, para_id, text, _ in COMMENTS:
        initials = author.split(" ")[0][:1] + author.split(" ")[-1][:1]
        comment_blocks.append(
            f'<w:comment w:id="{cid}" w:author="{author}" w:initials="{initials}" '
            f'w:date="2026-11-06T09:00:00Z">'
            f'<w:p w14:paraId="{para_id}">'
            f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p></w:comment>'
        )
    # A follow-up inside an existing thread.
    comment_blocks.append(
        '<w:comment w:id="7" w:author="M. Okafor (Legal)" w:initials="MO" '
        'w:date="2026-11-06T11:30:00Z">'
        '<w:p w14:paraId="1A2B3C07"><w:r><w:t xml:space="preserve">'
        'Adding: the retention period should be stated explicitly.'
        '</w:t></w:r></w:p></w:comment>'
    )
    # Orphaned: defined, but anchored to no passage.
    comment_blocks.append(
        '<w:comment w:id="8" w:author="R. Alvarez (Planning Dept)" w:initials="RA" '
        'w:date="2026-11-06T09:05:00Z">'
        '<w:p w14:paraId="1A2B3C08"><w:r><w:t xml:space="preserve">'
        'General note: please proofread throughout.'
        '</w:t></w:r></w:p></w:comment>'
    )

    comments_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:comments xmlns:w="{W}" xmlns:w14="{W14}">'
        + "".join(comment_blocks) + "</w:comments>"
    )

    extended = []
    for cid, author, para_id, text, _ in COMMENTS:
        extended.append(f'<w15:commentEx w15:paraId="{para_id}" w15:done="0"/>')
    extended.append(
        '<w15:commentEx w15:paraId="1A2B3C07" w15:paraIdParent="1A2B3C03" w15:done="0"/>')
    extended.append('<w15:commentEx w15:paraId="1A2B3C08" w15:done="0"/>')
    comments_extended_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w15:commentsEx xmlns:w15="{W15}">' + "".join(extended) + "</w15:commentsEx>"
    )

    ids = []
    for cid, author, para_id, text, _ in COMMENTS:
        ids.append(
            f'<w16cid:commentId w16cid:paraId="{para_id}" '
            f'w16cid:durableId="{cid}0A0B0C{para_id[-2:]}"/>')
    ids.append('<w16cid:commentId w16cid:paraId="1A2B3C07" w16cid:durableId="70A0B0C07"/>')
    ids.append('<w16cid:commentId w16cid:paraId="1A2B3C08" w16cid:durableId="80A0B0C08"/>')
    comments_ids_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w16cid:commentsIds xmlns:w16cid="{W16CID}">' + "".join(ids) + "</w16cid:commentsIds>"
    )

    extensible = []
    for cid, author, para_id, text, _ in COMMENTS:
        extensible.append(
            f'<w16cex:commentExtensible w16cex:durableId="{cid}0A0B0C{para_id[-2:]}" '
            'w16cex:dateUtc="2026-11-06T09:00:00Z"/>')
    extensible.append(
        '<w16cex:commentExtensible w16cex:durableId="70A0B0C07" '
        'w16cex:dateUtc="2026-11-06T11:30:00Z"/>')
    extensible.append(
        '<w16cex:commentExtensible w16cex:durableId="80A0B0C08" '
        'w16cex:dateUtc="2026-11-06T09:05:00Z"/>')
    comments_extensible_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w16cex:commentsExtensible xmlns:w16cex="{W16CEX}">'
        + "".join(extensible) + "</w16cex:commentsExtensible>"
    )

    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Types xmlns="{CT}">'
        '<Default Extension="rels" '
        'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.'
        'openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.'
        'openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '<Override PartName="/word/comments.xml" ContentType="application/vnd.'
        'openxmlformats-officedocument.wordprocessingml.comments+xml"/>'
        '<Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.'
        'openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>'
        '<Override PartName="/word/commentsIds.xml" ContentType="application/vnd.'
        'openxmlformats-officedocument.wordprocessingml.commentsIds+xml"/>'
        '<Override PartName="/word/commentsExtensible.xml" ContentType="application/vnd.'
        'openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml"/>'
        '</Types>'
    )

    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{RELS}">'
        f'<Relationship Id="rId1" Type="{R}/officeDocument" Target="word/document.xml"/>'
        '</Relationships>'
    )

    doc_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{RELS}">'
        f'<Relationship Id="rId1" Type="{R}/styles" Target="styles.xml"/>'
        f'<Relationship Id="rId20" Type="{R}/comments" Target="comments.xml"/>'
        '<Relationship Id="rId21" Type="http://schemas.microsoft.com/office/2011/'
        'relationships/commentsExtended" Target="commentsExtended.xml"/>'
        '<Relationship Id="rId22" Type="http://schemas.microsoft.com/office/2016/09/'
        'relationships/commentsIds" Target="commentsIds.xml"/>'
        '<Relationship Id="rId23" Type="http://schemas.microsoft.com/office/2018/08/'
        'relationships/commentsExtensible" Target="commentsExtensible.xml"/>'
        '</Relationships>'
    )

    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles xmlns:w="{W}">'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
        '<w:name w:val="Normal"/></w:style>'
        '<w:style w:type="paragraph" w:styleId="CommentText">'
        '<w:name w:val="annotation text"/></w:style>'
        '<w:style w:type="character" w:styleId="CommentReference">'
        '<w:name w:val="annotation reference"/></w:style>'
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

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    print(f"built {out_path}")


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else "examples/sample-incoming.docx"
    build(target)
