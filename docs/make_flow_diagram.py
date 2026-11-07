#!/usr/bin/env python3
"""Generate the review-round swimlane diagram.

Writes two files next to this script:

    review-round-flow.svg    standalone SVG, linked from the README
    review-round-flow.html   self-contained page with the same drawing

Both come from the one spec below, so they cannot drift apart.

    python3 docs/make_flow_diagram.py

The output is plain UTF-8 text on purpose: this repository is published through
an API that transfers file contents as text, so anything binary or with stray
bytes would not survive the trip.
"""

from __future__ import annotations

import pathlib
from xml.sax.saxutils import escape

HERE = pathlib.Path(__file__).resolve().parent

# --------------------------------------------------------------------------
# Layout
# --------------------------------------------------------------------------

VIEW_W, VIEW_H = 1320, 576

PAPER = "#f5f5f5"
INK = "#2d3142"
MUTED = "#4f5d75"
ACCENT = "#eb6c36"

LANE_LEFT, LANE_RIGHT = 40, 1280
ACTOR_DIVIDER = 160

# Lane order is chosen so the flow hops as little as possible: down into the
# machinery, back up through the reviewer, out to the department.
LANES = [
    ("DEPARTMENT", 80),
    ("REVIEWER", 184),
    ("SKILL", 288),
    ("COLLEAGUE", 392),
]
LANE_H = 104
NODE_H = 48
NODE_W = 144
COL_X = [176, 360, 544, 728, 912, 1096]

# id, lane index, column index, title, sublabel, focal
NODES = [
    ("n1", 0, 0, "Commented .docx", "incoming file", False),
    ("n2", 2, 1, "Extract comments", "extract_comments.py", False),
    ("n3", 1, 2, "Triage comments", "status · owner · author", False),
    ("n4", 3, 3, "Answer flagged", "needs_input only", False),
    ("n5", 1, 3, "Build change plan", "before any edit", False),
    ("n6", 1, 4, "Approve plan", "human gate", True),
    ("n7", 2, 4, "Apply replies", "apply_replies.py", False),
    ("n8", 2, 5, "Render files", "clean · marked · review", False),
    ("n9", 0, 5, "Receive reply", "email + clean PDF", False),
]

# path, dashed, accent, label text, label x, label y
ARROWS = [
    ("M 248 156 V 332 Q 248 340 256 340 H 358", False, False, None, 0, 0),
    ("M 504 340 H 516 Q 524 340 524 332 V 244 Q 524 236 532 236 H 542", False, False, None, 0, 0),
    ("M 688 236 H 726", False, False, None, 0, 0),
    ("M 616 260 V 290 Q 616 298 624 298 H 752 Q 760 298 760 306 V 418", True, False, "NEEDS INPUT", 800, 353),
    ("M 840 418 V 262", True, False, "ANSWER", 876, 329),
    ("M 872 236 H 910", False, False, None, 0, 0),
    ("M 984 260 V 314", False, True, "APPROVED", 1028, 279),
    ("M 1056 340 H 1094", False, False, None, 0, 0),
    ("M 1168 316 V 158", False, False, "SEND", 1136, 241),
]

# The label masks are sized from the text so the stroke never runs under a glyph.
LABEL_WIDTH = {
    "NEEDS INPUT": 64,
    "ANSWER": 56,
    "APPROVED": 72,
    "SEND": 48,
}

LEGEND = [
    ("box", "Step"),
    ("gate", "Approval gate"),
    ("arrow", "Next step"),
    ("dashed", "Colleague branch"),
    ("accent", "Gated handoff"),
]
LEGEND_X = [140, 220, 356, 496, 688]

TITLE = "Government comment review round"
DESC = (
    "Swimlane diagram of one review round across four lanes: department, "
    "reviewer, skill, and colleague. The department returns a commented Word "
    "document; the skill extracts each comment into a ledger; the reviewer "
    "triages them and routes some to a colleague; a change plan is built and "
    "must be approved before any edit is applied; the skill then writes the "
    "replies with tracked changes, renders the reply files and three PDFs, and "
    "the round closes by delivering the reply back to the department."
)

SANS = "'Geist', system-ui, sans-serif"
MONO = "'Geist Mono', ui-monospace, monospace"


# --------------------------------------------------------------------------
# Rendering
# --------------------------------------------------------------------------


def _node_box(node) -> tuple[int, int]:
    """Return the (x, y) of a node's top-left corner."""
    _, lane, col, _, _, _ = node
    return COL_X[col], LANES[lane][1] + 28


def _node_centre(node) -> tuple[int, int]:
    x, y = _node_box(node)
    return x + NODE_W // 2, y + NODE_H // 2


def render_svg(root_attrs: str = "") -> str:
    """Render the diagram body. Shared by the SVG and the HTML.

    ``root_attrs`` is added to the root element. The standalone file passes an
    intrinsic size so a browser draws it at a known scale; the HTML leaves it
    off because its stylesheet sizes the element itself.
    """
    out: list[str] = []
    add = out.append

    add(f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {VIEW_W} {VIEW_H}"{root_attrs}')
    add('     role="img" aria-labelledby="gcrf-title gcrf-desc">')
    add(f'  <title id="gcrf-title">{escape(TITLE)}</title>')
    add(f'  <desc id="gcrf-desc">{escape(DESC)}</desc>')
    add("  <defs>")
    add(f'    <marker id="arrow" markerWidth="8" markerHeight="6" refX="7" refY="3" orient="auto">'
        f'<polygon points="0 0, 8 3, 0 6" fill="{MUTED}"/></marker>')
    add(f'    <marker id="arrow-accent" markerWidth="8" markerHeight="6" refX="7" refY="3" orient="auto">'
        f'<polygon points="0 0, 8 3, 0 6" fill="{ACCENT}"/></marker>')
    add("  </defs>")
    add(f'  <rect width="100%" height="100%" fill="{PAPER}"/>')

    # Lane dividers
    top = LANES[0][1]
    bottom = LANES[-1][1] + LANE_H
    for i, (_, y) in enumerate(LANES):
        edge = 0.22 if i in (0, len(LANES) - 1) else 0.10
        add(f'  <line x1="{LANE_LEFT}" y1="{y}" x2="{LANE_RIGHT}" y2="{y}" '
            f'stroke="rgba(45,49,66,{edge})" stroke-width="1"/>')
    add(f'  <line x1="{LANE_LEFT}" y1="{bottom}" x2="{LANE_RIGHT}" y2="{bottom}" '
        f'stroke="rgba(45,49,66,0.22)" stroke-width="1"/>')
    add(f'  <line x1="{ACTOR_DIVIDER}" y1="{top}" x2="{ACTOR_DIVIDER}" y2="{bottom}" '
        f'stroke="rgba(45,49,66,0.22)" stroke-width="1"/>')

    # Lane labels
    for name, y in LANES:
        add(f'  <text x="60" y="{y + 56}" fill="{MUTED}" font-size="8" '
            f'font-family="{MONO}" letter-spacing="0.18em">{escape(name)}</text>')

    # Arrows, drawn before the nodes so the boxes sit on top
    for path, dashed, accent, label, lx, ly in ARROWS:
        stroke = ACCENT if accent else MUTED
        width = 1.4 if accent else 1.2
        dash = ' stroke-dasharray="5,4"' if dashed else ""
        marker = "arrow-accent" if accent else "arrow"
        add(f'  <path d="{path}" fill="none" stroke="{stroke}" stroke-width="{width}"'
            f'{dash} marker-end="url(#{marker})"/>')

    # Arrow labels, on an opaque mask so the stroke never runs under a glyph
    for path, dashed, accent, label, lx, ly in ARROWS:
        if not label:
            continue
        w = LABEL_WIDTH[label]
        fill = ACCENT if accent else MUTED
        add(f'  <rect x="{lx - w // 2}" y="{ly - 9}" width="{w}" height="12" rx="2" fill="{PAPER}"/>')
        add(f'  <text x="{lx}" y="{ly}" fill="{fill}" font-size="8" font-family="{MONO}" '
            f'text-anchor="middle" letter-spacing="0.06em">{escape(label)}</text>')

    # Nodes
    for node in NODES:
        _, _, _, title, sub, focal = node
        x, y = _node_box(node)
        cx = x + NODE_W // 2
        if focal:
            fill = f"rgba(235,108,54,0.08)"
            stroke = ACCENT
        else:
            fill = "#ffffff"
            stroke = INK
        add(f'  <rect x="{x}" y="{y}" width="{NODE_W}" height="{NODE_H}" rx="6" '
            f'fill="{fill}" stroke="{stroke}" stroke-width="1"/>')
        add(f'  <text x="{cx}" y="{y + 22}" fill="{INK}" font-size="12" font-weight="600" '
            f'font-family="{SANS}" text-anchor="middle">{escape(title)}</text>')
        add(f'  <text x="{cx}" y="{y + 36}" fill="{MUTED}" font-size="8" '
            f'font-family="{MONO}" text-anchor="middle">{escape(sub)}</text>')

    # Legend, below the diagram
    add(f'  <line x1="{LANE_LEFT}" y1="516" x2="{LANE_RIGHT}" y2="516" '
        f'stroke="rgba(45,49,66,0.10)" stroke-width="0.8"/>')
    add(f'  <text x="40" y="538" fill="{MUTED}" font-size="8" font-family="{MONO}" '
        f'letter-spacing="0.18em">LEGEND</text>')
    for kind, text in LEGEND:
        x = LEGEND_X[LEGEND.index((kind, text))]
        if kind == "box":
            add(f'  <rect x="{x}" y="528" width="16" height="12" rx="2" fill="#ffffff" '
                f'stroke="{INK}" stroke-width="1"/>')
            tx = x + 24
        elif kind == "gate":
            add(f'  <rect x="{x}" y="528" width="16" height="12" rx="2" '
                f'fill="rgba(235,108,54,0.08)" stroke="{ACCENT}" stroke-width="1"/>')
            tx = x + 24
        elif kind == "arrow":
            add(f'  <line x1="{x}" y1="534" x2="{x + 28}" y2="534" stroke="{MUTED}" '
                f'stroke-width="1.2" marker-end="url(#arrow)"/>')
            tx = x + 36
        elif kind == "dashed":
            add(f'  <line x1="{x}" y1="534" x2="{x + 28}" y2="534" stroke="{MUTED}" '
                f'stroke-width="1.2" stroke-dasharray="4,3" marker-end="url(#arrow)"/>')
            tx = x + 36
        else:
            add(f'  <line x1="{x}" y1="534" x2="{x + 28}" y2="534" stroke="{ACCENT}" '
                f'stroke-width="1.4" marker-end="url(#arrow-accent)"/>')
            tx = x + 36
        add(f'  <text x="{tx}" y="538" fill="{MUTED}" font-size="8" '
            f'font-family="{SANS}">{escape(text)}</text>')

    add("</svg>")
    return "\n".join(out)


SVG_HEADER = '<?xml version="1.0" encoding="UTF-8"?>\n'

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{title} · Swimlane</title>
  <link href="https://fonts.googleapis.com/css2?family=Instrument+Serif:ital@0;1&family=Geist:wght@400;500;600&family=Geist+Mono:wght@400;500;600&display=swap" rel="stylesheet">
  <style>
    *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
    :root {{
      --color-paper:  {paper};
      --color-ink:    {ink};
      --color-muted:  {muted};
      --color-accent: {accent};
      --font-sans:    'Geist', system-ui, sans-serif;
      --font-serif:   'Instrument Serif', serif;
      --font-mono:    'Geist Mono', ui-monospace, monospace;
    }}
    body {{
      font-family: var(--font-sans);
      background: var(--color-paper);
      color: var(--color-ink);
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 3rem 2rem;
    }}
    .frame {{ max-width: 1320px; width: 100%; }}
    .eyebrow {{
      font-family: var(--font-mono);
      font-size: 0.66rem;
      font-weight: 500;
      letter-spacing: 0.18em;
      text-transform: uppercase;
      color: var(--color-muted);
      margin-bottom: 0.5rem;
    }}
    h1 {{
      font-family: var(--font-serif);
      font-size: clamp(1.5rem, 2.4vw + 0.75rem, 2rem);
      font-weight: 400;
      letter-spacing: -0.02em;
      line-height: 1.15;
      margin-bottom: 1.5rem;
    }}
    svg {{ width: 100%; min-width: 1080px; display: block; }}
  </style>
</head>
<body>
  <div class="frame">
    <p class="eyebrow">Swimlane · Diagram Design</p>
    <h1>{title}</h1>
{svg}
  </div>
</body>
</html>
"""


def main() -> None:
    svg = render_svg()
    (HERE / "review-round-flow.svg").write_text(
        SVG_HEADER + render_svg(root_attrs=f' width="{VIEW_W}" height="{VIEW_H}"') + "\n",
        encoding="utf-8",
    )
    html = HTML_TEMPLATE.format(
        title=escape(TITLE),
        paper=PAPER,
        ink=INK,
        muted=MUTED,
        accent=ACCENT,
        svg=svg,
    )
    (HERE / "review-round-flow.html").write_text(html, encoding="utf-8")
    print("wrote review-round-flow.svg and review-round-flow.html")


if __name__ == "__main__":
    main()
