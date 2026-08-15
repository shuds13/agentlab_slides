#!/usr/bin/env python3
"""Build the AgentLab workspace slide.

The campaign shown is invented. The file list and what each file holds are the
real ones, from framework/SYSTEM.md and framework/tools.py:

  results.jsonl      SYSTEM.md:42  one JSON object per line, appended as each
                                   result lands
  claims.jsonl       tools.py:162  agents claim work here under a file lock so
                                   two never run the same piece
  LOGBOOK.md         SYSTEM.md:44  terse memory across runs, append-only
  ANNOUNCEMENTS.md                 the board a person or Slack posts to mid-run
  logs/                            the agent's own run log, round by round
                                   ("===== ROUND 1 =====", then its commentary)
  runs/                            the prompts each run started from

The journal page is built by journal_excerpt/ at letter size, then shown small,
so the type sits at the proportions of a real document. Rebuild it with:

    cd journal_excerpt && python3 figure.py && pdflatex excerpt.tex \
        && pdftoppm -r 200 -png -f 1 -l 1 -singlefile excerpt.pdf excerpt

Output: agentlab_workspace.pptx
"""

import math
import os

from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

from slidekit import (
    arc_arrow, round_picture,
    new_deck, blank_slide, title_block, box, label_box, text, arrow,
    MONO, MUTED, FAINT, BLUE, BLUE_BG, SLATE, GREY_BG, HAIRLINE, ARROW, WHITE,
)

HERE = os.path.dirname(os.path.abspath(__file__))
PAGE = os.path.join(HERE, "journal_excerpt", "excerpt.png")
PAGE_ASPECT = 1700 / 2200        # letter

def add(prs):
    slide = blank_slide(prs)

    title_block(
        slide,
        "The workspace",
        "One directory per campaign. Everything the agent produces lands in it.",
    )

    # ---------------------------------------------------------- campaign agent
    CX = 2.20                                   # shared centre line for the column

    agent = box(slide, 0.75, 1.48, 2.90, 0.95, fill=BLUE_BG, stroke=BLUE, line_w=1.75)
    label_box(agent, [("Campaign agent", 14, BLUE, True, False),
                      ("Claude Agent SDK", 9, MUTED, False, False)])

    arrow(slide, 3.65, 1.955, 4.35, 1.955)
    text(slide, 3.58, 1.63, 0.84, 0.28,
         [("writes", 9.5, ARROW, False, False)])

    # -------------------------------------------------------------- the cycle
    # A straight stack with a return arrow up the left-hand side.
    STEPS = ["Hypothesize", "Run tasks", "Interpret", "Goal met?"]
    PILL_X, PILL_W, PILL_H, GAP = 1.00, 2.30, 0.42, 0.18
    TOP = 2.85

    arrow(slide, CX, 2.43, CX, TOP, color=BLUE, width=1.4)

    centres = []
    for i, step in enumerate(STEPS):
        y = TOP + i * (PILL_H + GAP)
        centres.append(y + PILL_H / 2)
        p = box(slide, PILL_X, y, PILL_W, PILL_H,
                fill=WHITE, stroke=BLUE, line_w=1.4, radius=0.5)
        label_box(p, [(step, 11, BLUE, True, False)])
        if i:
            arrow(slide, CX, y - GAP, CX, y, color=BLUE, width=1.4)

    BOTTOM = TOP + len(STEPS) * PILL_H + (len(STEPS) - 1) * GAP

    # back round for another cycle
    LOOP_X = 0.55
    arrow(slide, PILL_X, centres[-1], LOOP_X, centres[-1], color=BLUE, width=1.6,
          head=False)
    arrow(slide, LOOP_X, centres[-1], LOOP_X, centres[0], color=BLUE, width=1.6,
          head=False)
    arrow(slide, LOOP_X, centres[0], PILL_X, centres[0], color=BLUE, width=1.6)

    # or out, when the goal check says so
    arrow(slide, CX, BOTTOM, CX, BOTTOM + 0.38, color=BLUE, width=1.75)
    text(slide, PILL_X, BOTTOM + 0.40, PILL_W, 0.28,
         [("stop, write up", 9.5, BLUE, True, False)])

    text(slide, 0.35, BOTTOM + 0.86, 3.85, 0.50,
         [("Repeats until the goal is met, the leads run out,", 9.5, MUTED, False, True),
          ("or the run meets one of its budget constraints.", 9.5, MUTED, False, True)])

    # -------------------------------------------------------------- workspace
    WS_X, WS_Y, WS_W, WS_H = 4.35, 1.42, 8.45, 4.78
    box(slide, WS_X, WS_Y, WS_W, WS_H, fill=None, stroke=HAIRLINE, line_w=1.0, dash=True)
    text(slide, WS_X + 0.23, WS_Y + 0.13, 5.00, 0.30,
         [("workspace/collective-tuning/", 12, SLATE, True, False)],
         align=PP_ALIGN.LEFT, font=MONO)

    # the journal, shown as a page rather than described
    THUMB_H = 3.60
    THUMB_W = THUMB_H * PAGE_ASPECT
    THUMB_X, THUMB_Y = 4.72, 2.05
    pic = slide.shapes.add_picture(PAGE, Inches(THUMB_X), Inches(THUMB_Y),
                                   Inches(THUMB_W), Inches(THUMB_H))
    pic.line.color.rgb = HAIRLINE
    pic.line.width = Pt(1.0)
    round_picture(pic)
    text(slide, THUMB_X, THUMB_Y + THUMB_H + 0.08, THUMB_W, 0.26,
         [("JOURNAL.pdf", 10.5, SLATE, True, False)], font=MONO)

    # the rest of what a campaign leaves behind
    FILES = [
        ("results.jsonl",     "one line per result, as it lands"),
        ("claims.jsonl",      "which agent is running what"),
        ("LOGBOOK.md",        "terse memory across runs"),
        ("ANNOUNCEMENTS.md",  "messages in, mid-run"),
        ("logs/",             "the agent's own log, round by round"),
        ("runs/",             "the prompt each run started from"),
    ]
    CH_X, CH_W, CH_H, CH_GAP = 7.95, 4.55, 0.46, 0.15
    for i, (name, what) in enumerate(FILES):
        y = 2.05 + i * (CH_H + CH_GAP)
        box(slide, CH_X, y, CH_W, CH_H, fill=GREY_BG, stroke=HAIRLINE,
            line_w=1.0, radius=0.12)
        text(slide, CH_X + 0.14, y + 0.10, 1.75, 0.28,
             [(name, 10.5, SLATE, True, False)], align=PP_ALIGN.LEFT, font=MONO)
        text(slide, CH_X + 1.92, y + 0.12, 2.50, 0.28,
             [(what, 9.5, MUTED, False, False)], align=PP_ALIGN.LEFT)

    # ------------------------------------------------------------- footnote
    text(slide, WS_X, 6.32, WS_W, 0.32,
         [("Several agents can share one workspace — claims.jsonl is how they avoid "
           "running the same work twice.", 9.5, FAINT, False, True)],
         align=PP_ALIGN.LEFT)
    return slide


if __name__ == "__main__":
    deck = new_deck()
    add(deck)
    dest = os.path.join(HERE, "agentlab_workspace.pptx")
    deck.save(dest)
    print("wrote", dest)
