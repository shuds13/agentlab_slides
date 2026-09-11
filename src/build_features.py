#!/usr/bin/env python3
"""Build the other-features and example-campaigns slide, for the longer talk.

Output: agentlab_features.pptx
"""

import os

from pptx.enum.text import PP_ALIGN

from slidekit import ROOT, new_deck, blank_slide, title_block, text, BLUE, INK, SLATE


HEAD_H, ITEM_H = 0.40, 0.44
COL_W = 5.85
LX, RX = 0.62, 6.86

COLUMNS = [
    (LX, "Other features", [
        "Globus transfer tools: Use to transfer files, within allowed dirs.",
        "Critic agent: Built in; the main agent cannot avoid running it.",
        "Efficiency reviewer: Run to review your workflow for inefficiencies.",
        "Reader Subagent: Used to read large files and summarize to main agent",
    ]),
    (RX, "Example campaigns", [
        "Direct Search and optimization",
        "Layer above ensemble tools (e.g. libEnsemble as a task)",
        "Running & optimizing benchmarks — keeping figures & skills up-to-date.",
    ]),
]


def add(prs):
    slide = blank_slide(prs)

    title_block(slide, "Other features and example campaigns", "")

    for x, head, items in COLUMNS:
        text(slide, x, 1.62, COL_W, HEAD_H,
             [(head, 13.5, SLATE, True, False)], align=PP_ALIGN.LEFT)
        y = 1.62 + HEAD_H
        for item in items:
            text(slide, x + 0.10, y, 0.20, 0.30,
                 [("•", 11, BLUE, True, False)], align=PP_ALIGN.LEFT)
            text(slide, x + 0.36, y, COL_W - 0.36, 0.34,
                 [(item, 11, INK, False, False)], align=PP_ALIGN.LEFT)
            y += ITEM_H
    return slide


if __name__ == "__main__":
    deck = new_deck()
    add(deck)
    dest = os.path.join(ROOT, "agentlab_features.pptx")
    deck.save(dest)
    print("wrote", dest)
