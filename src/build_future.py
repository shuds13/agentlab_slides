#!/usr/bin/env python3
"""Build the future-work slide, for the longer talk.

What is planned rather than built. Nothing here is claimed as working
elsewhere in the deck — IRI, American Science Cloud and Genesis appear on this
slide only.

Output: agentlab_future.pptx
"""

import os

from pptx.enum.text import PP_ALIGN

from slidekit import ROOT, new_deck, blank_slide, title_block, text, BLUE, INK, SLATE


HEAD_H, ITEM_H, GROUP_GAP = 0.40, 0.36, 0.44
WIDTH = 12.10
X = 0.62

GROUPS = [
    ("IRI support", [
        "alternative to Globus Compute for submission",
        "alternative tools for reading and writing files, and for querying "
        "system and job status",
    ]),
    ("American Science Cloud / Genesis", [
        "Tools for services · query dashboards",
        "Links to skills · catalogues · knowledge bases",
    ]),
    ("Integrate with other front-ends", [
        "AI workbenches · Experiment trackers",
    ]),
]


def add(prs):
    slide = blank_slide(prs)

    title_block(slide, "Future", "")

    y = 1.62
    for head, items in GROUPS:
        text(slide, X, y, WIDTH, HEAD_H,
             [(head, 13.5, SLATE, True, False)], align=PP_ALIGN.LEFT)
        y += HEAD_H
        for item in items:
            text(slide, X + 0.10, y, 0.20, 0.30,
                 [("•", 12, BLUE, True, False)], align=PP_ALIGN.LEFT)
            text(slide, X + 0.36, y, WIDTH - 0.36, 0.32,
                 [(item, 12, INK, False, False)], align=PP_ALIGN.LEFT)
            y += ITEM_H
        y += GROUP_GAP
    return slide


if __name__ == "__main__":
    deck = new_deck()
    add(deck)
    dest = os.path.join(ROOT, "agentlab_future.pptx")
    deck.save(dest)
    print("wrote", dest)
