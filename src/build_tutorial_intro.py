#!/usr/bin/env python3
"""Build the tutorial opener, for the longer talk.

Names the three things the tutorial covers. One slide follows per item, in the
order listed here; the order lives in build_deck.py's DECKS.

Output: agentlab_tutorial_intro.pptx
"""

import os

from pptx.enum.text import PP_ALIGN

from slidekit import ROOT, new_deck, blank_slide, title_block, text, BLUE, INK


ITEMS = [
    "Install and run the simple example",
    "Run with Globus Compute",
    "Add Slack integration",
]


def add(prs):
    slide = blank_slide(prs)

    title_block(slide, "Tutorial", "Three things.")

    for i, item in enumerate(ITEMS):
        y = 2.20 + i * 1.05
        text(slide, 0.90, y - 0.06, 0.70, 0.60,
             [(str(i + 1), 34, BLUE, True, False)], align=PP_ALIGN.LEFT)
        text(slide, 1.80, y, 10.90, 0.50,
             [(item, 22, INK, False, False)], align=PP_ALIGN.LEFT)
    return slide


if __name__ == "__main__":
    deck = new_deck()
    add(deck)
    dest = os.path.join(ROOT, "agentlab_tutorial_intro.pptx")
    deck.save(dest)
    print("wrote", dest)
