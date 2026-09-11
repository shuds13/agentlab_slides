#!/usr/bin/env python3
"""Build the alternative title slide, for the longer talk.

Same as the short deck's title slide, plus the full author list with
affiliations and the ANL and FASTMath logos.

Affiliations: Argonne National Laboratory for all four; Magesh Rajasekaran is
also Louisiana State University.

Output: agentlab_title_alt.pptx
"""

import os

from pptx.util import Inches
from pptx.enum.text import PP_ALIGN

from slidekit import (
    ROOT, IMAGES,
    new_deck, blank_slide, box, text, MONO, INK, MUTED, BLUE, SLATE,
)

REPO = "github.com/shuds13/AgentLab"

LOGOS = [("anl.png", 366 / 137), ("fastmath_logo.png", 308 / 90)]
LOGO_H = 0.45
LOGO_Y = 6.02
LOGO_GAP = 0.45


def add(prs):
    slide = blank_slide(prs)

    text(slide, 1.0, 2.05, 11.33, 1.10,
         [("AgentLab", 60, INK, True, False)], align=PP_ALIGN.LEFT)
    text(slide, 1.06, 3.18, 11.33, 0.50,
         [("A research lab run by agents", 22, BLUE, False, True)],
         align=PP_ALIGN.LEFT)
    box(slide, 1.08, 3.92, 2.10, 0.045, fill=BLUE, stroke=BLUE, line_w=0.5,
        radius=None)

    text(slide, 1.06, 4.25, 11.33, 0.34,
         [("Stephen Hudson¹ · Magesh Rajasekaran¹² · John-Luke Navarro¹ · "
           "Jeff Larson¹", 15, SLATE, False, False)], align=PP_ALIGN.LEFT)
    text(slide, 1.06, 4.66, 11.33, 0.30,
         [("¹ Argonne National Laboratory     ² Louisiana State University",
           11, MUTED, False, False)], align=PP_ALIGN.LEFT)

    text(slide, 1.06, 5.18, 11.33, 0.30,
         [(REPO, 12, MUTED, False, False)], align=PP_ALIGN.LEFT, font=MONO)

    x = 1.06
    for name, aspect in LOGOS:
        w = LOGO_H * aspect
        slide.shapes.add_picture(os.path.join(IMAGES, name), Inches(x),
                                 Inches(LOGO_Y), Inches(w), Inches(LOGO_H))
        x += w + LOGO_GAP
    return slide


if __name__ == "__main__":
    deck = new_deck()
    add(deck)
    dest = os.path.join(ROOT, "agentlab_title_alt.pptx")
    deck.save(dest)
    print("wrote", dest)
