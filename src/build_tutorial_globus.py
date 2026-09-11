#!/usr/bin/env python3
"""Build tutorial slide 2a — set the Globus Compute endpoint up.

Host names and commands are the user's, verbatim. Starting the endpoint and
running the campaign are the next slide, build_tutorial_globus_run.py.

Output: agentlab_tutorial_globus.pptx
"""

import os

from pptx.enum.text import PP_ALIGN

from slidekit import (
    ROOT,
    new_deck, blank_slide, title_block, box, text,
    MONO, INK, FAINT, BLUE, SLATE, GREY_BG, HAIRLINE,
)


HEAD_H, ITEM_H = 0.36, 0.32
COL_W = 5.85
LX, RX = 0.62, 6.86

NODE = [
    "no scheduler / no queue",
    "You will need to ssh to the node — check it is not too busy",
    "FQDN (e.g. compute-386-03.cels.anl.gov)",
]

REMOTE = [
    "python3 -m venv ~/gc_env",
    "source ~/gc_env/bin/activate",
    "pip install globus-compute-endpoint",
    "globus-compute-endpoint configure agentlab",
]

COPY = ["scp systems/endpoints/no_scheduler/user_config_template.yaml.j2 "
        "compute-386-03.cels.anl.gov:.globus_compute/agentlab/"]


def _heading(slide, x, y, head, width=COL_W):
    """Give a heading the second line it needs when it wraps, or it lands on
    the first bullet under it."""
    h = HEAD_H if len(head) <= 62 else HEAD_H + 0.26
    text(slide, x, y, width, h,
         [(head, 12.5, SLATE, True, False)], align=PP_ALIGN.LEFT)
    return y + h


def _bullet(slide, x, y, item):
    text(slide, x + 0.10, y, 0.20, 0.28,
         [("•", 11, BLUE, True, False)], align=PP_ALIGN.LEFT)
    text(slide, x + 0.34, y, COL_W - 0.34, 0.30,
         [(item, 11, INK, False, False)], align=PP_ALIGN.LEFT)
    return y + ITEM_H


def _commands(slide, x, y, where, lines, width=COL_W, size=11.5):
    """A labelled command block. The label says where the commands are run."""
    text(slide, x, y, width, 0.28,
         [(where, 11, SLATE, True, False)], align=PP_ALIGN.LEFT)
    y += 0.32
    h = 0.38 + 0.36 * len(lines)
    box(slide, x, y, width, h, fill=GREY_BG, stroke=HAIRLINE, line_w=1.0,
        radius=0.08)
    for i, line in enumerate(lines):
        text(slide, x + 0.28, y + 0.17 + i * 0.36, width - 0.56, 0.30,
             [(line, size, INK, False, False)], align=PP_ALIGN.LEFT, font=MONO)
    return y + h


def add(prs):
    slide = blank_slide(prs)

    title_block(
        slide,
        "2(a)   Run with Globus Compute",
        "Set the endpoint up on the remote node.",
    )

    y = _heading(slide, LX, 1.58, "Use a remote endpoint on a GCE node")
    for item in NODE:
        y = _bullet(slide, LX, y, item)

    _commands(slide, RX, 1.58, "On compute-386-03:", REMOTE)

    y = _commands(slide, LX, 4.10, "Laptop — copy the endpoint:", COPY,
                  width=12.10, size=10)
    text(slide, LX, y + 0.18, 12.10, 0.34,
         [("The agent can do this for you, once it has the FQDN.",
           11, FAINT, False, True)], align=PP_ALIGN.LEFT)
    return slide


if __name__ == "__main__":
    deck = new_deck()
    add(deck)
    dest = os.path.join(ROOT, "agentlab_tutorial_globus.pptx")
    deck.save(dest)
    print("wrote", dest)
