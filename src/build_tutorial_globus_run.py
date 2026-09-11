#!/usr/bin/env python3
"""Build tutorial slide 2b — start the endpoint and run the campaign.

Commands and paths are the user's, verbatim. Setting the endpoint up is the
previous slide, build_tutorial_globus.py.

Output: agentlab_tutorial_globus_run.pptx
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

REMOTE_RUN = [
    "globus-compute-endpoint start agentlab --detach",
    "globus-compute-endpoint list",
]
GCE_JSON = [
    '{"endpoint": "<uuid from list>",',
    ' "work_dir": "/home/<you>/agentlab-work"}',
]
LOCAL_RUN = ["WATCH=true AGENT_MODEL=haiku ./run_remote.sh"]


def _heading(slide, x, y, head):
    """Give a heading the second line it needs when it wraps, or it lands on
    the first bullet under it."""
    h = HEAD_H if len(head) <= 62 else HEAD_H + 0.26
    text(slide, x, y, COL_W, h,
         [(head, 12.5, SLATE, True, False)], align=PP_ALIGN.LEFT)
    return y + h


def _bullet(slide, x, y, item):
    text(slide, x + 0.10, y, 0.20, 0.28,
         [("•", 11, BLUE, True, False)], align=PP_ALIGN.LEFT)
    text(slide, x + 0.34, y, COL_W - 0.34, 0.30,
         [(item, 11, INK, False, False)], align=PP_ALIGN.LEFT)
    return y + ITEM_H


def _commands(slide, x, y, where, lines):
    """A labelled command block. The label says where the commands are run."""
    text(slide, x, y, COL_W, 0.28,
         [(where, 11, SLATE, True, False)], align=PP_ALIGN.LEFT)
    y += 0.32
    h = 0.38 + 0.36 * len(lines)
    box(slide, x, y, COL_W, h, fill=GREY_BG, stroke=HAIRLINE, line_w=1.0,
        radius=0.08)
    for i, line in enumerate(lines):
        text(slide, x + 0.28, y + 0.17 + i * 0.36, COL_W - 0.56, 0.30,
             [(line, 11.5, INK, False, False)], align=PP_ALIGN.LEFT, font=MONO)
    return y + h


def add(prs):
    slide = blank_slide(prs)

    title_block(
        slide,
        "2(b)   Run with Globus Compute",
        "Start the endpoint, then start the campaign.",
    )

    y = _heading(slide, LX, 1.58,
                 "You start it and authenticate with Globus Compute "
                 "(interactive)")
    y = _commands(slide, LX, y, "On compute-386-03:", REMOTE_RUN)

    text(slide, LX + 0.10, y + 0.16, COL_W, 0.30,
         [("list gives you the endpoint UUID", 11, INK, False, False)],
         align=PP_ALIGN.LEFT)
    text(slide, LX, y + 0.62, COL_W, 0.50,
         [("Once the detached endpoint is running, the ssh connection is no "
           "longer needed.", 11, FAINT, False, True)], align=PP_ALIGN.LEFT)

    y = _heading(slide, RX, 1.58, "Then on your own machine")
    y = _bullet(slide, RX, y,
                "Agent uses the UUID and puts your personal config under "
                "users/<user_id>")
    y = _commands(slide, RX, y + 0.10, "On the laptop, users/<you>/gce.json:",
                  GCE_JSON)
    text(slide, RX, y + 0.16, COL_W, 0.52,
         [("Note the work_dir field is required for remote machines, but not "
           "used in this case as no files created.", 11, FAINT, False, True)],
         align=PP_ALIGN.LEFT)

    y = _commands(slide, RX, y + 0.76, "Laptop:", LOCAL_RUN)

    text(slide, RX, y + 0.20, COL_W, 0.34,
         [("The quick test does not produce files on the remote machine.",
           11, FAINT, False, True)], align=PP_ALIGN.LEFT)
    return slide


if __name__ == "__main__":
    deck = new_deck()
    add(deck)
    dest = os.path.join(ROOT, "agentlab_tutorial_globus_run.pptx")
    deck.save(dest)
    print("wrote", dest)
