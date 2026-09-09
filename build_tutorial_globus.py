#!/usr/bin/env python3
"""Build tutorial slide 2 — run with Globus Compute.

Endpoint name, paths and commands are the user's, verbatim. Each command block
is labelled with the machine it runs on: the first two are on the remote node,
the last is on your own machine.

Output: agentlab_tutorial_globus.pptx
"""

import os

from pptx.enum.text import PP_ALIGN

from slidekit import (
    new_deck, blank_slide, title_block, box, text,
    MONO, INK, FAINT, BLUE, SLATE, GREY_BG, HAIRLINE,
)

HERE = os.path.dirname(os.path.abspath(__file__))

HEAD_H, ITEM_H, GROUP_GAP = 0.36, 0.32, 0.30
TOP = 1.62
COL_W = 5.85
RX = 6.86

# (text, is_command) — commands are set in the mono face
LEFT = [
    ("Use a remote endpoint on a GCE node", [
        ("no scheduler / no queue", False),
        ("You will need to ssh to the node — check it is not too busy", False),
        ("ssh compute-386-03.cels.anl", True),
    ]),
    ("Give your agent full domain name (FQDN) and ask to set up endpoint", [
        ("FQDN (e.g. compute-386-03.cels.anl)", False),
        ("Create a virtual environment on the remote", False),
        ("Activate that environment", False),
        ("pip install globus-compute-endpoint", True),
        ("Move a templated endpoint into place", False),
    ]),
]

REMOTE_RUN = [
    "globus-compute-endpoint start agentlab --detach",
    "globus-compute-endpoint list",
]
LOCAL_RUN = ["WATCH=true AGENT_MODEL=sonnet ./run_remote.sh"]


def _heading(slide, x, y, head):
    """Headings wrap at roughly 62 characters in this column; give a wrapped
    one the second line it needs, or it lands on the first bullet."""
    h = HEAD_H if len(head) <= 62 else HEAD_H + 0.26
    text(slide, x, y, COL_W, h,
         [(head, 12.5, SLATE, True, False)], align=PP_ALIGN.LEFT)
    return y + h


def _bullet(slide, x, y, item, command=False):
    text(slide, x + 0.10, y, 0.20, 0.28,
         [("•", 11, BLUE, True, False)], align=PP_ALIGN.LEFT)
    text(slide, x + 0.34, y, COL_W - 0.34, 0.30,
         [(item, 11, INK, False, False)], align=PP_ALIGN.LEFT,
         font=MONO if command else None)


def _commands(slide, x, y, where, lines):
    """A labelled command block. The label says which machine it runs on."""
    text(slide, x, y, COL_W, 0.26,
         [(where, 10, FAINT, False, True)], align=PP_ALIGN.LEFT)
    y += 0.28
    h = 0.42 + 0.40 * len(lines)
    box(slide, x, y, COL_W, h, fill=GREY_BG, stroke=HAIRLINE, line_w=1.0,
        radius=0.08)
    for i, line in enumerate(lines):
        text(slide, x + 0.30, y + 0.20 + i * 0.40, COL_W - 0.60, 0.32,
             [(line, 12.5, INK, False, False)], align=PP_ALIGN.LEFT, font=MONO)
    return y + h


def add(prs):
    slide = blank_slide(prs)

    title_block(
        slide,
        "2   Run with Globus Compute",
        "The agent sets the endpoint up; you start it and sign in.",
    )

    y = TOP
    for head, items in LEFT:
        y = _heading(slide, 0.62, y, head)
        for item, command in items:
            _bullet(slide, 0.62, y, item, command)
            y += ITEM_H
        y += GROUP_GAP

    text(slide, 0.62, y + 0.10, COL_W, 0.34,
         [("Once the detached endpoint is running, the ssh connection is no "
           "longer needed.", 11, FAINT, False, True)], align=PP_ALIGN.LEFT)

    y = _heading(slide, RX, TOP,
                 "You start it and authenticate with Globus Compute "
                 "(interactive)")
    y = _commands(slide, RX, y, "on the remote node", REMOTE_RUN)

    text(slide, RX, y + 0.10, COL_W, 0.30,
         [("list gives you the endpoint UUID", 11, INK, False, False)],
         align=PP_ALIGN.LEFT)

    y = _heading(slide, RX, y + 0.62, "Then on your own machine")
    _bullet(slide, RX, y,
            "Agent uses the UUID and puts your personal config under "
            "users/<user_id>")
    y = _commands(slide, RX, y + ITEM_H + 0.06, "on your own machine",
                  LOCAL_RUN)

    text(slide, RX, y + 0.16, COL_W, 0.34,
         [("The quick test does not produce files on the remote machine.",
           11, FAINT, False, True)], align=PP_ALIGN.LEFT)
    return slide


if __name__ == "__main__":
    deck = new_deck()
    add(deck)
    dest = os.path.join(HERE, "agentlab_tutorial_globus.pptx")
    deck.save(dest)
    print("wrote", dest)
