#!/usr/bin/env python3
"""Build tutorial slide 3 — add Slack integration.

Invite link, scope names and wording are the user's, verbatim.

Output: agentlab_tutorial_slack.pptx
"""

import os

from pptx.enum.text import PP_ALIGN

from slidekit import (
    ROOT,
    new_deck, blank_slide, title_block, text,
    MONO, INK, FAINT, BLUE, SLATE,
)


HEAD_H, ITEM_H = 0.36, 0.32
COL_W = 5.85
LX, RX = 0.62, 6.86

INVITE = ("https://join.slack.com/t/agentlab-co/shared_invite/"
          "zt-47fcg5wpv-PnypDIKMB3TYBUZNJitwqA")

SCOPES = ("OAuth scopes: channels:history → public channel,  "
          "groups:history → private channel.")


def _heading(slide, x, y, head, width=COL_W):
    """Give a heading that wraps the second line it needs, or it lands on the
    first bullet under it."""
    h = HEAD_H if len(head) <= 62 else HEAD_H + 0.26
    text(slide, x, y, width, h,
         [(head, 12.5, SLATE, True, False)], align=PP_ALIGN.LEFT)
    return y + h


def _bullet(slide, x, y, item, width=COL_W):
    text(slide, x + 0.10, y, 0.20, 0.28,
         [("•", 11, BLUE, True, False)], align=PP_ALIGN.LEFT)
    text(slide, x + 0.34, y, width - 0.34, 0.30,
         [(item, 11, INK, False, False)], align=PP_ALIGN.LEFT)
    return y + ITEM_H


def add(prs):
    slide = blank_slide(prs)

    title_block(
        slide,
        "3   Add Slack integration",
        "Notifications one way, or talk to the agents.",
    )

    y = _heading(slide, LX, 1.58, "Please join the AgentLab Slack workspace.",
                 width=12.10)
    text(slide, LX + 0.10, y, 11.90, 0.30,
         [(INVITE, 10, INK, False, False)], align=PP_ALIGN.LEFT, font=MONO)
    _bullet(slide, LX, y + 0.34, "Create a new channel in the workspace — you will point your app at this channel",
            width=12.10)
    text(slide, LX + 0.10, y + 0.72, 11.90, 0.30,
         [("link expires…", 11, FAINT, False, True)], align=PP_ALIGN.LEFT)

    y = _heading(slide, LX, 3.16,
                 "Ask your agent (in AgentLab repo) to help setup slack.")
    y = _bullet(slide, LX, y, "It will take you through the process")
    y = _bullet(slide, LX, y, "Choose: setting up new lab")
    _bullet(slide, LX, y, "You can set up notifications only, or 2-way comms.")

    y = _heading(slide, RX, 3.16,
                 "Create a slack app for any workspace — some may require "
                 "approval.")
    _bullet(slide, RX, y, "Approval request happens in the process.")

    y = _heading(slide, LX, 4.86,
                 "Note to allow messaging Agent from slack:", width=12.10)
    text(slide, LX + 0.10, y, 11.90, 0.32,
         [(SCOPES, 11, INK, False, False)], align=PP_ALIGN.LEFT, font=MONO)

    text(slide, LX, 5.86, 12.10, 0.34,
         [("Your agent should set up a lab.yaml. You can then start the slack "
           "bridge with", 12, SLATE, False, False)], align=PP_ALIGN.LEFT)
    text(slide, LX + 0.10, 6.24, 11.90, 0.32,
         [("bin/lab.sh start", 12, INK, False, False)],
         align=PP_ALIGN.LEFT, font=MONO)
    return slide


if __name__ == "__main__":
    deck = new_deck()
    add(deck)
    dest = os.path.join(ROOT, "agentlab_tutorial_slack.pptx")
    deck.save(dest)
    print("wrote", dest)
