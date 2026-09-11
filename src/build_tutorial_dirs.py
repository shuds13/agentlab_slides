#!/usr/bin/env python3
"""Build the second slide of tutorial part 1 — the user and campaign directories.

File names and descriptions are the user's, verbatim. The blank row between
the third and fourth campaign file keeps their grouping: the prompts and the
method first, then the task, its settings and the launcher.

Output: agentlab_tutorial_dirs.pptx
"""

import os

from pptx.enum.text import PP_ALIGN

from slidekit import (
    ROOT,
    new_deck, blank_slide, title_block, text,
    MONO, INK, FAINT, BLUE, SLATE,
)


NAME_X, DESC_X = 0.62, 2.37
ROW_H, GROUP_GAP = 0.42, 0.22

FILES = [
    [("prompt.md", "The overall challenge / campaign."),
     ("method.md", "Describes research cycles (defaults to methods/standard.md)."),
     ("user_prompt.md", "Instructions and goal for this run of the agent.")],
    [("task.py", "Task or tasks to run locally or on remote system."),
     ("campaign.json", "task run settings."),
     ("run.sh", "Setting and starts agent.")],
]


def add(prs):
    slide = blank_slide(prs)

    title_block(
        slide,
        "User and campaign directories",
        "What you keep once, and what each campaign holds.",
    )

    text(slide, NAME_X, 1.58, 6.00, 0.34,
         [("users/<user_id>", 13, SLATE, True, False)],
         align=PP_ALIGN.LEFT, font=MONO)
    text(slide, NAME_X + 0.10, 2.00, 0.20, 0.28,
         [("•", 11, BLUE, True, False)], align=PP_ALIGN.LEFT)
    text(slide, NAME_X + 0.34, 2.00, 10.00, 0.30,
         [("For each remote system — Endpoint UUID / Account / Work-dir",
           11, INK, False, False)], align=PP_ALIGN.LEFT)

    text(slide, NAME_X, 2.72, 6.00, 0.34,
         [("Campaign dir files", 13, SLATE, True, False)], align=PP_ALIGN.LEFT)

    y = 3.16
    for group in FILES:
        for name, desc in group:
            text(slide, NAME_X, y, 1.70, 0.32,
                 [(name, 12, INK, False, False)], align=PP_ALIGN.LEFT,
                 font=MONO)
            text(slide, DESC_X, y, 10.30, 0.32,
                 [(desc, 11.5, SLATE, False, False)], align=PP_ALIGN.LEFT)
            y += ROW_H
        y += GROUP_GAP

    text(slide, NAME_X, y + 0.10, 12.10, 0.34,
         [("user_prompt.md is most likely to be modified between agent runs "
           "in an ongoing campaign.", 11, FAINT, False, True)],
         align=PP_ALIGN.LEFT)
    return slide


if __name__ == "__main__":
    deck = new_deck()
    add(deck)
    dest = os.path.join(ROOT, "agentlab_tutorial_dirs.pptx")
    deck.save(dest)
    print("wrote", dest)
