#!/usr/bin/env python3
"""Build tutorial slide 1 — install and run the simple example.

Commands and dependency names are the user's, verbatim.

Output: agentlab_tutorial_install.pptx
"""

import os

from pptx.enum.text import PP_ALIGN

from slidekit import (
    new_deck, blank_slide, title_block, box, text,
    MONO, INK, FAINT, SLATE, GREY_BG, HAIRLINE,
)

HERE = os.path.dirname(os.path.abspath(__file__))

RUN = [
    "cd campaigns/example-quick-optimum/",
    "WATCH=true AGENT_MODEL=haiku ./run.sh",
]


def add(prs):
    slide = blank_slide(prs)

    title_block(
        slide,
        "1   Install and run the simple example",
        "Start your agent harness and say “Help me set up.”",
    )

    text(slide, 0.62, 1.80, 12.10, 0.40,
         [("The agent installs what's missing", 17, SLATE, False, False)],
         align=PP_ALIGN.LEFT)
    text(slide, 0.62, 2.24, 12.10, 0.30,
         [("claude-agent-sdk, globus-compute-sdk", 11, FAINT, False, False)],
         align=PP_ALIGN.LEFT, font=MONO)

    text(slide, 0.62, 3.00, 12.10, 0.40,
         [("Then run it, with a watch", 17, SLATE, False, False)],
         align=PP_ALIGN.LEFT)

    box(slide, 0.62, 3.52, 5.90, 1.40, fill=GREY_BG, stroke=HAIRLINE,
        line_w=1.0, radius=0.08)
    for i, line in enumerate(RUN):
        text(slide, 0.92, 3.76 + i * 0.44, 5.30, 0.34,
             [(line, 13, INK, False, False)], align=PP_ALIGN.LEFT, font=MONO)

    text(slide, 0.62, 5.16, 12.10, 0.40,
         [("Click on the watch URL to get an interactive view of the workspace",
           17, SLATE, False, False)], align=PP_ALIGN.LEFT)

    text(slide, 0.62, 5.78, 12.10, 0.64,
         [("If you do not specify AGENT_MODEL it uses your default Claude Code "
           "model.", 11, FAINT, False, True),
          ("This is a toy example — the smartest model is not needed.",
           11, FAINT, False, True)], align=PP_ALIGN.LEFT)

    return slide


if __name__ == "__main__":
    deck = new_deck()
    add(deck)
    dest = os.path.join(HERE, "agentlab_tutorial_install.pptx")
    deck.save(dest)
    print("wrote", dest)
