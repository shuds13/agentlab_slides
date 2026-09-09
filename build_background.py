#!/usr/bin/env python3
"""Build the background slide, for the longer talk.

The origin of this work, in the order it happened: the campaign it was built
for, the constraints that campaign imposed, and the two frameworks that came
out of it. Specific by intent — the general form is the short deck's job.

Output: agentlab_background.pptx
"""

import os

from pptx.enum.text import PP_ALIGN

from slidekit import (
    new_deck, blank_slide, title_block, text, BLUE, SLATE, INK,
)

HERE = os.path.dirname(os.path.abspath(__file__))

HEAD_H, ITEM_H, GROUP_GAP = 0.36, 0.32, 0.30
TOP = 1.60

COLUMNS = [
    (0.62, [
        ("ANL / JPMC — quantum search", SLATE, [
            "A complex problem space",
            "Agent reasoning driving HPC work",
            "Framework developed — Hudson and Rajasekaran",
        ]),
        ("Claude Agent SDK", SLATE, [
            "Claude Code's features in programmatic form",
        ]),
        ("Keep the agents off the HPC systems", SLATE, [
            "Agent on a persistent node with an LLM gateway, submits work to HPC",
            "Stop and restart a campaign at any time",
        ]),
    ]),
    (6.86, [
        ("A shared workspace the agents work in", SLATE, [
            "KKT (explore) and Polaris (explore / deepen)",
            "Read common files, don't overlap work",
            "Hypothesis cycles, logbook and journal",
        ]),
        ("Comms matter", SLATE, [
            "Long-running agents — talk to them, get status",
            "Guidance without stopping them",
        ]),
        ("Extract a generic framework", BLUE, [
            "CAS — Collaborative Agentic Search: many agents on one campaign",
            "AgentLab — the framework reused for long-running agent / HPC tasks",
        ]),
    ]),
]
COL_W = 5.85


def add(prs):
    slide = blank_slide(prs)

    title_block(
        slide,
        "Background",
        "Where this came from, and what the work it was built for demanded.",
    )

    for x, groups in COLUMNS:
        y = TOP
        for head, accent, items in groups:
            text(slide, x, y, COL_W, HEAD_H,
                 [(head, 12.5, accent, True, False)], align=PP_ALIGN.LEFT)
            y += HEAD_H
            for item in items:
                text(slide, x + 0.10, y, 0.20, 0.28,
                     [("•", 11, BLUE, True, False)], align=PP_ALIGN.LEFT)
                text(slide, x + 0.34, y, COL_W - 0.34, 0.30,
                     [(item, 11, INK, False, False)], align=PP_ALIGN.LEFT)
                y += ITEM_H
            y += GROUP_GAP
    return slide


if __name__ == "__main__":
    deck = new_deck()
    add(deck)
    dest = os.path.join(HERE, "agentlab_background.pptx")
    deck.save(dest)
    print("wrote", dest)
