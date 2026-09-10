#!/usr/bin/env python3
"""Assemble the AgentLab decks.

Two decks are built from the same slide modules. DECKS below is the whole
definition of what each one contains; a slide joins a deck by being listed
there, and belongs to both by being in BODY.

  agentlab       the short talk, and the one the README links
  agentlab_long  the alternative title, a background slide, then the same

  1  title                          (build_title_alt.py in the long deck)
  2  what it is, and where the code is
  3  how a campaign runs            (build_diagram.py)
  4  the workspace                  (build_workspace.py)
  5  human in the loop              (build_slack.py)
  6  built on the Claude Agent SDK  (build_sdk.py)
  7  getting started

Each slide module exposes add(prs) and can still be run on its own to produce a
single-slide pptx.

    python3 build_deck.py

Output: agentlab.pptx / .pdf and agentlab_long.pptx / .pdf
"""

import os
import subprocess

from pptx.enum.text import PP_ALIGN

import build_background
import build_diagram
import build_sdk
import build_slack
import build_title_alt
import build_tutorial
import build_tutorial_dirs
import build_tutorial_globus
import build_tutorial_install
import build_tutorial_slack
import build_workspace
from slidekit import (
    new_deck, blank_slide, title_block, box, label_box, text,
    MONO, INK, MUTED, FAINT, BLUE, BLUE_BG, TEAL, TEAL_BG, AMBER, SLATE,
    GREY_BG, HAIRLINE, VIOLET, VIOLET_BG,
)

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = "github.com/shuds13/AgentLab"
CAS = "github.com/shuds13/cas-framework"


def title_slide(prs):
    slide = blank_slide(prs)
    text(slide, 1.0, 2.35, 11.33, 1.10,
         [("AgentLab", 60, INK, True, False)], align=PP_ALIGN.LEFT)
    text(slide, 1.06, 3.48, 11.33, 0.50,
         [("A research lab run by agents", 22, BLUE, False, True)],
         align=PP_ALIGN.LEFT)
    box(slide, 1.08, 4.22, 2.10, 0.045, fill=BLUE, stroke=BLUE, line_w=0.5,
        radius=None)
    text(slide, 1.06, 4.55, 11.33, 0.60,
         [("Stephen Hudson · Magesh Rajasekaran", 15, SLATE, False, False)],
         align=PP_ALIGN.LEFT)
    text(slide, 1.06, 5.05, 11.33, 0.30,
         [(REPO, 12, MUTED, False, False)], align=PP_ALIGN.LEFT, font=MONO)
    return slide


def intro_slide(prs):
    slide = blank_slide(prs)
    title_block(
        slide,
        "What it is",
        "You give an agent a goal and a system. It runs the investigation itself.",
    )

    # the three sentences that actually define it, from the repo README
    POINTS = [
        ("A lab run by agents",
         "A campaign agent submits work to an HPC system, reads what comes "
         "back, decides what to try next, and keeps going until it has an "
         "answer — for hours or days, with nobody in the loop.",
         BLUE, BLUE_BG),
        ("One campaign per question",
         "Each investigation is a campaign with its own goal, task definition "
         "and records. Campaigns are independent and run side by side, sharing "
         "the framework and the system definitions.",
         TEAL, TEAL_BG),
        ("Visible and steerable",
         "Any number of campaigns, all watchable from one Slack channel, and "
         "redirectable while they run. Nobody stands up new infrastructure "
         "per question.",
         VIOLET, VIOLET_BG),
    ]
    x, w, gap = 0.62, 3.94, 0.30
    for i, (head, bodytext, accent, tint) in enumerate(POINTS):
        left = x + i * (w + gap)
        box(slide, left, 1.78, w, 2.30, fill=tint, stroke=accent, line_w=1.5)
        text(slide, left + 0.26, 2.04, w - 0.52, 0.36,
             [(head, 15, accent, True, False)], align=PP_ALIGN.LEFT)
        text(slide, left + 0.26, 2.54, w - 0.52, 1.45,
             [(bodytext, 11.5, SLATE, False, False)], align=PP_ALIGN.LEFT)

    # where the code is
    box(slide, 0.62, 4.86, 12.10, 0.92, fill=GREY_BG, stroke=HAIRLINE,
        line_w=1.0, radius=0.10)
    text(slide, 0.90, 5.02, 6.00, 0.30,
         [(REPO, 15, INK, True, False)], align=PP_ALIGN.LEFT, font=MONO)
    text(slide, 0.90, 5.34, 11.60, 0.28,
         [(f"built on the same machinery as {CAS}, which coordinates many "
           "agents on a single search campaign", 10.5, MUTED, False, False)],
         align=PP_ALIGN.LEFT)

    text(slide, 0.62, 6.06, 12.10, 0.30,
         [("Runs anywhere the Claude Agent SDK and a Globus Compute endpoint "
           "can be reached — the agent does not need to be on the HPC system.",
           10.5, FAINT, False, True)], align=PP_ALIGN.LEFT)
    return slide


def getting_started_slide(prs):
    """Sources, so the claims stay checkable:

      AGENTS.md:6      "get them from a clone to a running campaign"
      AGENTS.md:11-28  the points that need a human: the compute-system login,
                       the first Globus login, Transfer activation, Slack setup
      AGENTS.md:36     what only they know: system, allocation, queue, writable space
      AGENTS.md:149    start it in tmux; a campaign runs for hours or days and
                       closing the terminal kills it with jobs in flight
      docs/setup.md:7  Python 3.10+, the agent CLI on PATH and authenticated
      docs/setup.md:13 runs on a workstation, a login node or a VM, not on the
                       compute system
    """
    slide = blank_slide(prs)
    title_block(
        slide,
        "Getting started",
        "Clone the repository, start your agent in it, and ask.",
    )

    STEPS = [
        ("1   Clone it", f"git clone https://{REPO}", False),
        ("2   Start your agent in it", "cd AgentLab && claude", False),
        ("3   Say", "“Help me set up.”", True),
    ]
    x, w, gap = 0.62, 3.94, 0.30
    for i, (head, line, quoted) in enumerate(STEPS):
        left = x + i * (w + gap)
        box(slide, left, 1.72, w, 1.32, fill=BLUE_BG, stroke=BLUE, line_w=1.5)
        text(slide, left + 0.26, 1.94, w - 0.52, 0.32,
             [(head, 13, BLUE, True, False)], align=PP_ALIGN.LEFT)
        if quoted:
            text(slide, left + 0.26, 2.40, w - 0.52, 0.40,
                 [(line, 14, INK, True, True)], align=PP_ALIGN.LEFT)
        else:
            text(slide, left + 0.26, 2.44, w - 0.52, 0.34,
                 [(line, 9.5, SLATE, False, False)], align=PP_ALIGN.LEFT,
                 font=MONO)

    PANELS = [
        (0.62, 5.98, TEAL, "What you need", [
            "Claude Agent SDK, and an LLM service it can reach",
            "Python 3.10+, and the claude CLI on PATH and authenticated",
            "Access to a compute system, and a project to charge",
        ]),
        (6.90, 5.82, AMBER, "Where to run", [
            "Preferably a shared persistent node, so the team reaches the "
            "same campaigns",
            "Inside tmux. A campaign runs for hours or days, and closing the "
            "terminal kills it with jobs still in flight",
        ]),
    ]
    for left, w, accent, head, items in PANELS:
        box(slide, left, 3.36, w, 2.90, fill=GREY_BG, stroke=HAIRLINE,
            line_w=1.0, radius=0.08)
        text(slide, left + 0.28, 3.56, w - 0.56, 0.34,
             [(head, 15, accent, True, False)], align=PP_ALIGN.LEFT)
        for j, item in enumerate(items):
            text(slide, left + 0.30, 4.02 + j * 0.55, 0.20, 0.30,
                 [("•", 11, accent, True, False)], align=PP_ALIGN.LEFT)
            text(slide, left + 0.54, 4.00 + j * 0.55, w - 0.86, 0.52,
                 [(item, 11, SLATE, False, False)], align=PP_ALIGN.LEFT)

    text(slide, 0.62, 6.48, 12.10, 0.30,
         [("AGENTS.md in the repository root is the onboarding script — the "
           "agent reads it and walks you through the rest.",
           10.5, MUTED, False, True)], align=PP_ALIGN.LEFT)
    return slide


# everything after the title slide, shared by both decks
BODY = [intro_slide, build_diagram.add, build_workspace.add, build_slack.add,
        build_sdk.add, getting_started_slide]

# the tutorial closes the long talk; swap these two to reorder 2 and 3
TUTORIAL = [build_tutorial.add, build_tutorial_install.add,
            build_tutorial_dirs.add, build_tutorial_globus.add,
            build_tutorial_slack.add]

DECKS = {
    "agentlab": [title_slide] + BODY,
    "agentlab_long": ([build_title_alt.add, build_background.add] + BODY
                      + TUTORIAL),
}


def build(name, slides):
    prs = new_deck()
    for add_slide in slides:
        add_slide(prs)

    pptx = os.path.join(HERE, name + ".pptx")
    prs.save(pptx)
    print("wrote", pptx)

    subprocess.run(["soffice", "--headless", "--convert-to", "pdf",
                    pptx, "--outdir", HERE],
                   check=True, stdout=subprocess.DEVNULL,
                   stderr=subprocess.DEVNULL)
    print("wrote", os.path.join(HERE, name + ".pdf"))


def main():
    for name, slides in DECKS.items():
        build(name, slides)


if __name__ == "__main__":
    main()
