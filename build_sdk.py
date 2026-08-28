#!/usr/bin/env python3
"""Build the "what the agent is built on" slide.

Sources, so the claims stay checkable — the Claude Agent SDK documentation:

  code.claude.com/docs/en/agent-sdk/overview
      the Capabilities table, listed on this slide in the docs' own order:
      built-in tools, hooks, subagents, MCP, permissions, sessions, skills /
      commands / memory, plugins. The subtitle is that page's own sentence.

  code.claude.com/docs/en/agent-sdk/sessions
      what makes this possible: a session is written to disk, resume picks one
      up by its id, and a fork leaves the original unchanged. Sessions live in
      ~/.claude/projects/<encoded-cwd>/*.jsonl — the same store interactive
      Claude Code writes, and the file has to be on the machine resuming it.

The commands shown are AgentLab's own, not the SDK's option names:

  bin/list_agents.sh --all               the session ids
  RESUME_SESSION=<session-id> ./run.sh   start a campaign from a session
  claude -r <session-id>                 open one by hand

The LiteLLM note is about this deployment, not the SDK.

Output: agentlab_sdk.pptx
"""

import os

from pptx.enum.text import PP_ALIGN

from slidekit import (
    new_deck, blank_slide, title_block, box, text,
    MONO, MUTED, BLUE, INK, SLATE, GREY_BG, HAIRLINE, VIOLET,
)

HERE = os.path.dirname(os.path.abspath(__file__))


def add(prs):
    slide = blank_slide(prs)

    title_block(
        slide,
        "Built on the Claude Agent SDK",
        "Everything that makes Claude Code powerful is available in the SDK.",
    )

    # ------------------------------------------------ the capabilities, as a list
    text(slide, 0.62, 1.58, 5.60, 0.32,
         [("What comes with it", 13, SLATE, True, False)], align=PP_ALIGN.LEFT)

    CAPS = [
        ["Built-in tools", "Subagents", "Hooks", "MCP"],
        ["Permissions", "Sessions", "Skills, commands and memory", "Plugins"],
    ]
    for items, cx, cw in zip(CAPS, (0.72, 3.45), (2.50, 2.90)):
        for row, item in enumerate(items):
            y = 2.16 + row * 0.62
            text(slide, cx, y, 0.20, 0.30,
                 [("•", 13, BLUE, True, False)], align=PP_ALIGN.LEFT)
            text(slide, cx + 0.24, y, cw, 0.34,
                 [(item, 13, INK, False, False)], align=PP_ALIGN.LEFT)

    # ------------------------------------------- the model behind it is not fixed
    box(slide, 0.62, 5.05, 5.60, 1.10, fill=GREY_BG, stroke=HAIRLINE,
        line_w=1.0, radius=0.08)
    text(slide, 0.90, 5.30, 5.04, 0.30,
         [("LiteLLM", 12.5, VIOLET, True, False)], align=PP_ALIGN.LEFT)
    text(slide, 0.90, 5.66, 5.04, 0.30,
         [("Easily deploy a proxy for multi-vendor models.",
           11, SLATE, False, False)], align=PP_ALIGN.LEFT)

    # ------------------------------------------------------ what sessions give you
    RX, RW = 6.72, 6.00
    text(slide, RX, 1.58, RW, 0.32,
         [("A session is written to disk, so a run can be picked up later",
           13, SLATE, True, False)], align=PP_ALIGN.LEFT)

    # how you find a session in the first place
    text(slide, RX, 2.10, 2.50, 0.28,
         [("bin/list_agents.sh --all", 11, INK, False, False)],
         align=PP_ALIGN.LEFT, font=MONO)
    text(slide, RX + 2.62, 2.11, 3.30, 0.28,
         [("lists the session ids", 11, MUTED, False, False)],
         align=PP_ALIGN.LEFT)

    SESSIONS = [
        (2.68, "Fork a session from interactive Claude Code",
         "Work the problem by hand, then hand the session to a campaign agent.",
         "RESUME_SESSION=<session-id> ./run.sh"),
        (4.57, "Run a postmortem with the research agent",
         "Reopen a finished campaign's session and ask what happened.",
         "claude -r <session-id>"),
    ]
    for top, head, bodytext, call in SESSIONS:
        box(slide, RX, top, RW, 1.58, fill=GREY_BG, stroke=HAIRLINE,
            line_w=1.0, radius=0.08)
        text(slide, RX + 0.28, top + 0.24, RW - 0.56, 0.32,
             [(head, 12.5, BLUE, True, False)], align=PP_ALIGN.LEFT)
        text(slide, RX + 0.28, top + 0.64, RW - 0.56, 0.32,
             [(bodytext, 11, SLATE, False, False)], align=PP_ALIGN.LEFT)
        text(slide, RX + 0.28, top + 1.06, RW - 0.56, 0.28,
             [(call, 11, INK, False, False)], align=PP_ALIGN.LEFT, font=MONO)

    return slide


if __name__ == "__main__":
    deck = new_deck()
    add(deck)
    dest = os.path.join(HERE, "agentlab_sdk.pptx")
    deck.save(dest)
    print("wrote", dest)
