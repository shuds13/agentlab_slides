#!/usr/bin/env python3
"""Build the AgentLab flow diagram slide.

Facts encoded here come from AGENTLAB-PRESENTATION-HANDOFF.md:
  - the endpoint runs on the compute system, not with the agent
  - Globus Compute runs tasks; the endpoint turns tasks into batch jobs
  - the agent runs anywhere; it does not need to be on the HPC system

Output: agentlab_diagram.pptx  (16:9, white background, native editable shapes)
"""

import os

from pptx.enum.text import PP_ALIGN

from slidekit import (
    new_deck, blank_slide, title_block, box, label_box, text, arrow,
    INK, MUTED, FAINT, BLUE, BLUE_BG, VIOLET, VIOLET_BG, TEAL, TEAL_BG,
    AMBER, AMBER_BG, SLATE, GREY, GREY_BG, PANEL_BG, HAIRLINE, ARROW,
    RETURN, WHITE,
)

HERE = os.path.dirname(os.path.abspath(__file__))
FLOW_Y = 3.20   # every box on the main flow is centred on this line


def add(prs):
    slide = blank_slide(prs)

    title_block(
        slide,
        "How a campaign runs",
        "You give an agent a goal and a system. It runs the investigation itself.",
    )

    # ---- LLM service
    llm = box(slide, 0.5, 1.32, 3.10, 0.72, fill=VIOLET_BG, stroke=VIOLET)
    label_box(llm, [("LLM service", 13, VIOLET, True, False),
                    ("Argo · the agent's reasoning", 9.5, MUTED, False, False)])

    # ---- your machine panel
    box(slide, 0.5, 2.30, 3.10, 3.30, fill=PANEL_BG, stroke=HAIRLINE,
        line_w=1.0, dash=True)

    agent = box(slide, 0.75, 2.61, 2.60, 1.18, fill=BLUE_BG, stroke=BLUE, line_w=1.75)
    label_box(agent, [("Campaign agent", 15, BLUE, True, False),
                      ("Claude Agent SDK", 9.5, MUTED, False, False),
                      ("long-lived process", 9.5, MUTED, False, False)])

    ws = box(slide, 0.75, 4.12, 2.60, 0.88, fill=GREY_BG, stroke=GREY, line_w=1.25)
    label_box(ws, [("workspace/", 13, SLATE, True, False),
                   ("results · journal · logbook", 9, MUTED, False, False)])

    text(slide, 0.62, 5.08, 2.86, 0.44,
         [("Your machine — workstation,", 9, FAINT, False, False),
          ("login node, or GCE node", 9, FAINT, False, False)])

    # LLM <-> agent
    arrow(slide, 2.05, 2.04, 2.05, 2.61, color=VIOLET, width=1.5, head=True, tail=True)
    text(slide, 2.16, 2.13, 1.40, 0.26,
         [("what to try next", 8.5, VIOLET, False, True)], align=PP_ALIGN.LEFT)

    # agent -> workspace
    arrow(slide, 2.05, 3.79, 2.05, 4.12, color=GREY, width=1.4)
    text(slide, 2.16, 3.82, 1.20, 0.26,
         [("writes", 8.5, MUTED, False, True)], align=PP_ALIGN.LEFT)

    # ---- Globus Compute
    gc = box(slide, 4.70, FLOW_Y - 0.725, 2.15, 1.45, fill=TEAL_BG, stroke=TEAL, line_w=1.75)
    label_box(gc, [("Globus Compute", 15, TEAL, True, False),
                   ("one task =", 9.5, MUTED, False, False),
                   ("one Python function", 9.5, MUTED, False, False)])

    arrow(slide, 3.60, FLOW_Y, 4.70, FLOW_Y)
    text(slide, 3.58, FLOW_Y - 0.36, 1.14, 0.28,
         [("submit task", 9.5, ARROW, False, False)])

    # the function is not necessarily small - say so where the phrase appears
    text(slide, 4.48, 4.02, 2.60, 0.50,
         [("…and that function can launch", 9, MUTED, False, True),
          ("a multi-node MPI run", 9, MUTED, False, True)])

    # ---- compute systems
    box(slide, 8.10, 1.75, 4.73, 3.85, fill=None, stroke=HAIRLINE, line_w=1.0, dash=True)
    text(slide, 8.30, 1.87, 4.33, 0.30,
         [("Compute systems", 11.5, SLATE, True, False)], align=PP_ALIGN.LEFT)

    box(slide, 8.35, 2.32, 4.23, 1.72, fill=AMBER_BG, stroke=AMBER, line_w=1.5)
    text(slide, 8.52, 2.40, 2.00, 0.28,
         [("Aurora", 13, AMBER, True, False)], align=PP_ALIGN.LEFT)

    ep = box(slide, 8.55, FLOW_Y - 0.34, 1.78, 0.68, fill=WHITE, stroke=AMBER, line_w=1.25)
    label_box(ep, [("Endpoint", 11, AMBER, True, False),
                   ("runs on the system", 8, MUTED, False, False)])

    bj = box(slide, 10.75, FLOW_Y - 0.34, 1.60, 0.68, fill=WHITE, stroke=SLATE, line_w=1.25)
    label_box(bj, [("Batch job", 11, SLATE, True, False),
                   ("via the scheduler", 8, MUTED, False, False)])

    arrow(slide, 10.33, FLOW_Y, 10.75, FLOW_Y, width=1.5)

    text(slide, 8.50, 3.60, 3.95, 0.32,
         [("the endpoint turns tasks into batch jobs", 8.5, MUTED, False, True)],
         align=PP_ALIGN.LEFT)

    pol = box(slide, 8.35, 4.22, 4.23, 0.52, fill=AMBER_BG, stroke=AMBER, line_w=1.25)
    label_box(pol, [("Polaris — same pattern", 11.5, AMBER, True, False)])

    text(slide, 8.35, 4.92, 4.23, 0.32,
         [("…or anywhere with an endpoint", 10.5, MUTED, False, True)])

    # Globus -> endpoint on Aurora
    arrow(slide, 6.85, FLOW_Y, 8.55, FLOW_Y)
    text(slide, 6.83, FLOW_Y - 0.36, 1.29, 0.28,
         [("route to endpoint", 9, ARROW, False, False)])

    # ---- the return path: result comes back the same way
    arrow(slide, 10.40, 5.60, 10.40, 6.15, color=RETURN, width=1.75, head=False)
    arrow(slide, 10.40, 6.15, 2.05, 6.15, color=RETURN, width=1.75, head=False)
    arrow(slide, 2.05, 6.15, 2.05, 5.60, color=RETURN, width=1.75, head=True)
    text(slide, 4.90, 5.82, 2.70, 0.30,
         [("result returns the same way", 9.5, RETURN, True, False)])

    # ---- footer
    text(slide, 0.5, 6.58, 12.33, 0.40,
         [("The loop is the point: submit, read, decide, submit again — "
           "for hours or days, without a person in it.", 12, SLATE, False, False)],
         align=PP_ALIGN.LEFT)
    return slide


if __name__ == "__main__":
    deck = new_deck()
    add(deck)
    dest = os.path.join(HERE, "agentlab_diagram.pptx")
    deck.save(dest)
    print("wrote", dest)
