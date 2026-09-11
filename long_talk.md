# Longer talk — `long`

No PDF is tracked for this one. Build it with
`python3 src/build_deck.py long`.

| | |
|---|---|
| 1 | title, with the full author list and logos — `src/build_title_alt.py` |
| 2 | background — `src/build_background.py` |
| 3 | what it is, and where the code is |
| 4 | how a campaign runs — `src/build_diagram.py` |
| 5 | the workspace — `src/build_workspace.py` |
| 6 | human in the loop — `src/build_slack.py` |
| 7 | built on the Claude Agent SDK — `src/build_sdk.py` |
| 8 | other features and example campaigns — `src/build_features.py` |
| 9 | future — `src/build_future.py` |
| 10 | getting started |
| 11–16 | the tutorial — see [tutorial.md](tutorial.md) |

Slides 3–7 and 10 are the lightning talk's, unchanged. The title slide is a
separate module rather than a variant of the short deck's, so the two can drift
without either breaking.
