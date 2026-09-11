# AgentLab slides

**[Lightning talk →](agentlab.pdf)** · **[Tutorial →](agentlab_tutorial.pdf)**

Talks on [AgentLab](https://github.com/shuds13/AgentLab), built with
`python-pptx`. Every slide is native PowerPoint shapes, so the `.pptx` can be
edited directly — but a rebuild overwrites it, so edit by hand only once you
have stopped regenerating.

## Build

```
pip install python-pptx
python3 build_deck.py
```

That writes a `.pptx` for each deck and, if LibreOffice is on the PATH, a
`.pdf` beside it.

## Decks

`DECKS` in `build_deck.py` lists the slides of each deck, in order. A slide is
in a deck because it is named in that deck's list.

| | | |
|---|---|---|
| `agentlab` | the lightning talk | [slides](short_talk.md) |
| `agentlab_long` | the longer talk | [slides](long_talk.md) |
| `agentlab_tutorial` | Getting started, then the tutorial | [slides](tutorial.md) |

Each slide is also a module exposing `add(prs)`, and each runs standalone to
produce a single-slide file for iterating on one slide in isolation:

```
python3 build_workspace.py      # -> agentlab_workspace.pptx
```

`slidekit.py` holds the palette and the drawing helpers. Change a colour there
and it changes everywhere.

Where a slide makes a factual claim, its build script's docstring cites the
source. Check those against the AgentLab repository when it changes.
