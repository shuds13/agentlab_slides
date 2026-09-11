# AgentLab slides

**[Lightning talk →](agentlab.pdf)** · **[Tutorial →](agentlab_tutorial.pdf)**

Talks on [AgentLab](https://github.com/shuds13/AgentLab), built with
`python-pptx`. Every slide is native PowerPoint shapes.

## Build

```
pip install python-pptx
python3 src/build_deck.py
```

or selected deck.

```
python3 src/build_deck.py short
```


That writes a `.pptx` for each deck built, and if
LibreOffice is on the PATH a `.pdf` beside it.


## Decks

`DECKS` in `src/build_deck.py` lists the slides of each deck, in order. A
slide is in a deck because it is named in that deck's list.

| name | | builds | |
|---|---|---|---|
| `short` | the lightning talk | `agentlab.pdf` | [slides](short_talk.md) |
| `long` | the longer talk | `agentlab_long.pdf` | [slides](long_talk.md) |
| `tutorial` | Getting started, then the tutorial | `agentlab_tutorial.pdf` | [slides](tutorial.md) |

```
python3 src/build_deck.py short              # one deck
python3 src/build_deck.py long tutorial      # or several
```

Each slide is also a module exposing `add(prs)`, and each runs standalone to
produce a single-slide file for iterating on one slide in isolation:

```
python3 src/build_workspace.py      # -> agentlab_workspace.pptx
```

`src/slidekit.py` holds the palette and the drawing helpers. Change a colour
there and it changes everywhere.

Where a slide makes a factual claim, its build script's docstring cites the
source. Check those against the AgentLab repository when it changes.
