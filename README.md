# AgentLab slides

**[View the deck →](agentlab.pdf)**

A talk on [AgentLab](https://github.com/shuds13/AgentLab), built with
`python-pptx`. Every slide is native PowerPoint shapes, so the `.pptx` can be
edited directly — but a rebuild overwrites it, so edit by hand only once you
have stopped regenerating.

## Build

```
pip install python-pptx
python3 build_deck.py
```

That writes `agentlab.pptx` and, if LibreOffice is on the PATH,
`agentlab.pdf`.

Each slide is also a module exposing `add(prs)`, and each runs standalone to
produce a single-slide file for iterating on one slide in isolation:

```
python3 build_workspace.py      # -> agentlab_workspace.pptx
```

## Slides

| | |
|---|---|
| 1 | title |
| 2 | what it is, and where the code is |
| 3 | how a campaign runs — `build_diagram.py` |
| 4 | the workspace — `build_workspace.py` |
| 5 | human in the loop — `build_slack.py` |
| 6 | built on the Claude Agent SDK — `build_sdk.py` |
| 7 | getting started |

`slidekit.py` holds the palette and the drawing helpers. Change a colour there
and it changes everywhere.

## The journal page

Slide 4 shows a page of a research journal. It is a real LaTeX document
compiled at letter size and then shown small, which is what makes the type sit
at the proportions of a real paper rather than an oversized mock-up.

The campaign in it is invented. Its numbers come from `figure.py`, which
computes them from the cost models of the two collectives and writes both the
plot and the table rows, so the prose, the table and the figure cannot
disagree.

`journal_excerpt/excerpt.png` is tracked, so building the deck needs only
`python-pptx`. To change the journal page you need LaTeX and matplotlib:

```
cd journal_excerpt
python3 figure.py
pdflatex excerpt.tex
pdftoppm -r 200 -png -f 1 -l 1 -singlefile excerpt.pdf excerpt
```

## Sources

Slides 3–7 make factual claims about AgentLab. Where a claim comes from a
specific file, the build script's docstring cites it — for example the file
list on slide 4 is from `framework/SYSTEM.md` and `framework/tools.py`. Check
those references against the repository when it changes.

Slide 6 describes the Claude Agent SDK rather than AgentLab, so `build_sdk.py`
cites the SDK documentation instead.

