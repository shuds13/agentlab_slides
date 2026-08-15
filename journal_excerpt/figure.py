#!/usr/bin/env python3
"""Figure and table for the illustrative journal page.

The campaign is invented. Latencies come from the standard cost model for the
two collectives, so the curves, the table and the prose are all the same
numbers:

    ring:               2(N-1) steps,  2(N-1)/N * bytes moved per rank
    recursive doubling:  2 log2(N) steps,  log2(N) * bytes moved per rank

Writes crossover.pdf and table_rows.tex, both \\input by excerpt.tex, so the
page cannot disagree with itself.
"""

import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D

A = 3.0     # us per communication step
B = 0.10    # us per KiB moved

RANKS = [32, 128, 512]
COLOURS = {32: "#1f77b4", 128: "#d62728", 512: "#2ca02c"}


def ring(size_kib, n):
    return A * 2 * (n - 1) + B * size_kib * 2 * (n - 1) / n


def recursive_doubling(size_kib, n):
    return A * 2 * np.log2(n) + B * size_kib * np.log2(n)


def crossover(n):
    """Size at which ring becomes the cheaper of the two."""
    lg = np.log2(n)
    return A * 2 * (n - 1 - lg) / (B * (lg - 2 * (n - 1) / n))


plt.rcParams.update({
    "font.family": "serif",
    "font.size": 9,
    "axes.linewidth": 0.7,
    "xtick.major.width": 0.7,
    "ytick.major.width": 0.7,
})

size = np.logspace(np.log2(16), np.log2(32768), 240, base=2)   # KiB
fig, ax = plt.subplots(figsize=(6.4, 2.7))

for n in RANKS:
    ax.plot(size, ring(size, n), color=COLOURS[n], lw=1.4)
    ax.plot(size, recursive_doubling(size, n), color=COLOURS[n], lw=1.4, ls="--")
    x = crossover(n)
    ax.plot(x, ring(x, n), "o", color=COLOURS[n], ms=5.0,
            mec="black", mew=0.6, zorder=5)

# colour carries the rank count; solid/dashed is explained in the caption
handles = [Line2D([], [], color=COLOURS[n], lw=1.4, label=f"{n} ranks")
           for n in RANKS]

ax.set_xscale("log", base=2)
ax.set_yscale("log")
ax.set_xlabel("message size (KiB)")
ax.set_ylabel("all-reduce latency (µs)")
ax.set_xticks([16, 128, 1024, 8192])
ax.set_xticklabels(["16", "128", "1 Ki", "8 Ki"])
ax.grid(True, which="major", lw=0.35, alpha=0.4)
ax.legend(handles=handles, fontsize=8, ncol=3, loc="upper left",
          framealpha=0.9, borderpad=0.35, columnspacing=1.2, handlelength=1.8)
fig.tight_layout(pad=0.25)
fig.savefig("crossover.pdf")

# ---- the same numbers, as table rows for the page
MIB = 1024.0
with open("table_rows.tex", "w") as fh:
    for n in RANKS:
        x = crossover(n)
        r, d = ring(MIB, n), recursive_doubling(MIB, n)
        winner = "ring" if r < d else "rec.\\ doubling"
        fh.write(f"{n} & {x:,.0f} & {r:,.0f} & {d:,.0f} & {winner} \\\\\n")
        print(f"{n:4d} ranks: crossover {x:8.1f} KiB   "
              f"ring@1MiB {r:7.1f}   rd@1MiB {d:7.1f}   -> {winner}")
