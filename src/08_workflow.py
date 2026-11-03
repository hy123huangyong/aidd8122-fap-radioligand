"""
Step 8 -- Workflow schematic for the Methods section (fig0).
"""
from __future__ import annotations

import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

import aidd_utils as U

U.set_style()

INK = "#22303c"
MUTED = "#5f6f7a"


def box(ax, x, y, w, h, title, lines, color, num=None):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.010,rounding_size=0.018",
                                linewidth=1.0, edgecolor=color, facecolor=color + "12"))
    ax.text(x + 0.016, y + h - 0.048, title, fontsize=9.0, fontweight="bold", color=color,
            va="top", ha="left")
    if num is not None:
        ax.text(x + w - 0.016, y + h - 0.040, num, fontsize=14, color=color + "4d",
                va="top", ha="right", fontweight="bold")
    for i, ln in enumerate(lines):
        ax.text(x + 0.016, y + h - 0.100 - i * 0.058, ln, fontsize=6.7, color=INK,
                va="top", ha="left")


def arrow(ax, x0, x1, y):
    ax.add_patch(FancyArrowPatch((x0, y), (x1, y), arrowstyle="-|>", mutation_scale=9,
                                 lw=1.0, color=MUTED, shrinkA=0, shrinkB=0))


def main() -> None:
    fig, ax = plt.subplots(figsize=(9.2, 4.0))
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")

    y, h, w = 0.50, 0.44, 0.222
    xs = [0.012, 0.2525, 0.493, 0.7335]

    box(ax, xs[0], y, w, h, "1  Data", [
        "ChEMBL bioactivities",
        "FAP 972 → 384 cpds",
        "PREP 962 → 512",
        "DPP4 6,553 → 4,338",
        "curation: assay type,",
        "validity, duplicates",
    ], U.PALETTE["FAP"], "A")

    box(ax, xs[1], y, w, h, "2  Models", [
        "Morgan + MACCS + desc.",
        "ridge / SVR / RF / XGB",
        "splits: random,",
        "scaffold, time",
        "paralogue transfer feats.",
        "selectivity: direct / indir.",
    ], U.PALETTE["PREP"], "B")

    box(ax, xs[2], y, w, h, "3  Trust", [
        "applicability domain",
        "(max Tanimoto)",
        "split-conformal intervals",
        "temporal drift",
        "+ intercept recalibration",
        "AUROC, EF$_{10}$, P@10",
    ], U.PALETTE["DPP4"], "C")

    box(ax, xs[3], y, w, h, "4  Design", [
        "graph GA: fragment",
        "crossover + mutation",
        "NSGA-II, 4 objectives:",
        "affinity, margin, QED, SA",
        "filters: 250-650 Da, AD",
        "controls: single-obj, random",
    ], U.PALETTE["accent"], "D")

    for i in range(3):
        arrow(ax, xs[i] + w + 0.006, xs[i + 1] - 0.006, y + h / 2)

    # feedback loop from trust back to models
    ax.add_patch(FancyArrowPatch((xs[2] + w / 2, y - 0.012), (xs[1] + w / 2, y - 0.012),
                                 connectionstyle="arc3,rad=-0.26", arrowstyle="-|>",
                                 mutation_scale=9, lw=0.9, color=MUTED, ls=(0, (3, 2))))
    ax.text((xs[1] + xs[2] + w) / 2, y - 0.150,
            "error analysis feeds back into feature and model choice",
            fontsize=6.9, color=MUTED, ha="center", style="italic")

    ax.text(0.012, 0.285, "Outputs", fontsize=8.0, fontweight="bold", color=INK)
    outs = ["dataset + curation log", "benchmark, 5 models × 3 splits",
            "calibrated error bars", "Pareto set of novel warheads"]
    for i, o in enumerate(outs):
        x = 0.012 + i * 0.2485
        ax.text(x, 0.232, "•", fontsize=7.0, color=MUTED, va="top")
        ax.text(x + 0.013, 0.232, o, fontsize=6.9, color=MUTED, va="top")

    ax.text(0.012, 0.135,
            "Every experiment is repeated over 5 random seeds with identical feature blocks and "
            "hyper-parameters; curated tables, figures and code are released with the report.",
            fontsize=6.9, color=MUTED, style="italic")

    fig.savefig(U.FIG / "fig0_workflow.png", dpi=330, bbox_inches="tight")
    plt.close(fig)
    print("fig0 written")


if __name__ == "__main__":
    main()
