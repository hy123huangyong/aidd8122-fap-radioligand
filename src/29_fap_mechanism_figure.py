# -*- coding: utf-8 -*-
"""Draw the FAP mechanism panel used on the background slide.

Style target: a review-figure cross-section, not a UI mock-up. One continuous
scene — a phospholipid bilayer, a FAP dimer spanning it, and a 68Ga-FAPI
molecule approaching the active site. Only two short leader lines are used;
every other label sits next to what it names. Equal-aspect data coordinates
(x 0-1.4, y 0-1) keep circles round and the protein to scale.

Schematic only: the lipid and protein shapes are symbolic, not a structure
depiction. Drawn near final size so the type survives insertion.
"""
from __future__ import annotations

from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.patches import Circle, Ellipse, FancyArrowPatch, Polygon, Rectangle

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "assets" / "fap_mechanism.png"

INK = "#233240"
GREY = "#7B8794"
FAINT = "#A9B6C1"
MEMB = "#D5DEE6"
PROT = "#2C4A6B"
AMBER = "#A2701F"
TEAL = "#0F7C8A"

plt.rcParams.update({
    "font.family": "sans-serif",
    "font.sans-serif": ["Arial", "Helvetica", "DejaVu Sans"],
    "font.size": 7.0,
})


def hexagon(cx, cy, r, rot=0.0):
    a = np.linspace(0, 2 * np.pi, 6, endpoint=False) + rot
    return np.column_stack([cx + r * np.cos(a), cy + r * np.sin(a)])


def bilayer(ax) -> None:
    """Head groups at the two faces of a pale hydrophobic core.

    At the size this panel is used, lipid tails turn into visual noise, so the
    bilayer is read from the two rows of heads and the core band alone.
    """
    ax.add_patch(Rectangle((0.024, 0.360), 1.352, 0.122, facecolor="#E7EDF2",
                           edgecolor="none", zorder=1))
    for y_edge in (0.482, 0.360):
        ax.plot([0.024, 1.376], [y_edge, y_edge], color="#C7D2DB", lw=0.6, zorder=2)
    xs = np.arange(0.032, 1.374, 0.031)
    for x in xs:
        for y_head in (0.482, 0.360):
            ax.add_patch(Circle((x, y_head), 0.0098, facecolor=MEMB,
                                edgecolor=FAINT, lw=0.5, zorder=3))


def fap_dimer(ax) -> None:
    """Two subunits: catalytic ectodomain, transmembrane helix, cytosolic tail.

    Drawn large: at presentation size this is the one structure the audience has
    to read, so it carries more weight than strict scale would give it.
    """
    cx = 0.700
    for sign in (-1, 1):
        # transmembrane helix first, so the ectodomain covers its top end
        ax.add_patch(Rectangle((cx + sign * 0.042 - 0.023, 0.334), 0.046, 0.262,
                               facecolor=PROT, edgecolor="none", zorder=6))
        # catalytic ectodomain; the two overlap so the dimer reads as one body
        ax.add_patch(Ellipse((cx + sign * 0.062, 0.660), 0.150, 0.230,
                             angle=6 if sign > 0 else -6,
                             facecolor=PROT, edgecolor="none", zorder=7))
    # catalytic cleft: a shallow notch at the top face of the dimer
    ax.add_patch(Polygon([[cx - 0.030, 0.778], [cx + 0.030, 0.778], [cx, 0.690]],
                         closed=True, facecolor="white", edgecolor="none", zorder=8))


def tracer(ax) -> None:
    """Quinoline warhead, nitrile, short linker, DOTA ring and the 68Ga label."""
    tipy = 0.800
    ax.add_patch(Polygon(hexagon(0.990, tipy, 0.031), closed=True, facecolor="white",
                         edgecolor=AMBER, lw=0.9, zorder=8))
    ax.add_patch(Polygon(hexagon(1.043, tipy, 0.027, rot=np.pi / 6), closed=True,
                         facecolor="white", edgecolor=AMBER, lw=0.9, zorder=8))
    # nitrile warhead on the left of the bicycle
    ax.plot([0.959, 0.940], [tipy, tipy], color=AMBER, lw=1.0, zorder=8)
    ax.text(0.928, tipy, "N", fontsize=6.2, color=AMBER, ha="center", va="center", zorder=9)
    # linker to the chelator
    ax.plot([1.070, 1.086, 1.101], [tipy, tipy + 0.021, tipy], color=GREY, lw=0.9, zorder=8)
    # DOTA ring with the radiometal at its centre
    ax.add_patch(Circle((1.120, tipy), 0.0260, facecolor="none", edgecolor=GREY,
                        lw=0.9, zorder=8))
    ax.add_patch(Circle((1.120, tipy), 0.0062, facecolor=TEAL, edgecolor="none", zorder=9))
    ax.text(1.120, tipy - 0.050, "DOTA", fontsize=5.8, color=GREY, ha="center", va="top")


def draw() -> None:
    fig = plt.figure(figsize=(4.90, 2.80), dpi=300)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, 1.40)
    ax.set_ylim(0.20, 1.00)
    ax.set_aspect("equal")
    ax.axis("off")

    bilayer(ax)
    fap_dimer(ax)
    tracer(ax)

    # the only directional arrow: tracer approaching the active site (short)
    ax.add_patch(FancyArrowPatch((0.916, 0.812), (0.760, 0.788),
                                 arrowstyle="-|>", mutation_scale=6.5, lw=0.8,
                                 color=AMBER, shrinkA=1, shrinkB=2,
                                 connectionstyle="arc3,rad=0.18", zorder=9))

    # compartments: plain text at the margin, no leaders
    ax.text(0.030, 0.940, "Extracellular", fontsize=7.4, color=INK,
            fontweight="bold", va="center")
    ax.text(0.030, 0.246, "Cytosol", fontsize=7.4, color=INK,
            fontweight="bold", va="center")

    # two short leaders only
    ax.annotate("phospholipid bilayer", xy=(0.176, 0.482), xytext=(0.040, 0.552),
                fontsize=6.6, color=GREY, va="center",
                arrowprops=dict(arrowstyle="-", lw=0.5, color=FAINT, shrinkA=2, shrinkB=2))
    ax.annotate("FAP dimer", xy=(0.618, 0.716), xytext=(0.500, 0.742),
                fontsize=7.0, color=PROT, fontweight="bold", ha="right", va="center",
                arrowprops=dict(arrowstyle="-", lw=0.55, color=PROT, shrinkA=2, shrinkB=1))
    ax.text(0.492, 0.684, "type II membrane protease", fontsize=6.3, color=PROT,
            ha="right", va="center")

    # ligand label sits directly above the molecule
    ax.text(1.010, 0.912, "68Ga-FAPI", fontsize=7.6, color=AMBER,
            ha="center", va="center", fontweight="bold")
    
    OUT.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(OUT, dpi=300, facecolor="white")
    plt.close(fig)
    print(f"  {OUT.name:26s} {OUT.stat().st_size // 1024:4d} KB")


if __name__ == "__main__":
    draw()
