#!/usr/bin/env python3
"""Render clean flat-style Patna traffic/parking illustrations as PNGs for the PPT."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle, FancyBboxPatch, Circle, Polygon
import os

os.makedirs("ppt_images", exist_ok=True)


def car(ax, x, y, w, h, color, flip=False):
    ax.add_patch(FancyBboxPatch((x, y), w, h*0.55, boxstyle="round,pad=0.02,rounding_size=0.05",
                                fc=color, ec="white", lw=1.5, zorder=5))
    # cabin
    cx = x + w*0.22
    ax.add_patch(FancyBboxPatch((cx, y+h*0.5), w*0.55, h*0.4, boxstyle="round,pad=0.02,rounding_size=0.05",
                                fc=color, ec="white", lw=1.2, zorder=5))
    # windows
    ax.add_patch(Rectangle((cx+w*0.05, y+h*0.55), w*0.45, h*0.28, fc="#cfe3f2", ec="none", zorder=6))
    # wheels
    ax.add_patch(Circle((x+w*0.25, y), h*0.13, fc="#222", ec="#111", zorder=7))
    ax.add_patch(Circle((x+w*0.75, y), h*0.13, fc="#222", ec="#111", zorder=7))


def scene(filename, title, subtitle, congested=True):
    fig, ax = plt.subplots(figsize=(6.2, 6.4), dpi=150)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 10)
    ax.axis("off")

    # sky
    ax.add_patch(Rectangle((0, 5.2), 10, 4.8, fc="#cfe0ee", ec="none"))
    # distant buildings
    bcolors = ["#9fb3c4", "#8ca6bb", "#adc0d0", "#94aabd", "#a6bccb", "#8fa8bd"]
    bx = 0.2
    import random
    random.seed(3)
    while bx < 9.8:
        bw = random.uniform(0.8, 1.4)
        bh = random.uniform(1.6, 3.2)
        ax.add_patch(Rectangle((bx, 5.2), bw, bh, fc=random.choice(bcolors), ec="#e8eef4", lw=0.8))
        # windows
        for wy in range(int(bh/0.5)):
            for wx in range(int(bw/0.35)):
                if random.random() > 0.35:
                    ax.add_patch(Rectangle((bx+0.12+wx*0.32, 5.35+wy*0.5), 0.16, 0.22,
                                           fc="#dfe9f1", ec="none"))
        bx += bw + 0.1

    # footpath
    ax.add_patch(Rectangle((0, 4.7), 10, 0.5, fc="#b9c3cc", ec="none"))
    # road
    ax.add_patch(Rectangle((0, 0), 10, 4.7, fc="#53595f", ec="none"))
    # lane markings (center)
    for lx in range(0, 10):
        ax.add_patch(Rectangle((lx+0.2, 2.3), 0.5, 0.08, fc="#f5d94e", ec="none"))

    # vehicles
    if congested:
        # haphazard, overlapping parking near footpath + a jammed lane
        colors = ["#F28C28", "#C0392B", "#1B6CA8", "#2E8B57", "#8E44AD", "#E67E22"]
        xs = [0.3, 1.7, 3.2, 4.5, 6.1, 7.6, 8.9]
        for i, xx in enumerate(xs):
            car(ax, xx, 3.7 + (0.15 if i % 2 else 0), 1.5, 1.0, colors[i % len(colors)])
        # second crowded row (double parking)
        xs2 = [0.9, 2.6, 4.2, 6.8, 8.3]
        for i, xx in enumerate(xs2):
            car(ax, xx, 1.2, 1.4, 0.95, colors[(i+2) % len(colors)])
    else:
        colors = ["#1B6CA8", "#2E8B57"]
        for i, xx in enumerate([1.5, 6.0]):
            car(ax, xx, 1.4, 1.6, 1.0, colors[i % len(colors)])

    # title banner
    ax.add_patch(FancyBboxPatch((0.4, 8.7), 9.2, 1.0, boxstyle="round,pad=0.05,rounding_size=0.1",
                                fc="#0B2E4F", ec="none", zorder=10))
    ax.text(5, 9.2, title, ha="center", va="center", fontsize=15, fontweight="bold",
            color="white", zorder=11)
    ax.text(5, 0.35, subtitle, ha="center", va="center", fontsize=10.5,
            color="#0B2E4F", fontweight="bold")

    plt.subplots_adjust(left=0, right=1, top=1, bottom=0)
    fig.savefig(f"ppt_images/{filename}", bbox_inches="tight", pad_inches=0.05)
    plt.close(fig)
    print("saved", filename)


scene("patna_parking.png", "PATNA — ROADSIDE PARKING",
      "Illegal roadside parking narrowing a busy road", congested=True)
scene("patna_congestion.png", "PATNA — TRAFFIC CONGESTION",
      "Mixed traffic & on-street parking at a choke point", congested=True)
print("done")
