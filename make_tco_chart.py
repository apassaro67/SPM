"""Generate a clean TCO comparison chart based on the 'Auswertung' tab
of the uploaded Excel."""
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch

# Figures from Auswertung tab (5-year TCO)
SN  = 4_775_750
BOB = 7_022_175
DELTA = BOB - SN  # 2,246,425
SN_Y3  = 971_150
BOB_Y3 = 1_316_435

# Colors aligned with deck
NAVY     = "#2C3E50"
DARK     = "#1F2937"
ORANGE   = "#F07F12"
GREY_TXT = "#4B5563"
GREEN_OK = "#2E7D32"
RED_HL   = "#B91C1C"
LIGHT_BG = "#F7F7F7"

# 5.83 x 3.17 inches target → use exact dimensions, dpi=200
fig = plt.figure(figsize=(5.83, 3.17), dpi=200)

# Two panes: left = 5y TCO comparison, right = recurring annual costs
gs = fig.add_gridspec(1, 2, width_ratios=[2.0, 1.0], wspace=0.35,
                      left=0.07, right=0.97, top=0.88, bottom=0.16)

# ---------- LEFT: 5-Year TCO comparison ----------
ax = fig.add_subplot(gs[0])
ax.set_title("5-Jahres-TCO  (Year 0–5)", fontsize=10, fontweight="bold",
             color=DARK, pad=6, loc="left")

labels = ["ServiceNow\n(SPM-Plattform)", "Best-of-Breed\nEinzeltools"]
vals = [SN/1e6, BOB/1e6]
colors = [ORANGE, NAVY]
bars = ax.bar(labels, vals, color=colors, width=0.55, edgecolor="white")

for b, v in zip(bars, vals):
    ax.text(b.get_x() + b.get_width()/2, v + 0.18, f"{v:.2f} M€",
            ha="center", va="bottom", fontsize=10.5, fontweight="bold",
            color=DARK)

# Delta annotation between the two bars
xc = 0.5
ax.annotate("", xy=(1, vals[1]*0.50), xytext=(0, vals[1]*0.50),
            arrowprops=dict(arrowstyle="<->", color=GREEN_OK, lw=1.8))
ax.text(xc, vals[1]*0.50 + 0.18,
        f"Ersparnis  −{DELTA/1e6:.2f} M€\n( ≈ {DELTA/BOB*100:.0f} % )",
        ha="center", va="bottom", fontsize=9.5, color=GREEN_OK,
        fontweight="bold")

ax.set_ylim(0, max(vals) * 1.28)
ax.set_ylabel("Mio. €", fontsize=9, color=GREY_TXT)
ax.tick_params(axis="x", labelsize=9, colors=DARK)
ax.tick_params(axis="y", labelsize=8, colors=GREY_TXT)
for s in ("top", "right"):
    ax.spines[s].set_visible(False)
ax.spines["left"].set_color("#D1D5DB")
ax.spines["bottom"].set_color("#D1D5DB")
ax.grid(axis="y", color="#E5E7EB", linewidth=0.5)
ax.set_axisbelow(True)

# ---------- RIGHT: Annual recurring (from year 3) ----------
ax2 = fig.add_subplot(gs[1])
ax2.set_title("Laufend p.a.\nab Jahr 3", fontsize=10, fontweight="bold",
              color=DARK, pad=6, loc="left")
labels2 = ["SN", "BoB"]
vals2 = [SN_Y3/1e3, BOB_Y3/1e3]
bars2 = ax2.bar(labels2, vals2, color=[ORANGE, NAVY], width=0.5,
                edgecolor="white")
for b, v in zip(bars2, vals2):
    ax2.text(b.get_x() + b.get_width()/2, v + 25, f"{v:.0f} k€",
             ha="center", va="bottom", fontsize=9.5, fontweight="bold",
             color=DARK)

delta_yr = (BOB_Y3 - SN_Y3)
ax2.text(0.5, max(vals2)*1.16,
         f"−{delta_yr/1e3:.0f} k€ p.a.",
         transform=ax2.get_xaxis_transform(),
         ha="center", va="bottom", fontsize=9, color=GREEN_OK, fontweight="bold")

ax2.set_ylim(0, max(vals2) * 1.30)
ax2.set_ylabel("Tsd. €", fontsize=8.5, color=GREY_TXT)
ax2.tick_params(axis="x", labelsize=9, colors=DARK)
ax2.tick_params(axis="y", labelsize=8, colors=GREY_TXT)
for s in ("top", "right"):
    ax2.spines[s].set_visible(False)
ax2.spines["left"].set_color("#D1D5DB")
ax2.spines["bottom"].set_color("#D1D5DB")
ax2.grid(axis="y", color="#E5E7EB", linewidth=0.5)
ax2.set_axisbelow(True)

# Footer caption
fig.text(0.5, 0.025,
         "Quelle: TCO_Vergleich_5J_SN_vs_BoB.xlsx — Tab 'Auswertung'   |   Werte basieren auf Annahmen / Schätzungen",
         ha="center", fontsize=6.5, color=GREY_TXT, style="italic")

plt.savefig("/tmp/tco_chart.png", dpi=200, facecolor="white")
print("Saved chart: /tmp/tco_chart.png")
