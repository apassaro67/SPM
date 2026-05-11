"""TCO chart v2: left = absolute 5-year comparison, right = percent/delta KPI."""
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch

SN  = 4_775_750
BOB = 7_022_175
DELTA = BOB - SN          # 2,246,425
PCT = DELTA / BOB * 100   # 32.0
SN_Y3  = 971_150
BOB_Y3 = 1_316_435
DELTA_Y3 = BOB_Y3 - SN_Y3
PCT_Y3 = DELTA_Y3 / BOB_Y3 * 100

NAVY     = "#2C3E50"
DARK     = "#1F2937"
ORANGE   = "#F07F12"
GREY_TXT = "#4B5563"
GREEN_OK = "#2E7D32"
LIGHT_BG = "#F7F7F7"

fig = plt.figure(figsize=(5.83, 3.17), dpi=200)
gs = fig.add_gridspec(1, 2, width_ratios=[1.8, 1.0], wspace=0.25,
                      left=0.08, right=0.93, top=0.86, bottom=0.16)

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
            ha="center", va="bottom", fontsize=11, fontweight="bold", color=DARK)

# Delta annotation
ax.annotate("", xy=(1, vals[1]*0.50), xytext=(0, vals[1]*0.50),
            arrowprops=dict(arrowstyle="<->", color=GREEN_OK, lw=1.8))
ax.text(0.5, vals[1]*0.50 + 0.20,
        f"Ersparnis  −{DELTA/1e6:.2f} M€\n( −{PCT:.0f} % )",
        ha="center", va="bottom", fontsize=10, color=GREEN_OK, fontweight="bold")

ax.set_ylim(0, max(vals) * 1.28)
ax.set_ylabel("Mio. €", fontsize=9, color=GREY_TXT)
ax.tick_params(axis="x", labelsize=9, colors=DARK)
ax.tick_params(axis="y", labelsize=8, colors=GREY_TXT)
for s in ("top", "right"):
    ax.spines[s].set_visible(False)
ax.spines["left"].set_color("#D1D5DB"); ax.spines["bottom"].set_color("#D1D5DB")
ax.grid(axis="y", color="#E5E7EB", linewidth=0.5); ax.set_axisbelow(True)

# ---------- RIGHT: KPI tile, percent + delta ----------
ax2 = fig.add_subplot(gs[1])
ax2.set_title("Laufend p.a.\nab Jahr 3", fontsize=10, fontweight="bold",
              color=DARK, pad=6, loc="left")
ax2.axis("off")

# Big % number
ax2.text(0.5, 0.62, f"−{PCT_Y3:.0f} %", ha="center", va="center",
         fontsize=36, fontweight="bold", color=GREEN_OK,
         transform=ax2.transAxes)
# Sub-label with absolute delta
ax2.text(0.5, 0.30, f"= −{DELTA_Y3/1e3:.0f} k€ p.a.",
         ha="center", va="center", fontsize=11.5, fontweight="bold",
         color=DARK, transform=ax2.transAxes)
ax2.text(0.5, 0.18, "weniger Betriebskosten",
         ha="center", va="center", fontsize=8.5, style="italic",
         color=GREY_TXT, transform=ax2.transAxes)
# Surrounding rounded rectangle
ax2.add_patch(FancyBboxPatch((0.03, 0.05), 0.94, 0.85,
                             boxstyle="round,pad=0.01,rounding_size=0.04",
                             facecolor="white", edgecolor=GREEN_OK, lw=1.8,
                             transform=ax2.transAxes))

# Footer caption
fig.text(0.5, 0.025,
         "Quelle: TCO_Vergleich_5J_SN_vs_BoB.xlsx — Tab 'Auswertung'   |   Werte basieren auf Annahmen / Schätzungen",
         ha="center", fontsize=6.5, color=GREY_TXT, style="italic")

plt.savefig("/tmp/tco_chart.png", dpi=200, facecolor="white")
print("Saved")
