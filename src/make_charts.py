"""Charts for the Lumentum pitch, Sprouts-style greens."""
from pathlib import Path
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import date

OUT = Path(r"C:\Users\Dell\Desktop\uchicagoaicomp\final_slideshow\charts")
OUT.mkdir(parents=True, exist_ok=True)

GREEN = "#2F7A38"
GREEN2 = "#66A36B"
DARK = "#1B4D1E"
GOLD = "#C9A227"
RED = "#C62828"

# LITE closes from stockanalysis.com, accessed 18 Sep 2026
dates = [
    date(2026, 9, 8), date(2026, 9, 9), date(2026, 9, 10), date(2026, 9, 11),
    date(2026, 9, 14), date(2026, 9, 15), date(2026, 9, 16), date(2026, 9, 17),
]
px = [978.54, 988.98, 935.70, 927.03, 835.03, 838.96, 919.40, 893.61]

fig, ax = plt.subplots(figsize=(7.2, 3.15), dpi=180)
ax.plot(dates, px, color=GREEN, lw=2.2, marker="o", ms=5)
ax.axvspan(date(2026, 9, 12), date(2026, 9, 13), color="#C8E6C9", alpha=0.7, label="Amodei essay weekend")
ax.axhline(893.61, color=GOLD, ls="--", lw=1, label="17 Sep close $893.61")
ax.annotate("14 Sep −9.9%\nnarrative shock", xy=(date(2026, 9, 14), 835.03),
            xytext=(date(2026, 9, 8), 820), fontsize=8, color=RED,
            arrowprops=dict(arrowstyle="->", color=RED, lw=0.8))
ax.annotate("16 Sep recovered\nto $919", xy=(date(2026, 9, 16), 919.40),
            xytext=(date(2026, 9, 16), 980), fontsize=8, color=DARK,
            arrowprops=dict(arrowstyle="->", color=DARK, lw=0.8))
ax.set_title("LITE: one-day scare, not a capex revision", color=DARK, loc="left", fontsize=11, fontweight="bold")
ax.set_ylabel("USD / share", color=DARK, fontsize=8)
ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %d"))
ax.tick_params(labelsize=8, colors=DARK)
for spine in ax.spines.values():
    spine.set_color("#BDBDBD")
ax.set_ylim(800, 1030)
ax.legend(fontsize=7, frameon=False)
ax.grid(axis="y", color="#EEEEEE")
fig.tight_layout()
fig.savefig(OUT / "lite_tape.png", bbox_inches="tight", facecolor="white")
plt.close()

# Peer cash conversion
names = ["LITE", "COHR", "FN"]
cfo = [751.4, 79.5, 256.7]
capex = [451.3, 1102.9, 252.5]
fcf = [300.1, -1023.4, 4.2]
x = range(len(names))
fig, ax = plt.subplots(figsize=(6.4, 3.2), dpi=180)
w = 0.25
ax.bar([i - w for i in x], cfo, w, label="FY26 CFO", color=GREEN)
ax.bar(x, capex, w, label="FY26 cash capex", color=GREEN2)
ax.bar([i + w for i in x], fcf, w, label="CFO − capex", color=GOLD)
ax.axhline(0, color="#888", lw=0.6)
ax.set_xticks(list(x))
ax.set_xticklabels(names, fontsize=10, color=DARK, fontweight="bold")
ax.set_ylabel("USD millions", color=DARK, fontsize=8)
ax.set_title("LITE converts the capex cycle; COHR does not", color=DARK, loc="left", fontsize=11, fontweight="bold")
ax.legend(fontsize=7, frameon=False)
ax.tick_params(labelsize=8, colors=DARK)
for spine in ax.spines.values():
    spine.set_color("#BDBDBD")
ax.grid(axis="y", color="#EEEEEE")
fig.tight_layout()
fig.savefig(OUT / "peer_fcf.png", bbox_inches="tight", facecolor="white")
plt.close()
print("charts written", OUT)
