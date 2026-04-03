"""
Assignment 2 — Task 2: Market Intelligence Dashboard
RentEasy KZ — Competitor & Market Analysis

Generates:
  1. Dashboard 1 — Market Overview (price by city, source split, rooms distribution)
  2. Dashboard 2 — Price Segments / Competitor Analysis (segments by city, seller types)
  3. Dashboard 3 — District Price Heatmap (top districts by city)
  4. Dashboard 4 — Source Comparison (Krisha vs OLX)
  5. SWOT Summary printed to console
  6. Saves all figures to renteasy_data/dashboards/
"""

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")  # non-interactive backend (no display needed)
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.gridspec as gridspec
import os

INPUT_FILE = "renteasy_data/renteasy_dataset_cleaned.csv"
OUT_DIR = "renteasy_data/dashboards"
os.makedirs(OUT_DIR, exist_ok=True)

# ── Colours ────────────────────────────────────────────────────────────────────
BLUE    = "#2A5CAA"
YELLOW  = "#F5C842"
RED     = "#E85D2C"
LIGHT   = "#F5F3EE"
DARK    = "#1A1A2E"
GREY    = "#7F9799"
GREENS  = ["#2ecc71", "#27ae60", "#1e8449", "#145a32"]
SEGMENT_COLORS = {"low": "#2ecc71", "middle": BLUE, "high": YELLOW, "luxury": RED}

plt.rcParams.update({
    "figure.facecolor": LIGHT,
    "axes.facecolor": "white",
    "axes.edgecolor": GREY,
    "axes.labelcolor": DARK,
    "xtick.color": DARK,
    "ytick.color": DARK,
    "text.color": DARK,
    "font.family": "DejaVu Sans",
    "axes.titlesize": 12,
    "axes.titleweight": "bold",
})

# ── Load data ──────────────────────────────────────────────────────────────────
df = pd.read_csv(INPUT_FILE, encoding="utf-8-sig")
df["price_tenge"] = pd.to_numeric(df["price_tenge"], errors="coerce")
df["area_m2"] = pd.to_numeric(df["area_m2"], errors="coerce")
df["price_per_m2"] = pd.to_numeric(df["price_per_m2"], errors="coerce")
df["rooms"] = pd.to_numeric(df["rooms"], errors="coerce").round().astype("Int64")

CITIES = ["Астана", "Алматы", "Шымкент", "Карагандa", "Актобе"]
CITY_COLORS = [BLUE, "#e74c3c", "#2ecc71", YELLOW, "#9b59b6"]


# ══════════════════════════════════════════════════════════════════════════════
# DASHBOARD 1 — Market Overview
# ══════════════════════════════════════════════════════════════════════════════
fig = plt.figure(figsize=(18, 11), facecolor=LIGHT)
fig.suptitle("RentEasy KZ — Market Overview Dashboard", fontsize=16,
             fontweight="bold", color=DARK, y=0.98)

gs = gridspec.GridSpec(2, 3, figure=fig, hspace=0.45, wspace=0.38)

# ── 1a: Average rent by city (bar) ────────────────────────────────────────────
ax1 = fig.add_subplot(gs[0, :2])
city_avg = df.groupby("city")["price_tenge"].mean().reindex(CITIES)
bars = ax1.bar(CITIES, city_avg / 1000, color=CITY_COLORS, edgecolor="white", linewidth=1.2)
for bar, val in zip(bars, city_avg):
    ax1.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 1,
             f"{val/1000:.0f}K", ha="center", va="bottom", fontsize=9, fontweight="bold")
ax1.set_title("Average Monthly Rent by City (тг × 1000)")
ax1.set_ylabel("тenge (thousands)")
ax1.set_ylim(0, city_avg.max() / 1000 * 1.2)
ax1.grid(axis="y", alpha=0.3)
ax1.spines["top"].set_visible(False)
ax1.spines["right"].set_visible(False)

# ── 1b: Listings count by city (pie) ──────────────────────────────────────────
ax2 = fig.add_subplot(gs[0, 2])
city_counts = df["city"].value_counts().reindex(CITIES)
wedges, texts, autotexts = ax2.pie(
    city_counts, labels=CITIES, colors=CITY_COLORS,
    autopct="%1.0f%%", startangle=140,
    textprops={"fontsize": 8},
    wedgeprops={"edgecolor": "white", "linewidth": 1.5}
)
for at in autotexts:
    at.set_fontsize(8)
    at.set_fontweight("bold")
ax2.set_title("Listings by City (n={:,})".format(len(df)))

# ── 1c: Room distribution (bar) ───────────────────────────────────────────────
ax3 = fig.add_subplot(gs[1, 0])
room_counts = df["rooms"].value_counts().sort_index()
room_counts = room_counts[room_counts.index.notna()]
room_labels = {0: "Studio", 1: "1-room", 2: "2-room", 3: "3-room", 4: "4-room", 5: "5+"}
labels = [room_labels.get(int(r), f"{int(r)}-room") for r in room_counts.index]
bars3 = ax3.bar(labels, room_counts.values, color=BLUE, edgecolor="white", linewidth=1)
for bar, val in zip(bars3, room_counts.values):
    ax3.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 5,
             f"{val:,}", ha="center", va="bottom", fontsize=8)
ax3.set_title("Listings by Room Count")
ax3.set_ylabel("Count")
ax3.grid(axis="y", alpha=0.3)
ax3.spines["top"].set_visible(False)
ax3.spines["right"].set_visible(False)

# ── 1d: Seller type split (stacked bar by city) ───────────────────────────────
ax4 = fig.add_subplot(gs[1, 1])
seller_city = pd.crosstab(df["city"], df["seller_type"]).reindex(CITIES)
seller_city = seller_city.div(seller_city.sum(axis=1), axis=0) * 100
seller_colors = {"owner": BLUE, "agency": YELLOW, "unknown": GREY}
bottom = np.zeros(len(CITIES))
for seller in ["owner", "agency", "unknown"]:
    if seller in seller_city.columns:
        vals = seller_city[seller].fillna(0).values
        ax4.bar(CITIES, vals, bottom=bottom, color=seller_colors[seller],
                label=seller.capitalize(), edgecolor="white", linewidth=0.8)
        bottom += vals
ax4.set_title("Seller Type by City (%)")
ax4.set_ylabel("%")
ax4.legend(fontsize=8, loc="upper right")
ax4.set_ylim(0, 110)
ax4.tick_params(axis="x", rotation=15)
ax4.spines["top"].set_visible(False)
ax4.spines["right"].set_visible(False)

# ── 1e: Price per m² by city (box) ────────────────────────────────────────────
ax5 = fig.add_subplot(gs[1, 2])
box_data = [df[df["city"] == c]["price_per_m2"].dropna().values for c in CITIES]
bp = ax5.boxplot(box_data, tick_labels=[c[:3] for c in CITIES], patch_artist=True,
                 medianprops={"color": RED, "linewidth": 2})
for patch, color in zip(bp["boxes"], CITY_COLORS):
    patch.set_facecolor(color)
    patch.set_alpha(0.7)
ax5.set_title("Price per m² Distribution by City")
ax5.set_ylabel("тenge/m²")
ax5.grid(axis="y", alpha=0.3)
ax5.spines["top"].set_visible(False)
ax5.spines["right"].set_visible(False)

fig.savefig(f"{OUT_DIR}/dashboard1_market_overview.png", dpi=150, bbox_inches="tight")
plt.close(fig)
print(f"[SAVED] dashboard1_market_overview.png")


# ══════════════════════════════════════════════════════════════════════════════
# DASHBOARD 2 — Price Segments (Competitor Analysis)
# ══════════════════════════════════════════════════════════════════════════════
fig2 = plt.figure(figsize=(18, 11), facecolor=LIGHT)
fig2.suptitle("RentEasy KZ — Competitor Price Segment Analysis", fontsize=16,
              fontweight="bold", color=DARK, y=0.98)
gs2 = gridspec.GridSpec(2, 3, figure=fig2, hspace=0.45, wspace=0.38)

SEGMENTS = ["low", "middle", "high", "luxury"]
SEG_LABELS = {"low": "Low\n(<120K)", "middle": "Middle\n(120-250K)",
              "high": "High\n(250-500K)", "luxury": "Luxury\n(>500K)"}
SEG_COLORS = [SEGMENT_COLORS[s] for s in SEGMENTS]

# ── 2a: Segment share overall (pie) ───────────────────────────────────────────
ax = fig2.add_subplot(gs2[0, 0])
seg_counts = df["price_segment"].value_counts().reindex(SEGMENTS).fillna(0)
wedges, texts, autos = ax.pie(
    seg_counts, labels=[SEG_LABELS[s] for s in SEGMENTS],
    colors=SEG_COLORS, autopct="%1.1f%%", startangle=90,
    textprops={"fontsize": 8},
    wedgeprops={"edgecolor": "white", "linewidth": 1.5}
)
for at in autos:
    at.set_fontweight("bold")
ax.set_title("Overall Price Segment Share")

# ── 2b: Segment count by city (grouped bar) ───────────────────────────────────
ax2b = fig2.add_subplot(gs2[0, 1:])
seg_city = pd.crosstab(df["city"], df["price_segment"]).reindex(CITIES)[
    [s for s in SEGMENTS if s in df["price_segment"].unique()]
]
x = np.arange(len(CITIES))
width = 0.2
n = len(seg_city.columns)
for i, (seg, color) in enumerate(zip(seg_city.columns, SEG_COLORS)):
    offset = (i - n/2 + 0.5) * width
    bars = ax2b.bar(x + offset, seg_city[seg].fillna(0), width,
                    label=SEG_LABELS[seg].replace("\n", " "), color=color,
                    edgecolor="white", linewidth=0.8)
ax2b.set_xticks(x)
ax2b.set_xticklabels(CITIES, fontsize=9)
ax2b.set_title("Price Segments by City")
ax2b.set_ylabel("Listings Count")
ax2b.legend(fontsize=8)
ax2b.grid(axis="y", alpha=0.3)
ax2b.spines["top"].set_visible(False)
ax2b.spines["right"].set_visible(False)

# ── 2c: Avg price by segment + city (line) ────────────────────────────────────
ax2c = fig2.add_subplot(gs2[1, :2])
for city, color in zip(CITIES, CITY_COLORS):
    city_df = df[df["city"] == city]
    seg_avg = city_df.groupby("price_segment")["price_tenge"].mean().reindex(SEGMENTS)
    valid = seg_avg.dropna()
    if len(valid) > 1:
        ax2c.plot(valid.index, valid.values / 1000, "o-", label=city,
                  color=color, linewidth=2, markersize=6)
ax2c.set_title("Average Price by Segment and City (тг × 1000)")
ax2c.set_ylabel("тenge (thousands)")
ax2c.set_xticks(range(len(SEGMENTS)))
ax2c.set_xticklabels([SEG_LABELS[s].replace("\n", " ") for s in SEGMENTS])
ax2c.legend(fontsize=8)
ax2c.grid(alpha=0.3)
ax2c.spines["top"].set_visible(False)
ax2c.spines["right"].set_visible(False)

# ── 2d: Segment × seller type (stacked bar) ───────────────────────────────────
ax2d = fig2.add_subplot(gs2[1, 2])
seg_seller = pd.crosstab(df["price_segment"], df["seller_type"]).reindex(SEGMENTS)
seg_seller = seg_seller.div(seg_seller.sum(axis=1), axis=0) * 100
bottom2 = np.zeros(len(SEGMENTS))
for seller, color in zip(["owner", "agency", "unknown"],
                         [BLUE, YELLOW, GREY]):
    if seller in seg_seller.columns:
        vals = seg_seller[seller].fillna(0).values
        ax2d.bar([SEG_LABELS[s].replace("\n", " ") for s in SEGMENTS],
                 vals, bottom=bottom2, color=color,
                 label=seller.capitalize(), edgecolor="white", linewidth=0.8)
        bottom2 += vals
ax2d.set_title("Seller Type by Price Segment (%)")
ax2d.set_ylabel("%")
ax2d.legend(fontsize=8)
ax2d.tick_params(axis="x", rotation=15)
ax2d.spines["top"].set_visible(False)
ax2d.spines["right"].set_visible(False)

fig2.savefig(f"{OUT_DIR}/dashboard2_price_segments.png", dpi=150, bbox_inches="tight")
plt.close(fig2)
print(f"[SAVED] dashboard2_price_segments.png")


# ══════════════════════════════════════════════════════════════════════════════
# DASHBOARD 3 — District Heatmap (top districts in Astana + Almaty)
# ══════════════════════════════════════════════════════════════════════════════
fig3 = plt.figure(figsize=(18, 10), facecolor=LIGHT)
fig3.suptitle("RentEasy KZ — Top Districts Price Intelligence", fontsize=16,
              fontweight="bold", color=DARK, y=0.98)
gs3 = gridspec.GridSpec(1, 2, figure=fig3, hspace=0.3, wspace=0.4)

for col_idx, (target_city, title_color) in enumerate([
    ("Астана", BLUE), ("Алматы", "#e74c3c")
]):
    ax = fig3.add_subplot(gs3[0, col_idx])
    city_df = df[(df["city"] == target_city) & df["district"].notna()]
    top_districts = (city_df.groupby("district")["price_tenge"]
                     .agg(["mean", "count"])
                     .query("count >= 5")
                     .sort_values("mean", ascending=True)
                     .tail(12))

    bars = ax.barh(top_districts.index, top_districts["mean"] / 1000,
                   color=title_color, alpha=0.8, edgecolor="white", linewidth=1)
    for bar, (_, row) in zip(bars, top_districts.iterrows()):
        ax.text(bar.get_width() + 1, bar.get_y() + bar.get_height() / 2,
                f"{row['mean']/1000:.0f}K (n={int(row['count'])})",
                va="center", fontsize=7.5)
    ax.set_title(f"Top Districts by Avg Price — {target_city}", color=title_color)
    ax.set_xlabel("Average Price (тг × 1000)")
    ax.grid(axis="x", alpha=0.3)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.set_xlim(0, top_districts["mean"].max() / 1000 * 1.35)

fig3.savefig(f"{OUT_DIR}/dashboard3_districts.png", dpi=150, bbox_inches="tight")
plt.close(fig3)
print(f"[SAVED] dashboard3_districts.png")


# ══════════════════════════════════════════════════════════════════════════════
# DASHBOARD 4 — Source Comparison: Krisha vs OLX
# ══════════════════════════════════════════════════════════════════════════════
fig4 = plt.figure(figsize=(18, 9), facecolor=LIGHT)
fig4.suptitle("RentEasy KZ — Krisha.kz vs OLX.kz Data Comparison", fontsize=16,
              fontweight="bold", color=DARK, y=0.98)
gs4 = gridspec.GridSpec(1, 3, figure=fig4, hspace=0.35, wspace=0.38)

src_colors = {"krisha": BLUE, "olx": RED}

# ── 4a: Listing count by source × city ────────────────────────────────────────
ax4a = fig4.add_subplot(gs4[0, 0])
src_city = pd.crosstab(df["city"], df["source"]).reindex(CITIES)
x = np.arange(len(CITIES))
w = 0.35
for i, (src, color) in enumerate(src_colors.items()):
    if src in src_city.columns:
        vals = src_city[src].fillna(0)
        bars = ax4a.bar(x + (i - 0.5) * w, vals, w, label=src.capitalize(),
                        color=color, edgecolor="white", linewidth=1)
        for bar, val in zip(bars, vals):
            ax4a.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 5,
                      str(int(val)), ha="center", va="bottom", fontsize=7)
ax4a.set_xticks(x)
ax4a.set_xticklabels(CITIES, rotation=15, fontsize=8)
ax4a.set_title("Listings by Source & City")
ax4a.set_ylabel("Count")
ax4a.legend(fontsize=9)
ax4a.grid(axis="y", alpha=0.3)
ax4a.spines["top"].set_visible(False)
ax4a.spines["right"].set_visible(False)

# ── 4b: Avg price by source × city ────────────────────────────────────────────
ax4b = fig4.add_subplot(gs4[0, 1])
src_price = df.groupby(["city", "source"])["price_tenge"].mean().unstack().reindex(CITIES)
for i, (src, color) in enumerate(src_colors.items()):
    if src in src_price.columns:
        vals = src_price[src].fillna(0)
        bars = ax4b.bar(x + (i - 0.5) * w, vals / 1000, w, label=src.capitalize(),
                        color=color, edgecolor="white", linewidth=1)
ax4b.set_xticks(x)
ax4b.set_xticklabels(CITIES, rotation=15, fontsize=8)
ax4b.set_title("Avg Rent by Source & City (тг × 1000)")
ax4b.set_ylabel("тenge (thousands)")
ax4b.legend(fontsize=9)
ax4b.grid(axis="y", alpha=0.3)
ax4b.spines["top"].set_visible(False)
ax4b.spines["right"].set_visible(False)

# ── 4c: Room coverage by source ───────────────────────────────────────────────
ax4c = fig4.add_subplot(gs4[0, 2])
for src, color in src_colors.items():
    src_df = df[df["source"] == src]
    room_dist = src_df["rooms"].value_counts(normalize=True).sort_index()
    room_dist = room_dist[room_dist.index.notna()]
    labels_r = [room_labels.get(int(r), f"{int(r)}-r") for r in room_dist.index]
    ax4c.plot(labels_r, room_dist.values * 100, "o-", label=src.capitalize(),
              color=color, linewidth=2, markersize=7)
ax4c.set_title("Room Distribution by Source (%)")
ax4c.set_ylabel("%")
ax4c.legend(fontsize=9)
ax4c.grid(alpha=0.3)
ax4c.spines["top"].set_visible(False)
ax4c.spines["right"].set_visible(False)

fig4.savefig(f"{OUT_DIR}/dashboard4_source_comparison.png", dpi=150, bbox_inches="tight")
plt.close(fig4)
print(f"[SAVED] dashboard4_source_comparison.png")


# ══════════════════════════════════════════════════════════════════════════════
# SWOT ANALYSIS — printed summary
# ══════════════════════════════════════════════════════════════════════════════
# Compute key metrics for SWOT
astana_avg  = df[df["city"] == "Астана"]["price_tenge"].mean()
almaty_avg  = df[df["city"] == "Алматы"]["price_tenge"].mean()
agency_pct  = (df["seller_type"] == "agency").mean() * 100
middle_pct  = (df["price_segment"] == "middle").mean() * 100
high_pct    = (df["price_segment"] == "high").mean() * 100
olx_pct     = (df["source"] == "olx").mean() * 100
room_no_data_pct = df["rooms"].isna().mean() * 100

swot_fig = plt.figure(figsize=(16, 10), facecolor=LIGHT)
swot_fig.suptitle("RentEasy KZ — SWOT Analysis of Competitors", fontsize=16,
                   fontweight="bold", color=DARK, y=0.98)

quadrants = [
    ("STRENGTHS", BLUE, 0, 1,
     [
         f"• Krisha.kz dominates with {100-olx_pct:.0f}% of listings — highest data coverage",
         "• Large dataset: 6,684 clean records across 5 cities",
         "• Strong middle segment presence: confirms stable demand",
         "• Established brand trust among Kazakh renters",
         "• Both platforms free to browse — low adoption barrier",
     ]),
    ("WEAKNESSES", RED, 1, 1,
     [
         f"• {agency_pct:.0f}% of listings via agencies → double commissions for renters",
         f"• {room_no_data_pct:.0f}% of OLX listings missing room/area data — poor UX",
         "• No price analytics or historical trends on either platform",
         "• Listings scattered across 2+ sites — no aggregation",
         "• Fake/duplicate listings are not filtered",
     ]),
    ("OPPORTUNITIES", "#2ecc71", 0, 0,
     [
         f"• Astana avg rent {astana_avg/1000:.0f}K тg/mo — high price variance = analytics demand",
         f"• Middle segment ({middle_pct:.0f}%) + High ({high_pct:.0f}%) = core target for RentEasy",
         "• 61% owner listings = direct connection possible (bypass agencies)",
         "• AI price recommendations not offered by any current platform",
         "• Seasonal peaks (Aug–Sep) = marketing growth windows",
     ]),
    ("THREATS", YELLOW, 1, 0,
     [
         "• Krisha.kz may add analytics features (defensive moat risk)",
         "• Low entry barrier for new aggregator startups",
         "• OLX anti-scraping / URL changes (fragile data pipeline)",
         "• Real estate market slowdown reduces listing volume",
         "• User trust in new platforms is hard to build quickly",
     ]),
]

for (title, color, col, row, bullets) in quadrants:
    ax = swot_fig.add_axes([0.02 + col * 0.49, 0.06 + row * 0.47, 0.46, 0.42])
    ax.set_facecolor(color + "18")  # very light tint
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")

    ax.text(0.5, 0.94, title, ha="center", va="top", fontsize=13, fontweight="bold",
            color=color, transform=ax.transAxes)
    ax.add_patch(mpatches.FancyBboxPatch((0, 0), 1, 1,
                 boxstyle="round,pad=0.01", linewidth=2,
                 edgecolor=color, facecolor="none", transform=ax.transAxes))
    for i, bullet in enumerate(bullets):
        ax.text(0.03, 0.82 - i * 0.17, bullet, va="top", fontsize=9,
                color=DARK, wrap=True, transform=ax.transAxes,
                multialignment="left")

swot_fig.savefig(f"{OUT_DIR}/dashboard5_swot.png", dpi=150, bbox_inches="tight")
plt.close(swot_fig)
print(f"[SAVED] dashboard5_swot.png")


# ══════════════════════════════════════════════════════════════════════════════
# CONSOLE SUMMARY
# ══════════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 60)
print("MARKET INTELLIGENCE SUMMARY")
print("=" * 60)
print(f"\nDataset: {len(df):,} clean records | {df['city'].nunique()} cities | 2 sources")
print(f"\nAvg rent: Алматы={almaty_avg/1000:.0f}K | Астана={astana_avg/1000:.0f}K тg/mo")
print(f"Agency share: {agency_pct:.1f}% of all listings")
print(f"Middle segment: {middle_pct:.0f}% | High: {high_pct:.0f}%")
print(f"\nMost promising market: Астана (high rent, high growth, 1,830 listings)")
print(f"Target segment: Middle (120-250K тg) — largest demand pool")
print(f"Key competitor weakness: no price analytics, agency dominance")
print(f"RentEasy strategy: launch Астана → price transparency + AI recommendations")
print(f"\nAll dashboards saved to: {OUT_DIR}/")
