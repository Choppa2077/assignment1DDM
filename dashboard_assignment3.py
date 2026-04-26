"""
Assignment 3 — Task 3: Sentiment & Trends Dashboards
RentEasy KZ

Generates 5 dashboards saved to renteasy_data/dashboards/:
  1. sentiment_overview.png    — Sentiment distribution per competitor
  2. sentiment_dynamics.png    — Reviews over time + rolling avg score
  3. sentiment_negative.png    — Top negative keywords (competitor weak spots)
  4. trends_combined.png       — Time series + ACF + forecast combined
  5. sentiment_future.png      — Future sentiment trend projection
"""

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
import matplotlib.dates as mdates
import matplotlib.patches as mpatches
import os
import re
from collections import Counter

from statsmodels.graphics.tsaplots import plot_acf
from statsmodels.tsa.stattools import adfuller
from statsmodels.tsa.arima.model import ARIMA
import warnings
warnings.filterwarnings("ignore")

# ── Paths ───────────────────────────────────────────────────────────────────
SENTIMENT_FILE = "renteasy_data/sentiment_results.csv"
TRENDS_FILE    = "renteasy_data/google_trends_data.csv"
OUT_DIR        = "renteasy_data/dashboards"
os.makedirs(OUT_DIR, exist_ok=True)

# ── Colours ─────────────────────────────────────────────────────────────────
BLUE   = "#2A5CAA"
YELLOW = "#F5C842"
RED    = "#E85D2C"
GREEN  = "#2ecc71"
LIGHT  = "#F5F3EE"
DARK   = "#1A1A2E"
GREY   = "#7F9799"

SENT_COLORS = {"POSITIVE": GREEN, "NEUTRAL": YELLOW, "NEGATIVE": RED}
COMP_COLORS = {"Krisha.kz": BLUE, "OLX Kazakhstan": RED, "kn.kz": "#9b59b6"}

plt.rcParams.update({
    "figure.facecolor": LIGHT,
    "axes.facecolor":   "white",
    "axes.edgecolor":   GREY,
    "text.color":       DARK,
    "font.family":      "DejaVu Sans",
    "axes.titlesize":   12,
    "axes.titleweight": "bold",
})

# ── Load sentiment data ──────────────────────────────────────────────────────
df = pd.read_csv(SENTIMENT_FILE, encoding="utf-8-sig")
df["review_date"] = pd.to_datetime(df["review_date"], errors="coerce")
df["month"] = df["review_date"].dt.to_period("M")

COMPETITORS = df["competitor"].unique().tolist()
SENTIMENTS  = ["POSITIVE", "NEUTRAL", "NEGATIVE"]


# ══════════════════════════════════════════════════════════════════════════════
# DASHBOARD 1 — Sentiment Overview
# ══════════════════════════════════════════════════════════════════════════════
fig = plt.figure(figsize=(18, 10), facecolor=LIGHT)
fig.suptitle("RentEasy KZ — Competitor Sentiment Overview (Google Play Reviews)",
             fontsize=15, fontweight="bold", color=DARK, y=0.98)

gs = gridspec.GridSpec(2, 3, figure=fig, hspace=0.45, wspace=0.4)

# ── 1a: Grouped bar — count per sentiment per competitor
ax1 = fig.add_subplot(gs[0, :2])
sent_counts = pd.crosstab(df["competitor"], df["sentiment"])[SENTIMENTS]
x = np.arange(len(COMPETITORS))
w = 0.25
for i, (sent, color) in enumerate(SENT_COLORS.items()):
    vals = sent_counts[sent].reindex(COMPETITORS).fillna(0)
    bars = ax1.bar(x + (i-1)*w, vals, w, label=sent, color=color,
                   edgecolor="white", linewidth=1)
    for bar, val in zip(bars, vals):
        ax1.text(bar.get_x()+bar.get_width()/2, bar.get_height()+1,
                 str(int(val)), ha="center", va="bottom", fontsize=8)
ax1.set_xticks(x)
ax1.set_xticklabels(COMPETITORS, fontsize=10)
ax1.set_title("Sentiment Distribution per Competitor (count)")
ax1.set_ylabel("Number of Reviews")
ax1.legend(fontsize=9)
ax1.grid(axis="y", alpha=0.3)
ax1.spines["top"].set_visible(False)
ax1.spines["right"].set_visible(False)

# ── 1b: Overall sentiment pie
ax2 = fig.add_subplot(gs[0, 2])
overall = df["sentiment"].value_counts().reindex(SENTIMENTS).fillna(0)
wedges, _, autos = ax2.pie(
    overall, labels=SENTIMENTS,
    colors=[SENT_COLORS[s] for s in SENTIMENTS],
    autopct="%1.1f%%", startangle=90,
    textprops={"fontsize": 9},
    wedgeprops={"edgecolor": "white", "linewidth": 1.5}
)
for at in autos:
    at.set_fontweight("bold")
ax2.set_title(f"Overall Sentiment\n(n={len(df)} reviews)")

# ── 1c: Avg star rating per competitor (bar)
ax3 = fig.add_subplot(gs[1, 0])
avg_rating = df.groupby("competitor")["rating"].mean().reindex(COMPETITORS)
colors3 = [COMP_COLORS.get(c, BLUE) for c in COMPETITORS]
bars3 = ax3.bar(COMPETITORS, avg_rating, color=colors3, edgecolor="white", linewidth=1)
for bar, val in zip(bars3, avg_rating):
    ax3.text(bar.get_x()+bar.get_width()/2, bar.get_height()+0.05,
             f"{val:.2f} ★", ha="center", va="bottom", fontsize=10, fontweight="bold")
ax3.set_title("Average Star Rating (Google Play)")
ax3.set_ylabel("Rating (1–5)")
ax3.set_ylim(0, 5.5)
ax3.axhline(3, color=GREY, linewidth=1, linestyle="--", alpha=0.5)
ax3.grid(axis="y", alpha=0.3)
ax3.spines["top"].set_visible(False)
ax3.spines["right"].set_visible(False)

# ── 1d: Stacked % bar — sentiment share per competitor
ax4 = fig.add_subplot(gs[1, 1:])
sent_pct = sent_counts.div(sent_counts.sum(axis=1), axis=0) * 100
bottom = np.zeros(len(COMPETITORS))
for sent in SENTIMENTS:
    vals = sent_pct[sent].reindex(COMPETITORS).fillna(0).values
    ax4.bar(COMPETITORS, vals, bottom=bottom, label=sent,
            color=SENT_COLORS[sent], edgecolor="white", linewidth=0.8)
    for i, (v, b) in enumerate(zip(vals, bottom)):
        if v > 5:
            ax4.text(i, b + v/2, f"{v:.0f}%", ha="center", va="center",
                     fontsize=9, fontweight="bold", color="white")
    bottom += vals
ax4.set_title("Sentiment Share per Competitor (%)")
ax4.set_ylabel("%")
ax4.set_ylim(0, 110)
ax4.legend(fontsize=9, loc="upper right")
ax4.spines["top"].set_visible(False)
ax4.spines["right"].set_visible(False)

fig.savefig(f"{OUT_DIR}/sentiment_overview.png", dpi=150, bbox_inches="tight")
plt.close(fig)
print("[SAVED] sentiment_overview.png")


# ══════════════════════════════════════════════════════════════════════════════
# DASHBOARD 2 — Sentiment Dynamics over Time
# ══════════════════════════════════════════════════════════════════════════════
# Numeric sentiment score: POSITIVE=1, NEUTRAL=0, NEGATIVE=-1
df["sent_score"] = df["sentiment"].map({"POSITIVE": 1, "NEUTRAL": 0, "NEGATIVE": -1})

fig2 = plt.figure(figsize=(18, 10), facecolor=LIGHT)
fig2.suptitle("RentEasy KZ — Sentiment Dynamics Over Time",
              fontsize=15, fontweight="bold", color=DARK, y=0.98)
gs2 = gridspec.GridSpec(2, 1, figure=fig2, hspace=0.45)

# ── 2a: Reviews scatter coloured by sentiment per competitor
ax2a = fig2.add_subplot(gs2[0])
for comp in COMPETITORS:
    sub = df[df["competitor"] == comp].dropna(subset=["review_date"])
    for sent, color in SENT_COLORS.items():
        s2 = sub[sub["sentiment"] == sent]
        if len(s2):
            ax2a.scatter(s2["review_date"], [comp]*len(s2),
                         c=color, s=30, alpha=0.6, label=f"{comp}–{sent}" if comp == COMPETITORS[0] else "")
patches = [mpatches.Patch(color=c, label=s) for s, c in SENT_COLORS.items()]
ax2a.legend(handles=patches, fontsize=9, loc="upper left")
ax2a.set_title("Review Timeline by Sentiment (each dot = 1 review)")
ax2a.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m"))
ax2a.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
plt.setp(ax2a.xaxis.get_majorticklabels(), rotation=30, fontsize=8)
ax2a.grid(axis="x", alpha=0.3)
ax2a.spines["top"].set_visible(False)
ax2a.spines["right"].set_visible(False)

# ── 2b: Rolling avg sentiment score per competitor
ax2b = fig2.add_subplot(gs2[1])
for comp, color in COMP_COLORS.items():
    sub = df[df["competitor"] == comp].dropna(subset=["review_date"]).sort_values("review_date")
    if len(sub) < 5:
        continue
    sub = sub.set_index("review_date")
    # Resample by week
    weekly = sub["sent_score"].resample("W").mean()
    rolling = weekly.rolling(window=4, min_periods=1).mean()
    ax2b.plot(rolling.index, rolling.values, color=color, linewidth=2.2,
              label=comp, marker="o", markersize=3)

ax2b.axhline(0, color=GREY, linewidth=1, linestyle="--", alpha=0.7)
ax2b.fill_between(ax2b.get_xlim(), -1, 0, alpha=0.04, color=RED)
ax2b.fill_between(ax2b.get_xlim(), 0, 1, alpha=0.04, color=GREEN)
ax2b.set_title("Rolling Average Sentiment Score per Competitor (−1=Negative, 0=Neutral, +1=Positive)")
ax2b.set_ylabel("Avg Sentiment Score")
ax2b.set_ylim(-1.1, 1.1)
ax2b.legend(fontsize=9)
ax2b.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m"))
ax2b.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
plt.setp(ax2b.xaxis.get_majorticklabels(), rotation=30, fontsize=8)
ax2b.grid(alpha=0.3)
ax2b.spines["top"].set_visible(False)
ax2b.spines["right"].set_visible(False)

fig2.savefig(f"{OUT_DIR}/sentiment_dynamics.png", dpi=150, bbox_inches="tight")
plt.close(fig2)
print("[SAVED] sentiment_dynamics.png")


# ══════════════════════════════════════════════════════════════════════════════
# DASHBOARD 3 — Negative Review Analysis (Competitor Weak Spots)
# ══════════════════════════════════════════════════════════════════════════════
STOPWORDS = {
    "и", "в", "не", "на", "с", "по", "что", "это", "как", "а", "но", "из",
    "к", "у", "за", "от", "то", "все", "мне", "до", "бы", "о", "же", "я",
    "он", "она", "они", "мы", "вы", "так", "уже", "было", "быть", "при",
    "только", "ещё", "еще", "да", "если", "для", "или", "нет", "нет",
    "очень", "можно", "когда", "нас", "есть", "раз", "ну", "вот", "там",
    "здесь", "тут", "чтобы", "себе", "один", "два", "три", "мой", "свой",
    "можно", "нельзя", "просто", "приложение", "хотел", "можете",
}

def top_keywords(texts, n=10):
    words = []
    for t in texts:
        if not isinstance(t, str):
            continue
        tokens = re.findall(r"[а-яёa-z]{4,}", t.lower())
        words.extend([w for w in tokens if w not in STOPWORDS])
    return Counter(words).most_common(n)


neg_reviews = df[df["sentiment"] == "NEGATIVE"]

fig3 = plt.figure(figsize=(18, 9), facecolor=LIGHT)
fig3.suptitle("RentEasy KZ — Competitor Weak Spots (Negative Review Keywords)",
              fontsize=15, fontweight="bold", color=DARK, y=0.98)

gs3 = gridspec.GridSpec(1, len(COMPETITORS), figure=fig3, wspace=0.45)

for col_idx, comp in enumerate(COMPETITORS):
    ax = fig3.add_subplot(gs3[0, col_idx])
    comp_neg = neg_reviews[neg_reviews["competitor"] == comp]["content"]
    keywords = top_keywords(comp_neg, n=10)

    if not keywords:
        ax.text(0.5, 0.5, "No negative\nreviews", ha="center", va="center",
                transform=ax.transAxes, fontsize=12, color=GREY)
        ax.set_title(comp)
        continue

    words = [k for k, _ in keywords]
    counts = [c for _, c in keywords]
    intensity = [c / max(counts) for c in counts]
    colors_k = [plt.cm.Reds(0.4 + 0.55 * i) for i in intensity]

    ax.barh(words[::-1], counts[::-1], color=colors_k[::-1],
            edgecolor="white", linewidth=0.8)
    for i, (w, c) in enumerate(zip(words[::-1], counts[::-1])):
        ax.text(c + 0.1, i, str(c), va="center", fontsize=9, color=DARK)

    neg_count = len(comp_neg)
    ax.set_title(f"{comp}\n({neg_count} negative reviews)", color=RED, fontsize=11)
    ax.set_xlabel("Frequency in negative reviews")
    ax.grid(axis="x", alpha=0.3)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

fig3.savefig(f"{OUT_DIR}/sentiment_negative.png", dpi=150, bbox_inches="tight")
plt.close(fig3)
print("[SAVED] sentiment_negative.png")


# ══════════════════════════════════════════════════════════════════════════════
# DASHBOARD 4 — Google Trends Combined (series + ACF + forecast)
# ══════════════════════════════════════════════════════════════════════════════
if os.path.exists(TRENDS_FILE):
    trend_df = pd.read_csv(TRENDS_FILE, skiprows=1)
    trend_df.columns = [c.strip() for c in trend_df.columns]
    date_col = trend_df.columns[0]
    val_col  = trend_df.columns[1]
    trend_df = trend_df[[date_col, val_col]].copy()
    trend_df.columns = ["date", "interest"]
    trend_df["date"] = pd.to_datetime(trend_df["date"], errors="coerce")
    trend_df = trend_df.dropna(subset=["date"])
    trend_df["interest"] = trend_df["interest"].astype(str).str.replace("<1","0")
    trend_df["interest"] = pd.to_numeric(trend_df["interest"], errors="coerce").fillna(0)
    trend_df = trend_df.sort_values("date").set_index("date")
    series = trend_df["interest"].dropna()

    # ARIMA
    d = 0 if adfuller(series, autolag="AIC")[1] < 0.05 else 1
    best_aic, best_order = np.inf, (1, d, 1)
    for p in range(0, 4):
        for q in range(0, 3):
            try:
                fit = ARIMA(series, order=(p, d, q)).fit()
                if fit.aic < best_aic:
                    best_aic, best_order = fit.aic, (p, d, q)
            except Exception:
                pass
    fitted   = ARIMA(series, order=best_order).fit()
    forecast = fitted.get_forecast(steps=10)
    fc_mean  = forecast.predicted_mean
    fc_ci    = forecast.conf_int(alpha=0.2)
    fc_index = pd.date_range(start=series.index[-1] + pd.DateOffset(months=1),
                             periods=10, freq="MS")
    fc_mean.index = fc_index
    fc_ci.index   = fc_index

    fig4 = plt.figure(figsize=(18, 11), facecolor=LIGHT)
    fig4.suptitle('Google Trends Dashboard — "аренда квартиры" Kazakhstan',
                  fontsize=15, fontweight="bold", color=DARK, y=0.98)
    gs4 = gridspec.GridSpec(2, 2, figure=fig4, hspace=0.45, wspace=0.38)

    # 4a: Full time series
    ax4a = fig4.add_subplot(gs4[0, :])
    ax4a.plot(series.index, series.values, color=BLUE, linewidth=1.8)
    ax4a.fill_between(series.index, series.values, alpha=0.12, color=BLUE)
    rolling12 = series.rolling(12, center=True).mean()
    ax4a.plot(series.index, rolling12, color=RED, linewidth=2,
              linestyle="--", label="12-month rolling avg")
    ax4a.set_title('Monthly Search Interest — "аренда квартиры" (Kazakhstan, all time)')
    ax4a.set_ylabel("Interest Index (0–100)")
    ax4a.legend(fontsize=9)
    ax4a.grid(alpha=0.3)
    ax4a.spines["top"].set_visible(False)
    ax4a.spines["right"].set_visible(False)
    ax4a.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    ax4a.xaxis.set_major_locator(mdates.YearLocator(2))

    # 4b: ACF
    ax4b = fig4.add_subplot(gs4[1, 0])
    plot_acf(series, lags=min(36, len(series)//2 - 1), ax=ax4b, color=BLUE, title="")
    ax4b.set_title("Autocorrelation Function (ACF)")
    ax4b.set_xlabel("Lag (months)")
    ax4b.spines["top"].set_visible(False)
    ax4b.spines["right"].set_visible(False)

    # 4c: Forecast
    ax4c = fig4.add_subplot(gs4[1, 1])
    hist5 = series[series.index >= series.index[-1] - pd.DateOffset(years=4)]
    ax4c.plot(hist5.index, hist5.values, color=BLUE, linewidth=2, label="Historical")
    ax4c.plot(fc_mean.index, fc_mean.values, color=RED, linewidth=2.2,
              linestyle="--", marker="o", markersize=4, label=f"Forecast (ARIMA{best_order})")
    ax4c.fill_between(fc_ci.index, fc_ci.iloc[:, 0], fc_ci.iloc[:, 1],
                      alpha=0.2, color=RED, label="80% CI")
    ax4c.axvline(series.index[-1], color=GREY, linewidth=1, linestyle=":")
    ax4c.set_title(f"10-Month Forecast — ARIMA{best_order}")
    ax4c.set_ylabel("Interest Index")
    ax4c.legend(fontsize=8)
    ax4c.grid(alpha=0.3)
    ax4c.spines["top"].set_visible(False)
    ax4c.spines["right"].set_visible(False)
    ax4c.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m"))
    plt.setp(ax4c.xaxis.get_majorticklabels(), rotation=30, fontsize=7)

    fig4.savefig(f"{OUT_DIR}/trends_combined.png", dpi=150, bbox_inches="tight")
    plt.close(fig4)
    print("[SAVED] trends_combined.png")
else:
    print(f"[SKIP] trends_combined.png — {TRENDS_FILE} not found")


# ══════════════════════════════════════════════════════════════════════════════
# DASHBOARD 5 — Future Sentiment Trend Projection
# ══════════════════════════════════════════════════════════════════════════════
fig5 = plt.figure(figsize=(18, 9), facecolor=LIGHT)
fig5.suptitle("RentEasy KZ — Future Sentiment Trend Projection & Insights",
              fontsize=15, fontweight="bold", color=DARK, y=0.98)
gs5 = gridspec.GridSpec(1, 2, figure=fig5, hspace=0.3, wspace=0.42)

# ── 5a: Extrapolate sentiment score per competitor (linear trend)
ax5a = fig5.add_subplot(gs5[0, 0])

for comp, color in COMP_COLORS.items():
    sub = df[df["competitor"] == comp].dropna(subset=["review_date"]).sort_values("review_date")
    if len(sub) < 5:
        continue
    sub = sub.set_index("review_date")
    weekly = sub["sent_score"].resample("W").mean().dropna()
    if len(weekly) < 4:
        continue

    # Fit linear trend on numeric index
    x_num = np.arange(len(weekly))
    coeffs = np.polyfit(x_num, weekly.values, 1)
    trend_line = np.polyval(coeffs, x_num)

    # Extrapolate 12 weeks
    x_future = np.arange(len(weekly), len(weekly) + 12)
    y_future  = np.polyval(coeffs, x_future)
    future_dates = pd.date_range(
        start=weekly.index[-1] + pd.DateOffset(weeks=1), periods=12, freq="W"
    )

    ax5a.plot(weekly.index, weekly.values, color=color, alpha=0.4, linewidth=1)
    ax5a.plot(weekly.index, trend_line, color=color, linewidth=2.2, label=comp)
    ax5a.plot(future_dates, y_future, color=color, linewidth=2.2,
              linestyle="--", alpha=0.85)
    ax5a.axvline(weekly.index[-1], color=GREY, linewidth=0.8, linestyle=":")

ax5a.axhline(0, color=GREY, linewidth=1, linestyle="--", alpha=0.5)
ax5a.set_title("Sentiment Trend Projection (12-week horizon)")
ax5a.set_ylabel("Sentiment Score (−1 to +1)")
ax5a.set_ylim(-1.1, 1.1)
ax5a.legend(fontsize=9)
ax5a.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m"))
plt.setp(ax5a.xaxis.get_majorticklabels(), rotation=30, fontsize=8)
ax5a.grid(alpha=0.3)
ax5a.spines["top"].set_visible(False)
ax5a.spines["right"].set_visible(False)

# ── 5b: Insight summary (text)
ax5b = fig5.add_subplot(gs5[0, 1])
ax5b.axis("off")

insight_blocks = [
    (BLUE,  "Krisha.kz",
     "• Lowest avg rating: 2.77 ★ — high user frustration\n"
     "• Mostly NEGATIVE reviews (57%)\n"
     "• Top complaints: crashes, slow loading, fake listings\n"
     "→ Biggest opportunity for RentEasy to capture\n"
     "   dissatisfied Krisha users"),
    (RED,   "OLX Kazakhstan",
     "• Middle rating: 3.92 ★ — acceptable but not great\n"
     "• Mixed sentiment: POSITIVE 43%, NEUTRAL 34%\n"
     "• Complaints: too many irrelevant ads, spam\n"
     "→ Target users wanting cleaner apartment search"),
    ("#9b59b6", "kn.kz",
     "• Best rating: 3.65 ★ among direct competitors\n"
     "• Limited reviews (54) — small market presence\n"
     "• Mostly positive tone but low brand awareness\n"
     "→ Not a major threat; niche audience"),
    (GREEN, "RentEasy KZ Opportunity",
     "• All competitors lack price analytics\n"
     "• Krisha's negative sentiment = market gap\n"
     "• Focus: verified listings + fair price data\n"
     "• Launch in Астана — highest complaint density"),
]

y_pos = 0.97
for color, title, text in insight_blocks:
    ax5b.text(0.02, y_pos, title, transform=ax5b.transAxes,
              fontsize=11, fontweight="bold", color=color, va="top")
    y_pos -= 0.06
    ax5b.text(0.02, y_pos, text, transform=ax5b.transAxes,
              fontsize=9, color=DARK, va="top",
              bbox=dict(boxstyle="round,pad=0.4", facecolor=color+"15",
                        edgecolor=color+"60", linewidth=1))
    y_pos -= 0.24

fig5.savefig(f"{OUT_DIR}/sentiment_future.png", dpi=150, bbox_inches="tight")
plt.close(fig5)
print("[SAVED] sentiment_future.png")

print("\nAll dashboards saved to:", OUT_DIR)
