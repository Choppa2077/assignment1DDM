"""
Assignment 3 — Task 2: Google Trends Time-Series Analysis
RentEasy KZ

Reads: renteasy_data/google_trends_data.csv  (downloaded from trends.google.com)
Steps:
  1. Load and parse the Google Trends CSV
  2. Visualize the time series
  3. Stationarity: ACF plot + Augmented Dickey-Fuller test
  4. Forecast: ARIMA (horizon = 10 months) with confidence interval
Saves 3 PNG files to renteasy_data/dashboards/
"""

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import os
import sys

from statsmodels.graphics.tsaplots import plot_acf
from statsmodels.tsa.stattools import adfuller
from statsmodels.tsa.arima.model import ARIMA
import warnings
warnings.filterwarnings("ignore")

INPUT_FILE = "renteasy_data/google_trends_data.csv"
OUT_DIR    = "renteasy_data/dashboards"
os.makedirs(OUT_DIR, exist_ok=True)

BLUE  = "#2A5CAA"
RED   = "#E85D2C"
LIGHT = "#F5F3EE"
DARK  = "#1A1A2E"
GREY  = "#7F9799"

plt.rcParams.update({
    "figure.facecolor": LIGHT,
    "axes.facecolor": "white",
    "axes.edgecolor": GREY,
    "text.color": DARK,
    "font.family": "DejaVu Sans",
})


# ═══════════════════════════════════════════════════════════════════════════
# 1. LOAD DATA
# ═══════════════════════════════════════════════════════════════════════════
if not os.path.exists(INPUT_FILE):
    print(f"ERROR: {INPUT_FILE} not found.")
    print("Please download from https://trends.google.com/trends")
    print("Search: 'аренда квартиры', region: Kazakhstan, all time")
    print(f"Save CSV as: {INPUT_FILE}")
    sys.exit(1)

# Google Trends CSV — 1 header row, dates as YYYY-MM-DD
df = pd.read_csv(INPUT_FILE)
df.columns = [c.strip() for c in df.columns]

# The date column name varies — find it
date_col = df.columns[0]
val_col  = df.columns[1]

df = df[[date_col, val_col]].copy()
df.columns = ["date", "interest"]

# Parse dates (Google exports as YYYY-MM or YYYY-MM-DD)
df["date"] = pd.to_datetime(df["date"], errors="coerce")
df = df.dropna(subset=["date"])

# Interest is 0-100, sometimes "<1" for very low — replace with 0
df["interest"] = df["interest"].astype(str).str.replace("<1", "0").str.strip()
df["interest"] = pd.to_numeric(df["interest"], errors="coerce").fillna(0)

df = df.sort_values("date").reset_index(drop=True)
df = df.set_index("date")

print(f"Loaded {len(df)} data points: {df.index[0].date()} — {df.index[-1].date()}")
print(f"Interest range: {df['interest'].min()} — {df['interest'].max()}")
print(f"Mean: {df['interest'].mean():.1f}")


# ═══════════════════════════════════════════════════════════════════════════
# 2. VISUALIZE TIME SERIES
# ═══════════════════════════════════════════════════════════════════════════
fig, ax = plt.subplots(figsize=(14, 5), facecolor=LIGHT)

ax.plot(df.index, df["interest"], color=BLUE, linewidth=1.8, label="Interest index")
ax.fill_between(df.index, df["interest"], alpha=0.15, color=BLUE)

# 12-month rolling average
rolling = df["interest"].rolling(window=12, center=True).mean()
ax.plot(df.index, rolling, color=RED, linewidth=2.2, linestyle="--", label="12-month avg")

# Annotate key events
events = {
    "2020-03": ("COVID-19\npandemic", -18),
    "2022-09": ("Astana\nrename", -18),
    "2021-01": ("KZ housing\nboom", 5),
}
for date_str, (label, yoffset) in events.items():
    try:
        event_date = pd.Timestamp(date_str)
        if event_date in df.index or df.index.asof(event_date) is not None:
            y_val = df["interest"].asof(event_date)
            ax.axvline(event_date, color=GREY, linewidth=1, linestyle=":")
            ax.annotate(label, xy=(event_date, y_val),
                        xytext=(event_date, y_val + yoffset),
                        fontsize=7.5, color=GREY, ha="center",
                        arrowprops=dict(arrowstyle="-", color=GREY, lw=0.8))
    except Exception:
        pass

ax.set_title('Google Trends: "аренда квартиры" in Kazakhstan (Interest over Time)',
             fontsize=13, fontweight="bold", color=DARK)
ax.set_ylabel("Interest Index (0–100)", color=DARK)
ax.set_xlabel("Year", color=DARK)
ax.legend(fontsize=9)
ax.grid(alpha=0.3)
ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)
ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
ax.xaxis.set_major_locator(mdates.YearLocator(2))

fig.tight_layout()
fig.savefig(f"{OUT_DIR}/trends_timeseries.png", dpi=150, bbox_inches="tight")
plt.close(fig)
print(f"[SAVED] trends_timeseries.png")


# ═══════════════════════════════════════════════════════════════════════════
# 3. STATIONARITY — ACF + ADF Test
# ═══════════════════════════════════════════════════════════════════════════
series = df["interest"].dropna()

# ADF test
adf_result = adfuller(series, autolag="AIC")
adf_stat   = adf_result[0]
adf_pvalue = adf_result[1]
is_stationary = adf_pvalue < 0.05

print(f"\nAugmented Dickey-Fuller Test:")
print(f"  ADF Statistic: {adf_stat:.4f}")
print(f"  p-value:       {adf_pvalue:.4f}")
print(f"  Stationary:    {'YES' if is_stationary else 'NO'} (p {'<' if is_stationary else '>='} 0.05)")

# ACF plot
fig, axes = plt.subplots(1, 2, figsize=(14, 5), facecolor=LIGHT)

# Left: ACF
plot_acf(series, lags=36, ax=axes[0], color=BLUE, title="")
axes[0].set_title("Autocorrelation Function (ACF) — 36 lags", fontsize=12,
                  fontweight="bold", color=DARK)
axes[0].set_xlabel("Lag (months)", color=DARK)
axes[0].set_ylabel("ACF", color=DARK)
axes[0].spines["top"].set_visible(False)
axes[0].spines["right"].set_visible(False)

# Right: ADF summary box
axes[1].axis("off")
axes[1].set_facecolor(LIGHT)

summary_text = (
    f"Stationarity Analysis\n"
    f"{'─'*35}\n\n"
    f"Method: Augmented Dickey-Fuller (ADF)\n\n"
    f"ADF Statistic:  {adf_stat:.4f}\n"
    f"p-value:        {adf_pvalue:.4f}\n"
    f"Significance:   α = 0.05\n\n"
    f"Result: The series is\n"
    f"{'STATIONARY ✓' if is_stationary else 'NON-STATIONARY ✗'}\n\n"
)
if is_stationary:
    summary_text += "The mean and variance remain\nconstant over time — suitable\nfor direct ARIMA forecasting."
else:
    summary_text += "Trend component detected.\nFirst-order differencing\napplied before ARIMA (d=1)."

color = "#2ecc71" if is_stationary else RED
axes[1].text(0.1, 0.85, summary_text, transform=axes[1].transAxes,
             fontsize=11, verticalalignment="top", color=DARK,
             bbox=dict(boxstyle="round,pad=0.6", facecolor=color+"22",
                       edgecolor=color, linewidth=2))

fig.suptitle('Google Trends Stationarity Analysis — "аренда квартиры" Kazakhstan',
             fontsize=13, fontweight="bold", color=DARK, y=1.01)
fig.tight_layout()
fig.savefig(f"{OUT_DIR}/trends_acf.png", dpi=150, bbox_inches="tight")
plt.close(fig)
print(f"[SAVED] trends_acf.png")


# ═══════════════════════════════════════════════════════════════════════════
# 4. ARIMA FORECAST — 10 months horizon
# ═══════════════════════════════════════════════════════════════════════════
HORIZON = 10

# Choose d based on stationarity
d = 0 if is_stationary else 1

# Fit ARIMA — try several (p,d,q) and pick best AIC
best_aic = np.inf
best_order = (1, d, 1)
for p in range(0, 4):
    for q in range(0, 4):
        try:
            model = ARIMA(series, order=(p, d, q))
            fit   = model.fit()
            if fit.aic < best_aic:
                best_aic   = fit.aic
                best_order = (p, d, q)
        except Exception:
            pass

print(f"\nBest ARIMA order: {best_order} (AIC={best_aic:.1f})")

model     = ARIMA(series, order=best_order)
fitted    = model.fit()
forecast  = fitted.get_forecast(steps=HORIZON)
fc_mean   = forecast.predicted_mean
fc_ci     = forecast.conf_int(alpha=0.2)  # 80% confidence interval

# Build forecast index (monthly)
last_date = series.index[-1]
fc_index  = pd.date_range(start=last_date + pd.DateOffset(months=1),
                          periods=HORIZON, freq="MS")
fc_mean.index = fc_index
fc_ci.index   = fc_index

# Plot
fig, ax = plt.subplots(figsize=(14, 5), facecolor=LIGHT)

# Historical (last 5 years for clarity)
hist_start = series.index[-1] - pd.DateOffset(years=5)
hist = series[series.index >= hist_start]

ax.plot(hist.index, hist.values, color=BLUE, linewidth=2, label="Historical data")
ax.plot(fc_mean.index, fc_mean.values, color=RED, linewidth=2.5,
        linestyle="--", marker="o", markersize=4, label=f"ARIMA{best_order} forecast")
ax.fill_between(fc_ci.index,
                fc_ci.iloc[:, 0], fc_ci.iloc[:, 1],
                alpha=0.2, color=RED, label="80% confidence interval")

ax.axvline(last_date, color=GREY, linewidth=1.2, linestyle=":", label="Forecast start")

ax.set_title(f'ARIMA{best_order} Forecast — "аренда квартиры" KZ — Horizon: {HORIZON} months',
             fontsize=13, fontweight="bold", color=DARK)
ax.set_ylabel("Interest Index (0–100)", color=DARK)
ax.set_xlabel("Date", color=DARK)
ax.legend(fontsize=9)
ax.grid(alpha=0.3)
ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)
ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m"))
ax.xaxis.set_major_locator(mdates.MonthLocator(interval=6))
plt.xticks(rotation=30)

# Annotate forecast values
for date, val in zip(fc_mean.index[::2], fc_mean.values[::2]):
    ax.annotate(f"{val:.0f}", xy=(date, val), xytext=(0, 8),
                textcoords="offset points", ha="center", fontsize=8, color=RED)

fig.tight_layout()
fig.savefig(f"{OUT_DIR}/trends_forecast.png", dpi=150, bbox_inches="tight")
plt.close(fig)
print(f"[SAVED] trends_forecast.png")

print(f"\nForecast for next {HORIZON} months:")
for date, val in zip(fc_mean.index, fc_mean.values):
    lo = fc_ci.iloc[list(fc_ci.index).index(date), 0]
    hi = fc_ci.iloc[list(fc_ci.index).index(date), 1]
    print(f"  {date.strftime('%Y-%m')}: {val:.1f}  [{lo:.1f} — {hi:.1f}]")
