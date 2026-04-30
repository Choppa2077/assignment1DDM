#!/usr/bin/env python3
"""
Assignment 4 - Task 2: Dashboards (5 visualizations)
"""
import warnings
warnings.filterwarnings("ignore")

from pathlib import Path
from collections import Counter

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots

DATA_DIR = Path("data")
OUT = Path("output/figures")
OUT.mkdir(parents=True, exist_ok=True)

BLUE = "#1E3A8A"
LBLUE = "#3B82F6"
AMBER = "#F59E0B"
GREEN = "#10B981"
RED = "#EF4444"


def save(fig, name: str):
    fig.write_html(OUT / f"{name}.html")
    try:
        fig.write_image(OUT / f"{name}.png", scale=2, width=1000, height=520)
        print(f"  Saved {name}.png + .html")
    except Exception:
        print(f"  Saved {name}.html  (kaleido not installed – PNG skipped)")


def fig1_dynamics(df):
    yc = df.groupby("year").size().reset_index(name="n")
    yc = yc[yc["year"] >= 2015]
    z = np.polyfit(yc["year"], yc["n"], 1)
    trend = np.poly1d(z)(yc["year"])

    fig = go.Figure()
    fig.add_bar(x=yc["year"], y=yc["n"], name="Publications",
                marker_color=BLUE, opacity=0.8,
                text=yc["n"], textposition="outside")
    fig.add_scatter(x=yc["year"], y=trend, mode="lines", name="Trend",
                    line=dict(color=AMBER, width=2, dash="dash"))
    fig.update_layout(
        title="<b>Publication Dynamics 2015–2024</b><br><sup>ML / NLP for Real Estate Rental Markets</sup>",
        xaxis_title="Year", yaxis_title="Publications",
        template="plotly_white", height=480,
        legend=dict(yanchor="top", y=0.99, xanchor="left", x=0.01))
    save(fig, "fig1_publication_dynamics")


def fig2_journals(df):
    jdf = (df[df["journal"] != "Unknown Journal"]
           .groupby("journal")
           .agg(count=("title", "count"), avg_cit=("cited_by_count", "mean"))
           .reset_index()
           .sort_values("count", ascending=False)
           .head(10))

    fig = go.Figure(go.Bar(
        x=jdf["count"], y=jdf["journal"], orientation="h",
        marker=dict(color=jdf["avg_cit"], colorscale="Blues",
                    showscale=True, colorbar=dict(title="Avg Citations")),
        text=jdf["count"], textposition="outside"))
    fig.update_layout(
        title="<b>Top-10 Journals by Publication Count</b><br><sup>Color = average citation impact</sup>",
        xaxis_title="Publications", yaxis=dict(autorange="reversed"),
        template="plotly_white", height=500)
    save(fig, "fig2_top_journals")


def fig3_authors(df):
    counts = Counter(df["first_author"].dropna().replace("", np.nan).dropna())
    top = pd.DataFrame(counts.most_common(10), columns=["author", "n"])

    fig = go.Figure(go.Bar(
        x=top["n"], y=top["author"], orientation="h",
        marker_color=LBLUE, text=top["n"], textposition="outside"))
    fig.update_layout(
        title="<b>Top-10 Most Productive First Authors</b>",
        xaxis_title="Publications", yaxis=dict(autorange="reversed"),
        template="plotly_white", height=460)
    save(fig, "fig3_top_authors")


def fig4_countries(df):
    records = []
    for countries in df["country"].dropna():
        for c in str(countries).split(";"):
            c = c.strip()
            if c and c != "Unknown":
                records.append(c)
    cc = Counter(records)
    top = pd.DataFrame(cc.most_common(10), columns=["country", "n"])

    ISO = {
        "United States": "USA", "China": "CHN", "United Kingdom": "GBR",
        "Germany": "DEU", "India": "IND", "Australia": "AUS", "France": "FRA",
        "Canada": "CAN", "South Korea": "KOR", "Italy": "ITA", "Japan": "JPN",
        "Netherlands": "NLD", "Spain": "ESP", "Brazil": "BRA", "Singapore": "SGP",
        "Sweden": "SWE", "Switzerland": "CHE", "Poland": "POL", "Russia": "RUS",
        "Kazakhstan": "KAZ", "Malaysia": "MYS", "South Africa": "ZAF",
        "Portugal": "PRT", "Turkey": "TUR", "Indonesia": "IDN", "Taiwan": "TWN",
    }

    fig = make_subplots(
        rows=1, cols=2,
        subplot_titles=["Top-10 Countries", "Global Distribution"],
        column_widths=[0.45, 0.55],
        specs=[[{"type": "bar"}, {"type": "choropleth"}]])

    fig.add_trace(go.Bar(x=top["country"], y=top["n"],
                         marker_color=BLUE, text=top["n"],
                         textposition="outside", showlegend=False), row=1, col=1)

    all_df = pd.DataFrame(cc.most_common(60), columns=["country", "n"])
    all_df["iso"] = all_df["country"].map(ISO)
    all_df = all_df.dropna(subset=["iso"])
    fig.add_trace(go.Choropleth(locations=all_df["iso"], z=all_df["n"],
                                colorscale="Blues", showscale=True,
                                showlegend=False), row=1, col=2)

    fig.update_layout(title="<b>Geographic Distribution of Publications</b>",
                      template="plotly_white", height=480)
    save(fig, "fig4_top_countries")


def fig5_citations(df):
    clip95 = df["cited_by_count"].quantile(0.95)
    cit = df["cited_by_count"].clip(upper=clip95)

    fig = make_subplots(rows=1, cols=2,
                        subplot_titles=["Citation Distribution", "Citations by Year (box)"])

    fig.add_trace(go.Histogram(x=cit, nbinsx=50, marker_color=LBLUE,
                               opacity=0.8, name="Count"), row=1, col=1)

    for yr in sorted(df[df["year"] >= 2015]["year"].unique()):
        vals = df[df["year"] == yr]["cited_by_count"].clip(
            upper=df[df["year"] == yr]["cited_by_count"].quantile(0.90))
        fig.add_trace(go.Box(y=vals, name=str(yr), showlegend=False,
                             marker_color=BLUE), row=1, col=2)

    fig.update_xaxes(title_text="Citations", row=1, col=1)
    fig.update_yaxes(title_text="Papers", row=1, col=1)
    fig.update_xaxes(title_text="Year", row=1, col=2)
    fig.update_yaxes(title_text="Citations", row=1, col=2)
    fig.update_layout(title="<b>Citation Analysis</b>",
                      template="plotly_white", height=480)
    save(fig, "fig5_citation_distribution")


def main():
    df = pd.read_csv(DATA_DIR / "cleaned_publications.csv")
    print(f"Loaded {len(df)} records — generating dashboards...")
    fig1_dynamics(df)
    fig2_journals(df)
    fig3_authors(df)
    fig4_countries(df)
    fig5_citations(df)
    print(f"\nAll figures saved to {OUT}/")


if __name__ == "__main__":
    main()
