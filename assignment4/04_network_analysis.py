#!/usr/bin/env python3
"""
Assignment 4 - Task 2: Keyword Co-occurrence Network
"""
import warnings
warnings.filterwarnings("ignore")

from itertools import combinations
from pathlib import Path
from collections import Counter, defaultdict

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.cm as cm
import networkx as nx

DATA_DIR = Path("data")
OUT = Path("output/networks")
OUT.mkdir(parents=True, exist_ok=True)

STOP = {
    "the", "a", "an", "of", "for", "in", "on", "with", "and", "or",
    "to", "is", "are", "was", "were", "be", "been", "from", "by", "at",
    "based", "using", "proposed", "paper", "study", "approach", "method",
    "methods", "model", "models", "system", "systems", "data", "results",
    "analysis", "research", "work", "used", "new", "show", "this", "that",
    "computer", "science", "artificial", "intelligence", "business",
    "economics", "finance", "algorithm", "also", "high", "large", "different",
    "two", "three", "well", "case", "through", "each", "such", "which",
    "their", "these", "other", "than", "more", "about", "when", "have",
}

# Curated domain keywords to supplement sparse OpenAlex keyword field
DOMAIN_KWS = [
    "machine learning", "deep learning", "neural network", "random forest",
    "gradient boosting", "xgboost", "support vector", "regression",
    "price prediction", "property valuation", "hedonic pricing",
    "real estate", "rental market", "housing market",
    "sentiment analysis", "nlp", "text mining", "bert", "transformer",
    "review analysis", "user satisfaction", "opinion mining",
    "time series", "arima", "forecasting", "demand forecasting", "prophet",
    "lstm", "recurrent neural", "convolutional neural",
    "recommendation system", "collaborative filtering", "matrix factorization",
    "user experience", "mobile application", "platform design",
    "spatial analysis", "gis", "location", "neighbourhood",
    "feature selection", "feature engineering", "transfer learning",
    "explainable ai", "shap", "interpretability",
    "data collection", "web scraping", "crowdsourcing",
    "clustering", "classification", "anomaly detection",
    "image recognition", "computer vision",
]


def get_keywords(row) -> list:
    import re
    items = []

    # 1) structured keyword / concept fields
    for col in ("keywords", "concepts"):
        val = row.get(col, "")
        if pd.notna(val) and val:
            for kw in str(val).split(";"):
                kw = kw.strip().lower()
                if kw and len(kw) > 3 and kw not in STOP:
                    items.append(kw)

    # 2) domain keyword match in title + abstract
    combined = (str(row.get("title", "")) + " " +
                str(row.get("abstract", ""))).lower()
    combined = re.sub(r"[^a-z\s]", " ", combined)
    for dkw in DOMAIN_KWS:
        if dkw in combined:
            items.append(dkw)

    return list(dict.fromkeys(items[:20]))


def build_network(df, min_freq=4, min_cooc=3):
    kw_per_paper = [get_keywords(row) for _, row in df.iterrows()]
    freq = Counter(kw for kws in kw_per_paper for kw in kws)
    valid = {kw for kw, f in freq.items() if f >= min_freq}

    cooc = Counter()
    for kws in kw_per_paper:
        filtered = [k for k in kws if k in valid]
        for pair in combinations(sorted(filtered), 2):
            cooc[pair] += 1

    G = nx.Graph()
    for kw in valid:
        G.add_node(kw, freq=freq[kw])
    for (a, b), w in cooc.items():
        if w >= min_cooc:
            G.add_edge(a, b, weight=w)

    G.remove_nodes_from(list(nx.isolates(G)))
    print(f"  Network: {G.number_of_nodes()} nodes, {G.number_of_edges()} edges")
    return G, freq


def centrality_table(G) -> pd.DataFrame:
    deg = nx.degree_centrality(G)
    btw = nx.betweenness_centrality(G, weight="weight")
    try:
        eig = nx.eigenvector_centrality(G, weight="weight", max_iter=500)
    except Exception:
        eig = {n: 0.0 for n in G.nodes()}
    pr = nx.pagerank(G, weight="weight")

    rows = []
    for n in G.nodes():
        rows.append({
            "keyword": n,
            "frequency": G.nodes[n].get("freq", 0),
            "degree": round(deg[n], 4),
            "betweenness": round(btw[n], 4),
            "eigenvector": round(eig.get(n, 0), 4),
            "pagerank": round(pr[n], 4),
        })
    return pd.DataFrame(rows).sort_values("pagerank", ascending=False)


def detect_communities(G):
    from networkx.algorithms.community import greedy_modularity_communities
    comms = greedy_modularity_communities(G)
    node2comm = {}
    for i, c in enumerate(comms):
        for n in c:
            node2comm[n] = i
    return node2comm, len(comms)


def visualize(G, cent_df, node2comm):
    # Keep top 80 nodes by pagerank for readability
    top_nodes = cent_df.head(80)["keyword"].tolist()
    Gv = G.subgraph([n for n in top_nodes if n in G]).copy()

    pos = nx.spring_layout(Gv, k=1.8 / np.sqrt(max(Gv.number_of_nodes(), 1)),
                            seed=42, iterations=60)

    num_comm = len(set(node2comm.values()))
    cmap = cm.get_cmap("tab20", max(num_comm, 1))

    node_colors = [cmap(node2comm.get(n, 0) % 20) for n in Gv.nodes()]
    node_sizes = [max(80, Gv.nodes[n].get("freq", 1) * 10) for n in Gv.nodes()]
    edge_w = [Gv[u][v].get("weight", 1) for u, v in Gv.edges()]
    max_w = max(edge_w) if edge_w else 1
    edge_widths = [0.3 + 2.5 * (w / max_w) for w in edge_w]

    fig, axes = plt.subplots(1, 2, figsize=(22, 11))

    for ax_i, ax in enumerate(axes):
        nx.draw_networkx_edges(Gv, pos, ax=ax, width=edge_widths,
                               alpha=0.35, edge_color="#888")
        nx.draw_networkx_nodes(Gv, pos, ax=ax, node_color=node_colors,
                               node_size=node_sizes, alpha=0.9)
        ax.axis("off")

    # Left: top labels by pagerank
    pr_map = dict(zip(cent_df["keyword"], cent_df["pagerank"]))
    top_labels = sorted([n for n in Gv.nodes()],
                        key=lambda n: pr_map.get(n, 0), reverse=True)[:30]
    nx.draw_networkx_labels(Gv, pos, labels={n: n for n in top_labels},
                            ax=axes[0], font_size=7)
    axes[0].set_title(
        f"Keyword Co-occurrence Network\n({Gv.number_of_nodes()} nodes, "
        f"{Gv.number_of_edges()} edges)", fontsize=13, fontweight="bold")

    # Right: community cluster labels
    comm_labels = sorted(set(node2comm.values()),
                         key=lambda c: sum(1 for n in Gv.nodes() if node2comm.get(n) == c),
                         reverse=True)[:10]
    for ci, comm_id in enumerate(comm_labels):
        nodes_c = [n for n in Gv.nodes() if node2comm.get(n) == comm_id and n in pos]
        if not nodes_c:
            continue
        centroid = np.mean([pos[n] for n in nodes_c], axis=0)
        best = max(nodes_c, key=lambda n: pr_map.get(n, 0))
        axes[1].annotate(f"C{ci+1}: {best}", centroid, fontsize=8,
                         color=cmap(comm_id % 20),
                         bbox=dict(boxstyle="round,pad=0.2", fc="white", alpha=0.7))
    axes[1].set_title("Community Structure", fontsize=13, fontweight="bold")

    plt.suptitle("RentEasy KZ — Keyword Co-occurrence Network Analysis",
                 fontsize=15, fontweight="bold", y=1.01)
    plt.tight_layout()
    plt.savefig(OUT / "network_keyword_cooccurrence.png",
                dpi=150, bbox_inches="tight", facecolor="white")
    plt.close()
    print("  Saved network_keyword_cooccurrence.png")


def main():
    df = pd.read_csv(DATA_DIR / "cleaned_publications.csv")
    print(f"Loaded {len(df)} publications")

    print("Building keyword co-occurrence network ...")
    G, freq = build_network(df, min_freq=4, min_cooc=3)

    print("Calculating centrality metrics ...")
    cent_df = centrality_table(G)
    cent_df.to_csv(OUT / "centrality_metrics.csv", index=False)

    print("Detecting communities ...")
    node2comm, n_comm = detect_communities(G)
    print(f"  {n_comm} communities found")

    cent_df["community"] = cent_df["keyword"].map(node2comm)
    cent_df.to_csv(OUT / "centrality_metrics.csv", index=False)

    print("\nTop-10 keywords by PageRank:")
    for _, r in cent_df.head(10).iterrows():
        print(f"  {r['keyword']:<35} PR={r['pagerank']:.4f}  freq={r['frequency']}")

    print("Visualizing ...")
    visualize(G, cent_df, node2comm)

    nx.write_graphml(G, OUT / "keyword_network.graphml")
    print(f"\nNetwork analysis complete → {OUT}/")


if __name__ == "__main__":
    main()
