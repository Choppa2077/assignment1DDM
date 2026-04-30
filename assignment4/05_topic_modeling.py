#!/usr/bin/env python3
"""
Assignment 4 - Task 2: Topic Modeling (LDA) + Trend Detection
"""
import warnings
warnings.filterwarnings("ignore")
import re
from pathlib import Path
from collections import Counter

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec

DATA_DIR = Path("data")
OUT = Path("output/topics")
OUT.mkdir(parents=True, exist_ok=True)

CUSTOM_STOP = {
    "study", "paper", "research", "analysis", "using", "based", "proposed",
    "method", "results", "data", "model", "models", "approach", "system",
    "systems", "learning", "use", "used", "also", "new", "show", "work",
    "different", "two", "one", "three", "first", "second", "third",
    "well", "however", "thus", "therefore", "due", "although", "within",
    "across", "provide", "present", "identify", "found", "find",
}

TOPIC_LABELS = {
    0: "Price Prediction & Valuation",
    1: "Sentiment & Review Analysis",
    2: "Demand Forecasting & Time Series",
    3: "Platform Design & UX",
    4: "Recommendation Systems",
    5: "Property Features & Classification",
    6: "Market Dynamics & Policy",
}


def tokenize(text: str, stop_words: set) -> list:
    text = text.lower()
    text = re.sub(r"[^a-z\s]", " ", text)
    return [t for t in text.split() if len(t) > 3 and t not in stop_words]


def run_lda(tokenized, n_topics=7):
    try:
        from gensim import corpora
        from gensim.models import LdaModel

        dictionary = corpora.Dictionary(tokenized)
        dictionary.filter_extremes(no_below=3, no_above=0.8)
        corpus = [dictionary.doc2bow(t) for t in tokenized]

        lda = LdaModel(corpus=corpus, id2word=dictionary, num_topics=n_topics,
                       random_state=42, passes=10, alpha="auto", eta="auto",
                       iterations=200)

        topics = {i: [w for w, _ in lda.show_topic(i, topn=20)]
                  for i in range(n_topics)}
        doc_topics = []
        for bow in corpus:
            probs = lda.get_document_topics(bow, minimum_probability=0.0)
            doc_topics.append(max(probs, key=lambda x: x[1])[0])

        return topics, doc_topics, "LDA (gensim)"

    except ImportError:
        print("  gensim not found — falling back to NMF (sklearn)")
        from sklearn.feature_extraction.text import TfidfVectorizer
        from sklearn.decomposition import NMF

        joined = [" ".join(t) for t in tokenized]
        vec = TfidfVectorizer(max_features=3000, min_df=3, max_df=0.8)
        X = vec.fit_transform(joined)
        vocab = vec.get_feature_names_out()

        nmf = NMF(n_components=n_topics, random_state=42, max_iter=500)
        W = nmf.fit_transform(X)
        H = nmf.components_

        topics = {}
        for i in range(n_topics):
            top_idx = H[i].argsort()[-20:][::-1]
            topics[i] = [vocab[j] for j in top_idx]

        doc_topics = W.argmax(axis=1).tolist()
        return topics, doc_topics, "NMF (sklearn)"


def trend_keywords(df) -> pd.DataFrame:
    years = sorted(df["year"].unique())
    rows = []
    all_kws = set()

    for _, row in df.iterrows():
        for col in ("keywords", "concepts"):
            val = row.get(col, "")
            if pd.notna(val) and val:
                for kw in str(val).split(";"):
                    all_kws.add(kw.strip().lower())

    # Pick top 20 most-frequent keywords overall
    kw_counter = Counter()
    for _, row in df.iterrows():
        for col in ("keywords", "concepts"):
            val = row.get(col, "")
            if pd.notna(val) and val:
                for kw in str(val).split(";"):
                    kw_counter[kw.strip().lower()] += 1

    top_kws = [k for k, _ in kw_counter.most_common(20)
               if len(k) > 3 and k not in CUSTOM_STOP][:15]

    for yr in years:
        yr_df = df[df["year"] == yr]
        for kw in top_kws:
            count = yr_df.apply(
                lambda r: kw in str(r.get("keywords", "")).lower()
                          or kw in str(r.get("concepts", "")).lower(),
                axis=1
            ).sum()
            rows.append({"year": yr, "keyword": kw, "count": int(count)})

    return pd.DataFrame(rows)


def visualize(topics, doc_topics, df, trend_df, method):
    n_topics = len(topics)
    fig = plt.figure(figsize=(22, 20))
    gs = gridspec.GridSpec(3, 3, figure=fig, hspace=0.55, wspace=0.4)

    # — Word clouds per topic ---
    try:
        from wordcloud import WordCloud
        use_wc = True
    except ImportError:
        use_wc = False

    for tid in range(min(n_topics, 6)):
        row, col = tid // 3, tid % 3
        ax = fig.add_subplot(gs[row, col])
        words = topics[tid]

        if use_wc:
            freq_dict = {w: (20 - i) for i, w in enumerate(words)}
            wc = WordCloud(width=380, height=260, background_color="white",
                           colormap="Blues", max_words=25).generate_from_frequencies(freq_dict)
            ax.imshow(wc, interpolation="bilinear")
        else:
            ax.barh(range(len(words[:10])), range(len(words[:10]), 0, -1),
                    color=plt.cm.Blues(np.linspace(0.4, 0.9, 10)))
            ax.set_yticks(range(len(words[:10])))
            ax.set_yticklabels(words[:10], fontsize=8)

        label = TOPIC_LABELS.get(tid, f"Topic {tid + 1}")
        ax.set_title(f"Topic {tid + 1}: {label}", fontsize=9, fontweight="bold", pad=8)
        ax.axis("off")

    # — Topic share over time ---
    ax_time = fig.add_subplot(gs[2, :])
    df_t = df.copy()
    df_t["topic"] = doc_topics
    tby = df_t.groupby(["year", "topic"]).size().reset_index(name="n")
    tby = tby[tby["year"] >= 2015]

    colors = plt.cm.Set2(np.linspace(0, 1, n_topics))
    for tid in range(n_topics):
        sub = tby[tby["topic"] == tid]
        if not sub.empty:
            ax_time.plot(sub["year"], sub["n"], marker="o",
                         label=TOPIC_LABELS.get(tid, f"Topic {tid + 1}"),
                         color=colors[tid], linewidth=2)

    ax_time.set_xlabel("Year", fontsize=11)
    ax_time.set_ylabel("Papers", fontsize=11)
    ax_time.set_title("Research Topic Trends 2015–2024", fontsize=12, fontweight="bold")
    ax_time.legend(fontsize=7, ncol=2, loc="upper left")
    ax_time.grid(True, alpha=0.3)

    plt.suptitle(f"RentEasy KZ — Topic Modeling ({method})",
                 fontsize=15, fontweight="bold", y=1.01)
    plt.savefig(OUT / "topic_modeling.png", dpi=150,
                bbox_inches="tight", facecolor="white")
    plt.close()
    print("  Saved topic_modeling.png")

    # — Keyword trend plot ---
    if not trend_df.empty and len(trend_df["keyword"].unique()) > 0:
        top5 = (trend_df.groupby("keyword")["count"].sum()
                .sort_values(ascending=False).head(8).index.tolist())
        fig2, ax2 = plt.subplots(figsize=(14, 6))
        for kw in top5:
            sub = trend_df[trend_df["keyword"] == kw].sort_values("year")
            ax2.plot(sub["year"], sub["count"], marker="o", label=kw, linewidth=2)
        ax2.set_xlabel("Year")
        ax2.set_ylabel("Paper Count")
        ax2.set_title("Keyword Trend Detection — Top Keywords Over Time",
                      fontsize=13, fontweight="bold")
        ax2.legend(fontsize=9, ncol=2)
        ax2.grid(True, alpha=0.3)
        plt.tight_layout()
        plt.savefig(OUT / "keyword_trends.png", dpi=150, bbox_inches="tight")
        plt.close()
        print("  Saved keyword_trends.png")


def main():
    df = pd.read_csv(DATA_DIR / "cleaned_publications.csv")
    print(f"Loaded {len(df)} publications")

    try:
        import nltk
        nltk.download("stopwords", quiet=True)
        from nltk.corpus import stopwords
        stop = set(stopwords.words("english")) | CUSTOM_STOP
    except Exception:
        stop = CUSTOM_STOP

    print("Preprocessing text ...")
    df["_text"] = (df["title"].fillna("") + " "
                   + df["keywords"].fillna("") + " "
                   + df["abstract"].fillna(""))
    tokenized = [tokenize(t, stop) for t in df["_text"]]

    valid = [(toks, i) for i, toks in enumerate(tokenized) if len(toks) >= 5]
    tokenized_f = [t for t, _ in valid]
    idx_f = [i for _, i in valid]
    df_f = df.iloc[idx_f].reset_index(drop=True)
    print(f"  Valid documents: {len(tokenized_f)}")

    print("Training topic model ...")
    topics, doc_topics, method = run_lda(tokenized_f, n_topics=7)
    print(f"  Method: {method}")

    print("\nDiscovered topics:")
    for tid, words in topics.items():
        print(f"  Topic {tid + 1} [{TOPIC_LABELS.get(tid, '')}]: {', '.join(words[:8])}")

    topics_rows = [{"topic_id": tid + 1,
                    "label": TOPIC_LABELS.get(tid, f"Topic {tid+1}"),
                    "top_words": ", ".join(words[:15])}
                   for tid, words in topics.items()]
    pd.DataFrame(topics_rows).to_csv(OUT / "lda_topics.csv", index=False)

    print("\nDetecting keyword trends ...")
    trend_df = trend_keywords(df)
    trend_df.to_csv(OUT / "keyword_trends.csv", index=False)

    print("Creating visualizations ...")
    visualize(topics, doc_topics, df_f, trend_df, method)
    print(f"\nTopic modeling complete → {OUT}/")


if __name__ == "__main__":
    main()
