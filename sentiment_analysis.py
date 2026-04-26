"""
Assignment 3 — Task 1: Sentiment Analysis of Competitor Reviews
RentEasy KZ

Steps:
1. Collect Google Play reviews for 3 competitors via google-play-scraper
2. Classify sentiment (POSITIVE / NEUTRAL / NEGATIVE) using multilingual transformer model
3. Evaluate quality against star ratings
4. Analyse dynamics over time
5. Save reviews_dataset.csv and sentiment_results.csv
"""

import pandas as pd
import numpy as np
import os
import time
from datetime import datetime
from google_play_scraper import reviews, Sort
from transformers import pipeline
from sklearn.metrics import accuracy_score, classification_report, confusion_matrix

OUTPUT_DIR = "renteasy_data"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ── Competitors ────────────────────────────────────────────────────────────
COMPETITORS = [
    {"name": "Krisha.kz",      "app_id": "kz.krisha"},
    {"name": "OLX Kazakhstan", "app_id": "kz.slando"},
    {"name": "kn.kz",          "app_id": "io.cordova.knkz"},
]

REVIEWS_PER_APP = 200


# ═══════════════════════════════════════════════════════════════════════════
# STEP 1 — Collect reviews
# ═══════════════════════════════════════════════════════════════════════════
print("=" * 60)
print("STEP 1: Collecting Google Play reviews")
print("=" * 60)

all_reviews = []

for comp in COMPETITORS:
    name = comp["name"]
    app_id = comp["app_id"]
    print(f"\n→ {name} ({app_id})")

    fetched = []
    try:
        result, _ = reviews(
            app_id,
            lang="ru",
            country="kz",
            sort=Sort.NEWEST,
            count=REVIEWS_PER_APP,
        )
        for r in result:
            fetched.append({
                "competitor":    name,
                "app_id":        app_id,
                "review_id":     r.get("reviewId", ""),
                "author":        r.get("userName", ""),
                "rating":        r.get("score", 0),
                "content":       r.get("content", ""),
                "thumbs_up":     r.get("thumbsUpCount", 0),
                "app_version":   r.get("appVersion", ""),
                "review_date":   r["at"].strftime("%Y-%m-%d") if r.get("at") else "",
            })
        print(f"  Collected: {len(fetched)} reviews")
    except Exception as e:
        print(f"  ERROR: {e}")

    all_reviews.extend(fetched)
    time.sleep(2)

df_reviews = pd.DataFrame(all_reviews)

# Drop rows with empty content
df_reviews = df_reviews[df_reviews["content"].str.strip() != ""].reset_index(drop=True)

df_reviews.to_csv(f"{OUTPUT_DIR}/reviews_dataset.csv", index=False, encoding="utf-8-sig")
print(f"\n[SAVED] reviews_dataset.csv — {len(df_reviews)} reviews")
print(df_reviews.groupby("competitor")["rating"].agg(["count", "mean"]).round(2))


# ═══════════════════════════════════════════════════════════════════════════
# STEP 2 — Sentiment classification
# ═══════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 60)
print("STEP 2: Sentiment classification (multilingual transformer)")
print("=" * 60)

print("Loading model: nlptown/bert-base-multilingual-uncased-sentiment ...")
try:
    classifier = pipeline(
        "sentiment-analysis",
        model="nlptown/bert-base-multilingual-uncased-sentiment",
        max_length=128,
        truncation=True,
    )
    print("Model loaded successfully.")
    USE_MODEL = True
except Exception as e:
    print(f"Model load failed: {e}\nFalling back to keyword-based approach.")
    USE_MODEL = False


def keyword_sentiment(text):
    """Fast Russian keyword-based fallback."""
    text = text.lower()
    pos = ["отлично", "хорошо", "удобно", "нравится", "замечательно", "быстро",
           "удобный", "рекомендую", "спасибо", "класс", "супер", "топ", "круто",
           "лучший", "отличный", "прекрасный", "5 звезд", "молодцы"]
    neg = ["плохо", "ужасно", "не работает", "баг", "глючит", "медленно",
           "мошенники", "обман", "фейк", "не открывается", "вылетает",
           "ошибка", "не советую", "разочарование", "отстой", "хлам", "удалил"]
    p = sum(w in text for w in pos)
    n = sum(w in text for w in neg)
    if p > n:
        return "POSITIVE", 0.75
    elif n > p:
        return "NEGATIVE", 0.75
    return "NEUTRAL", 0.60


# nlptown model returns "1 star" to "5 stars"
def stars_to_sentiment(label):
    n = int(label.split()[0])
    if n <= 2:
        return "NEGATIVE"
    elif n == 3:
        return "NEUTRAL"
    return "POSITIVE"


def classify_batch(texts, batch_size=32):
    labels = []
    scores = []
    if USE_MODEL:
        for i in range(0, len(texts), batch_size):
            batch = texts[i:i + batch_size]
            results = classifier(batch)
            for r in results:
                lbl = stars_to_sentiment(r["label"])
                labels.append(lbl)
                scores.append(round(r["score"], 4))
            print(f"  Classified {min(i+batch_size, len(texts))}/{len(texts)}")
    else:
        for t in texts:
            lbl, sc = keyword_sentiment(t)
            labels.append(lbl)
            scores.append(sc)
    return labels, scores


texts = df_reviews["content"].tolist()
print(f"\nClassifying {len(texts)} reviews...")
sentiment_labels, sentiment_scores = classify_batch(texts)

df_reviews["sentiment"]       = sentiment_labels
df_reviews["sentiment_score"] = sentiment_scores


# ═══════════════════════════════════════════════════════════════════════════
# STEP 3 — Quality evaluation against star ratings
# ═══════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 60)
print("STEP 3: Quality evaluation (sentiment vs. star rating)")
print("=" * 60)


def rating_to_sentiment(r):
    if r <= 2:
        return "NEGATIVE"
    elif r == 3:
        return "NEUTRAL"
    return "POSITIVE"


df_reviews["rating_sentiment"] = df_reviews["rating"].apply(rating_to_sentiment)

y_true = df_reviews["rating_sentiment"]
y_pred = df_reviews["sentiment"]

acc = accuracy_score(y_true, y_pred)
print(f"\nOverall Accuracy: {acc:.3f} ({acc*100:.1f}%)")
print("\nClassification Report:")
print(classification_report(y_true, y_pred, labels=["POSITIVE", "NEUTRAL", "NEGATIVE"]))

print("Confusion Matrix (rows=true, cols=predicted):")
cm = confusion_matrix(y_true, y_pred, labels=["POSITIVE", "NEUTRAL", "NEGATIVE"])
cm_df = pd.DataFrame(cm,
    index=["True POS", "True NEU", "True NEG"],
    columns=["Pred POS", "Pred NEU", "Pred NEG"]
)
print(cm_df)


# ═══════════════════════════════════════════════════════════════════════════
# STEP 4 — Save results
# ═══════════════════════════════════════════════════════════════════════════
df_reviews.to_csv(f"{OUTPUT_DIR}/sentiment_results.csv", index=False, encoding="utf-8-sig")
print(f"\n[SAVED] sentiment_results.csv — {len(df_reviews)} rows")

print("\nSentiment distribution per competitor:")
print(pd.crosstab(df_reviews["competitor"], df_reviews["sentiment"]))

print("\nAverage rating per competitor:")
print(df_reviews.groupby("competitor")["rating"].mean().round(2))
