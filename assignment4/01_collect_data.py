#!/usr/bin/env python3
"""
Assignment 4 - Task 1: Data Collection from OpenAlex API
RentEasy KZ — Scientific and Network Analytics
Topic: ML/NLP for Real Estate Rental Platforms
"""
import requests
import json
import time
from pathlib import Path

MAILTO = "kaketika009@gmail.com"
BASE_URL = "https://api.openalex.org/works"
DATA_DIR = Path("data")

SELECT_FIELDS = ",".join([
    "id", "doi", "title", "publication_year", "publication_date",
    "authorships", "primary_location",
    "cited_by_count", "keywords", "concepts",
    "abstract_inverted_index", "open_access"
])

SEARCHES = [
    {
        "query": "rental price prediction machine learning regression",
        "filters": "publication_year:2015-2024,type:article",
        "max": 250,
        "label": "Price Prediction ML"
    },
    {
        "query": "real estate sentiment analysis user reviews NLP text mining",
        "filters": "publication_year:2015-2024,type:article",
        "max": 220,
        "label": "Sentiment Analysis"
    },
    {
        "query": "housing demand forecasting time series ARIMA deep learning",
        "filters": "publication_year:2015-2024,type:article",
        "max": 220,
        "label": "Demand Forecasting"
    },
    {
        "query": "property recommendation system collaborative filtering real estate",
        "filters": "publication_year:2015-2024,type:article",
        "max": 180,
        "label": "Recommendation Systems"
    },
    {
        "query": "apartment rental platform mobile application user experience proptech",
        "filters": "publication_year:2017-2024,type:article",
        "max": 160,
        "label": "Platform UX / PropTech"
    },
    {
        "query": "real estate price prediction neural network convolutional",
        "filters": "publication_year:2018-2024,type:article",
        "max": 150,
        "label": "Deep Learning Pricing"
    },
]


def fetch_works(query: str, filters: str = "", max_results: int = 250) -> list:
    publications = []
    cursor = "*"

    while len(publications) < max_results:
        per_page = min(200, max_results - len(publications))
        params = {
            "search": query,
            "filter": filters,
            "per-page": per_page,
            "cursor": cursor,
            "select": SELECT_FIELDS,
            "mailto": MAILTO,
        }

        try:
            resp = requests.get(BASE_URL, params=params, timeout=30)
            resp.raise_for_status()
            data = resp.json()
        except Exception as e:
            print(f"    [WARN] fetch error: {e}")
            break

        results = data.get("results", [])
        if not results:
            break

        publications.extend(results)
        print(f"    fetched {len(publications)} ...")

        cursor = data.get("meta", {}).get("next_cursor")
        if not cursor:
            break

        time.sleep(0.12)

    return publications[:max_results]


def main():
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    all_pubs = []
    seen_ids: set = set()

    print("=" * 60)
    print("OpenAlex Collection — RentEasy KZ Assignment 4")
    print("=" * 60)

    for search in SEARCHES:
        print(f"\n[{search['label']}]")
        print(f"  Query: {search['query'][:70]}")
        results = fetch_works(search["query"], search["filters"], search["max"])

        added = 0
        for pub in results:
            pid = pub.get("id", "")
            if pid and pid not in seen_ids:
                seen_ids.add(pid)
                all_pubs.append(pub)
                added += 1

        print(f"  Added {added} new | Total unique: {len(all_pubs)}")
        time.sleep(1.0)

    # Save raw JSON
    out = DATA_DIR / "raw_publications.json"
    with open(out, "w", encoding="utf-8") as f:
        json.dump(all_pubs, f, ensure_ascii=False, indent=2)

    size_mb = out.stat().st_size / 1024 / 1024
    print(f"\n{'=' * 60}")
    print(f"Total unique publications: {len(all_pubs)}")
    print(f"Saved to: {out}  ({size_mb:.1f} MB)")


if __name__ == "__main__":
    main()
