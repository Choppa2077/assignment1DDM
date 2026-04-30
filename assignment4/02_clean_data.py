#!/usr/bin/env python3
"""
Assignment 4 - Task 1: Data Cleaning and Dataset Construction
"""
import json
import re
import unicodedata
from pathlib import Path
from collections import Counter

import numpy as np
import pandas as pd

DATA_DIR = Path("data")

COUNTRY_CODES = {
    "US": "United States", "CN": "China", "GB": "United Kingdom",
    "DE": "Germany", "IN": "India", "AU": "Australia", "FR": "France",
    "CA": "Canada", "KR": "South Korea", "IT": "Italy", "JP": "Japan",
    "NL": "Netherlands", "ES": "Spain", "BR": "Brazil", "SG": "Singapore",
    "SE": "Sweden", "CH": "Switzerland", "PL": "Poland", "RU": "Russia",
    "KZ": "Kazakhstan", "PK": "Pakistan", "MY": "Malaysia", "ZA": "South Africa",
    "PT": "Portugal", "TR": "Turkey", "NG": "Nigeria", "EG": "Egypt",
    "SA": "Saudi Arabia", "IL": "Israel", "IR": "Iran", "TW": "Taiwan",
    "HK": "Hong Kong", "BE": "Belgium", "AT": "Austria", "NO": "Norway",
    "DK": "Denmark", "FI": "Finland", "CZ": "Czech Republic", "HU": "Hungary",
    "GR": "Greece", "RO": "Romania", "MX": "Mexico", "AR": "Argentina",
    "ID": "Indonesia", "TH": "Thailand", "VN": "Vietnam", "PH": "Philippines",
    "NZ": "New Zealand", "IE": "Ireland", "UA": "Ukraine",
}


def reconstruct_abstract(inv: dict) -> str:
    if not inv or not isinstance(inv, dict):
        return ""
    pairs = []
    for word, positions in inv.items():
        for pos in positions:
            pairs.append((pos, word))
    pairs.sort(key=lambda x: x[0])
    return " ".join(w for _, w in pairs)


def norm(text: str) -> str:
    if not text:
        return ""
    text = unicodedata.normalize("NFKC", str(text))
    return " ".join(text.strip().split())


def norm_journal(name: str) -> str:
    name = norm(name)
    name = re.sub(r"\s*\(.*?\)", "", name).strip()
    return name


def extract_record(pub: dict) -> dict:
    title = norm(pub.get("title", ""))
    year = pub.get("publication_year")
    doi = (pub.get("doi") or "").replace("https://doi.org/", "").strip()

    authorships = pub.get("authorships") or []
    authors, countries_raw = [], []

    for auth in authorships:
        name = norm((auth.get("author") or {}).get("display_name", "")).strip(".,;:")
        if name:
            authors.append(name)
        for inst in auth.get("institutions") or []:
            code = inst.get("country_code", "")
            if code:
                countries_raw.append(COUNTRY_CODES.get(code, code))

    authors_str = "; ".join(authors[:10])
    first_author = authors[0] if authors else ""
    countries_str = "; ".join(sorted(set(countries_raw))) if countries_raw else ""
    primary_country = countries_raw[0] if countries_raw else "Unknown"

    journal = ""
    primary_loc = pub.get("primary_location") or {}
    source = primary_loc.get("source") or {}
    if source:
        journal = norm_journal(source.get("display_name", ""))

    cited = pub.get("cited_by_count") or 0

    kw_list = []
    for kw in pub.get("keywords") or []:
        k = norm(kw.get("keyword", kw) if isinstance(kw, dict) else kw).lower()
        if k:
            kw_list.append(k)

    concept_names = []
    for c in sorted(pub.get("concepts") or [], key=lambda x: x.get("score", 0), reverse=True)[:5]:
        cn = norm(c.get("display_name", ""))
        if cn:
            concept_names.append(cn)

    abstract = reconstruct_abstract(pub.get("abstract_inverted_index") or {})
    is_oa = (pub.get("open_access") or {}).get("is_oa", False)

    return {
        "id": pub.get("id", ""),
        "doi": doi,
        "title": title,
        "year": int(year) if year else None,
        "first_author": first_author,
        "authors": authors_str,
        "num_authors": len(authors),
        "journal": journal or "Unknown Journal",
        "cited_by_count": int(cited),
        "keywords": "; ".join(dict.fromkeys(kw_list[:10])),
        "concepts": "; ".join(concept_names),
        "country": countries_str or "Unknown",
        "primary_country": primary_country,
        "abstract": abstract[:1200] if abstract else "",
        "is_open_access": bool(is_oa),
    }


def main():
    raw_path = DATA_DIR / "raw_publications.json"
    print(f"Loading {raw_path} ...")
    with open(raw_path, "r", encoding="utf-8") as f:
        raw = json.load(f)
    print(f"  {len(raw)} raw records loaded")

    records = [extract_record(p) for p in raw]
    df = pd.DataFrame(records)
    before = len(df)

    # — dedup by DOI
    has_doi = df["doi"].notna() & (df["doi"] != "")
    df_doi = df[has_doi].drop_duplicates(subset=["doi"])
    df_no_doi = df[~has_doi]

    # — dedup by title (case-insensitive)
    df_no_doi = df_no_doi.copy()
    df_no_doi["_tl"] = df_no_doi["title"].str.lower().str.strip()
    df_no_doi = df_no_doi.drop_duplicates(subset=["_tl"]).drop(columns=["_tl"])

    df = pd.concat([df_doi, df_no_doi], ignore_index=True)
    df["_tl"] = df["title"].str.lower().str.strip()
    df = df.drop_duplicates(subset=["_tl"]).drop(columns=["_tl"])

    print(f"  After deduplication: {len(df)} (removed {before - len(df)})")

    # — drop missing title / year
    df = df[df["title"].notna() & (df["title"] != "")]
    df = df[df["year"].notna()]
    df["year"] = df["year"].astype(int)
    df = df[(df["year"] >= 2010) & (df["year"] <= 2025)]

    # — fill missing
    df["journal"] = df["journal"].replace("", "Unknown Journal").fillna("Unknown Journal")
    df["country"] = df["country"].replace("", "Unknown").fillna("Unknown")
    df["primary_country"] = df["primary_country"].replace("", "Unknown").fillna("Unknown")
    df["cited_by_count"] = df["cited_by_count"].fillna(0).astype(int)
    df["keywords"] = df["keywords"].fillna("")
    df["abstract"] = df["abstract"].fillna("")

    df["log_citations"] = np.log1p(df["cited_by_count"])

    out = DATA_DIR / "cleaned_publications.csv"
    df.to_csv(out, index=False, encoding="utf-8-sig")

    print(f"\n{'=' * 50}")
    print(f"Final dataset: {len(df)} publications")
    print(f"  Year range   : {df['year'].min()} – {df['year'].max()}")
    print(f"  Journals     : {df['journal'].nunique()}")
    print(f"  First authors: {df['first_author'].nunique()}")
    print(f"  Countries    : {df['primary_country'].nunique()}")
    print(f"  Avg citations: {df['cited_by_count'].mean():.1f}")
    print(f"  With abstract: {(df['abstract'] != '').sum()}")
    print(f"Saved → {out}")


if __name__ == "__main__":
    main()
