"""
Assignment 2 — Task 1: Data Preparation
RentEasy KZ — Rental Market Dataset Cleaning

Steps:
1. Load merged dataset (Krisha.kz + OLX.kz)
2. Check and report: duplicates, missing values, outdated info, spelling errors
3. Remove outliers using 1.5×IQR method on numeric fields
4. Fix and clean data
5. Save cleaned CSV and cleaning log
"""

import pandas as pd
import numpy as np
import re
import os
from datetime import datetime

INPUT_FILE = "renteasy_data/renteasy_dataset.csv"
OUTPUT_FILE = "renteasy_data/renteasy_dataset_cleaned.csv"
LOG_FILE = "renteasy_data/cleaning_log.txt"

log_lines = []

def log(msg):
    print(msg)
    log_lines.append(msg)


# ─────────────────────────────────────────────
# 1. LOAD DATA
# ─────────────────────────────────────────────
log("=" * 60)
log("ASSIGNMENT 2 — DATA CLEANING LOG")
log(f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
log("=" * 60)

df = pd.read_csv(INPUT_FILE, encoding="utf-8-sig")
log(f"\n[LOAD] Rows: {len(df)}, Columns: {len(df.columns)}")
log(f"[LOAD] Columns: {list(df.columns)}")

original_count = len(df)


# ─────────────────────────────────────────────
# 2. CHECK DUPLICATES
# ─────────────────────────────────────────────
log("\n--- STEP 1: DUPLICATES ---")

# Full row duplicates
full_dupes = df.duplicated().sum()
log(f"Full row duplicates: {full_dupes}")

# Duplicate by (source, listing_id)
id_dupes = df.duplicated(subset=["source", "listing_id"]).sum()
log(f"Duplicate (source, listing_id): {id_dupes}")

# Remove duplicates by (source, listing_id), keep first
before = len(df)
df = df.drop_duplicates(subset=["source", "listing_id"], keep="first")
df = df.drop_duplicates(keep="first")
after = len(df)
log(f"Removed {before - after} duplicate rows → {after} rows remain")


# ─────────────────────────────────────────────
# 3. MISSING VALUES
# ─────────────────────────────────────────────
log("\n--- STEP 2: MISSING VALUES ---")

null_counts = df.isnull().sum()
null_pct = (df.isnull().sum() / len(df) * 100).round(1)
for col in df.columns:
    if null_counts[col] > 0:
        log(f"  {col}: {null_counts[col]} missing ({null_pct[col]}%)")

# Also check empty strings
empty_str_counts = (df == "").sum()
for col in df.columns:
    if empty_str_counts[col] > 0:
        log(f"  {col}: {empty_str_counts[col]} empty strings")

# Replace empty strings with NaN
df.replace("", np.nan, inplace=True)

# Force numeric types
numeric_cols = ["rooms", "area_m2", "floor", "total_floors", "price_tenge", "price_per_m2"]
for col in numeric_cols:
    df[col] = pd.to_numeric(df[col], errors="coerce")

# Drop rows with no price (critical field)
before = len(df)
df = df.dropna(subset=["price_tenge"])
log(f"\nDropped {before - len(df)} rows with missing price_tenge → {len(df)} rows remain")

# Drop rows with no city
before = len(df)
df = df.dropna(subset=["city"])
log(f"Dropped {before - len(df)} rows with missing city → {len(df)} rows remain")

# Fill missing seller_type with 'unknown'
filled_seller = df["seller_type"].isna().sum()
df["seller_type"] = df["seller_type"].fillna("unknown")
log(f"Filled {filled_seller} missing seller_type values with 'unknown'")

# Fill missing listing_date with today
filled_date = df["listing_date"].isna().sum()
df["listing_date"] = df["listing_date"].fillna(datetime.now().strftime("%Y-%m-%d"))
log(f"Filled {filled_date} missing listing_date values with today's date")


# ─────────────────────────────────────────────
# 4. OUTDATED LISTINGS
# ─────────────────────────────────────────────
log("\n--- STEP 3: OUTDATED LISTINGS ---")

df["listing_date"] = pd.to_datetime(df["listing_date"], errors="coerce")
cutoff_date = pd.Timestamp("2026-01-01")
before = len(df)
outdated = df[df["listing_date"] < cutoff_date]
if len(outdated) > 0:
    log(f"Removing {len(outdated)} listings older than 2026-01-01")
    df = df[df["listing_date"] >= cutoff_date]
else:
    log(f"No outdated listings found (all >= 2026-01-01)")
log(f"Date range: {df['listing_date'].min().date()} — {df['listing_date'].max().date()}")
df["listing_date"] = df["listing_date"].dt.strftime("%Y-%m-%d")


# ─────────────────────────────────────────────
# 5. SPELLING / STANDARDIZATION FIXES
# ─────────────────────────────────────────────
log("\n--- STEP 4: SPELLING & STANDARDIZATION ---")

# Standardize city names
city_fixes = {
    "Астана": "Астана",
    "Алматы": "Алматы",
    "Шымкент": "Шымкент",
    "Карагандa": "Кaрaганда",   # mixed Latin 'a' → Cyrillic
    "Кaрaганда": "Кaрaганда",
    "Актобе": "Актобе",
}
# Check for mixed-script city names
city_counts_before = df["city"].value_counts().to_dict()
df["city"] = df["city"].str.strip()
city_counts_after = df["city"].value_counts().to_dict()
log(f"Cities found: {list(df['city'].unique())}")

# Standardize seller_type
seller_valid = {"owner", "agency", "unknown"}
bad_sellers = df[~df["seller_type"].isin(seller_valid)]
if len(bad_sellers) > 0:
    log(f"Non-standard seller_type values: {bad_sellers['seller_type'].value_counts().to_dict()}")
    df["seller_type"] = df["seller_type"].apply(
        lambda x: x if x in seller_valid else "unknown"
    )

# Standardize source
source_valid = {"krisha", "olx"}
df["source"] = df["source"].str.lower().str.strip()
bad_source = df[~df["source"].isin(source_valid)]
if len(bad_source) > 0:
    log(f"Unexpected source values: {bad_source['source'].value_counts().to_dict()}")

# Fix rooms — round floats to int where appropriate
df["rooms"] = df["rooms"].apply(lambda x: int(round(x)) if pd.notna(x) else np.nan)

# Validate rooms range (0 = studio, max sensible = 8)
invalid_rooms = df[df["rooms"].notna() & ((df["rooms"] < 0) | (df["rooms"] > 10))]
if len(invalid_rooms) > 0:
    log(f"Invalid rooms values (outside 0-10): {len(invalid_rooms)} — setting to NaN")
    df.loc[(df["rooms"] < 0) | (df["rooms"] > 10), "rooms"] = np.nan

log(f"Seller type distribution: {df['seller_type'].value_counts().to_dict()}")
log(f"Source distribution: {df['source'].value_counts().to_dict()}")


# ─────────────────────────────────────────────
# 6. NUMERIC VALIDATION — 1.5×IQR OUTLIER REMOVAL
# ─────────────────────────────────────────────
log("\n--- STEP 5: OUTLIER REMOVAL (1.5×IQR) ---")

def remove_outliers_iqr(df, col, multiplier=1.5):
    """Remove outliers outside [Q1 - 1.5*IQR, Q3 + 1.5*IQR]."""
    series = df[col].dropna()
    Q1 = series.quantile(0.25)
    Q3 = series.quantile(0.75)
    IQR = Q3 - Q1
    lower = Q1 - multiplier * IQR
    upper = Q3 + multiplier * IQR
    before = len(df)
    mask = df[col].isna() | ((df[col] >= lower) & (df[col] <= upper))
    df_clean = df[mask].copy()
    removed = before - len(df_clean)
    log(f"  {col}: Q1={Q1:.0f}, Q3={Q3:.0f}, IQR={IQR:.0f} → range [{lower:.0f}, {upper:.0f}] → removed {removed} outliers")
    return df_clean

# Apply IQR to price_tenge
df = remove_outliers_iqr(df, "price_tenge")

# Apply IQR to area_m2 (only for Krisha records that have it)
df = remove_outliers_iqr(df, "area_m2")

# Apply IQR to price_per_m2
df = remove_outliers_iqr(df, "price_per_m2")

log(f"\nAfter outlier removal: {len(df)} rows")


# ─────────────────────────────────────────────
# 7. RECALCULATE DERIVED FIELDS
# ─────────────────────────────────────────────
log("\n--- STEP 6: RECALCULATE DERIVED FIELDS ---")

# Recalculate price_per_m2 from clean price + area
df["price_per_m2"] = np.where(
    df["area_m2"].notna() & (df["area_m2"] > 0),
    (df["price_tenge"] / df["area_m2"]).round(1),
    np.nan
)

# Add price segment column
def assign_price_segment(price):
    if pd.isna(price):
        return "unknown"
    if price < 120_000:
        return "low"
    elif price < 250_000:
        return "middle"
    elif price < 500_000:
        return "high"
    else:
        return "luxury"

df["price_segment"] = df["price_tenge"].apply(assign_price_segment)
seg_counts = df["price_segment"].value_counts().to_dict()
log(f"Price segments: {seg_counts}")


# ─────────────────────────────────────────────
# 8. RESET INDEX AND REASSIGN IDs
# ─────────────────────────────────────────────
df = df.reset_index(drop=True)
df["id"] = df.index + 1

# Reorder columns
col_order = [
    "id", "source", "city", "district", "rooms", "area_m2", "floor",
    "total_floors", "price_tenge", "price_per_m2", "price_segment",
    "seller_type", "listing_date", "title", "address", "url", "listing_id"
]
df = df[[c for c in col_order if c in df.columns]]


# ─────────────────────────────────────────────
# 9. FINAL SUMMARY
# ─────────────────────────────────────────────
log("\n--- FINAL SUMMARY ---")
log(f"Original records:  {original_count}")
log(f"Cleaned records:   {len(df)}")
log(f"Records removed:   {original_count - len(df)}")
log(f"\nCity distribution:")
for city, cnt in df["city"].value_counts().items():
    log(f"  {city}: {cnt}")
log(f"\nSource distribution:")
for src, cnt in df["source"].value_counts().items():
    log(f"  {src}: {cnt}")
log(f"\nMissing values in final dataset:")
final_nulls = df.isnull().sum()
for col in df.columns:
    if final_nulls[col] > 0:
        pct = final_nulls[col] / len(df) * 100
        log(f"  {col}: {final_nulls[col]} ({pct:.1f}%)")
log(f"\nNumeric fields summary:")
for col in ["price_tenge", "area_m2", "price_per_m2"]:
    if col in df.columns:
        s = df[col].dropna()
        log(f"  {col}: min={s.min():.0f}, max={s.max():.0f}, mean={s.mean():.0f}, median={s.median():.0f}")


# ─────────────────────────────────────────────
# 10. SAVE
# ─────────────────────────────────────────────
os.makedirs("renteasy_data", exist_ok=True)
df.to_csv(OUTPUT_FILE, index=False, encoding="utf-8-sig")
log(f"\n[SAVED] Cleaned CSV → {OUTPUT_FILE} ({len(df)} rows)")

with open(LOG_FILE, "w", encoding="utf-8") as f:
    f.write("\n".join(log_lines))
log(f"[SAVED] Cleaning log → {LOG_FILE}")
