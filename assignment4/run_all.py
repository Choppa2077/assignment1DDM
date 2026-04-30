#!/usr/bin/env python3
"""Run all Assignment 4 scripts in sequence."""
import subprocess
import sys
import time

STEPS = [
    ("01_collect_data.py",    "Task 1 — OpenAlex data collection (~5-10 min)"),
    ("02_clean_data.py",      "Task 1 — Data cleaning"),
    ("03_visualize.py",       "Task 2 — Dashboards"),
    ("04_network_analysis.py","Task 2 — Network analysis"),
    ("05_topic_modeling.py",  "Task 2 — Topic modeling"),
    ("06_generate_report.py", "Task 3 — DOCX report"),
]

print("=" * 65)
print("  RentEasy KZ — Assignment 4 Pipeline")
print("=" * 65)

for script, label in STEPS:
    print(f"\n▶  {label}")
    print(f"   Running: python {script}")
    t0 = time.time()
    result = subprocess.run([sys.executable, script])
    elapsed = time.time() - t0
    if result.returncode != 0:
        print(f"\n✗  {script} failed (exit {result.returncode}). Fix and re-run.")
        sys.exit(result.returncode)
    print(f"   ✓ done in {elapsed:.1f}s")

print("\n" + "=" * 65)
print("  All steps complete!")
print("  Outputs:")
print("    data/cleaned_publications.csv")
print("    output/figures/    (5 dashboards)")
print("    output/networks/   (co-occurrence network)")
print("    output/topics/     (LDA topics)")
print("    output/report/     (DOCX report)")
print("=" * 65)
