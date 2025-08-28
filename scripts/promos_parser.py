#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Bottlemart promos parser (robust positional, headerless fallback).

- Reads newest PDF from ../bottlemart_promos (sibling of scripts/) unless --pdf is provided.
- Extracts product rows: Code, Product, Retail Price Inc GST.
- Preferred: header-based page detection (when headers are text).
- Fallback: user-provided page ranges via --category-ranges (when headers are images or "ghost" text).

Categories (start -> stop, stop exclusive in the "conceptual" sense):
- ALM BEER              (ends at CUB BEER)
- CUB BEER              (ends at LION BEER)
- LION BEER             (ends at CRAFT BEER PROGRAM – ALM. OPTIONAL)
- ALM CIDER             (ends at SPIRITS - SINGLE SELL)
- SPIRITS - SINGLE SELL (ends at RTDS 'ANY' MULTIS (EG ANY 2-FOR))
- RTDS - SINGLE SELL    (ends at WINE 'ANY' MULTIS (EG ANY 2-FOR))
- WINE - SINGLE SELL    (ends at SPARKLING WINE)
- SPARKLING WINE        (ends at CASK 'ANY' MULTIS (EG ANY 2-FOR))

Output:
  - ONE Excel: bottlemart_products_all.xlsx (sheet: ALL_PRODUCTS) saved next to the PDF.

Usage (from scripts/):
  pip install pdfplumber pandas openpyxl
  # 1) Inspect page stats
  python promos_parser.py --print-page-stats
  # 2) Extract using manual ranges if headers not detected:
  python promos_parser.py --category-ranges "ALM BEER:6-8|CUB BEER:9-10|LION BEER:11-12|ALM CIDER:13-13|SPIRITS - SINGLE SELL:14-15|RTDS - SINGLE SELL:16-16|WINE - SINGLE SELL:17-18|SPARKLING WINE:19-20"
"""

import argparse
import re
import time
from pathlib import Path
from typing import List, Tuple, Optional, Dict

import pdfplumber
import pandas as pd

# ------------------ Regex & heuristics ------------------
PRICE_RE = re.compile(r"\$?\d{1,3}(?:,\d{3})*\.\d{2}")
CODE_AT_START_RE = re.compile(r"^\d{4,}$")

# Default code column thresholds (auto-calibrated unless --no-autocalib)
MIN_CODE_X1_DEFAULT = 60
MAX_CODE_X1_DEFAULT = 190
AUTO_MARGIN = 40  # +/- px around median x1

# Category order (canonical names)
CATEGORY_ORDER = [
    "ALM BEER",
    "CUB BEER",
    "LION BEER",
    "ALM CIDER",
    "SPIRITS - SINGLE SELL",
    "RTDS - SINGLE SELL",
    "WINE - SINGLE SELL",
    "SPARKLING WINE",
]

# ------------------ Positional grouping ------------------
def group_words_into_lines(words, y_tol: float = 2.2) -> List[list]:
    """Group words by 'top' proximity."""
    words = sorted(words, key=lambda w: (w["top"], w["x0"]))
    lines, current = [], []
    def same_line(a, b): return abs(a["top"] - b["top"]) <= y_tol
    for w in words:
        if not current:
            current = [w]; continue
        if same_line(current[-1], w): current.append(w)
        else:
            lines.append(sorted(current, key=lambda ww: ww["x0"]))
            current = [w]
    if current: lines.append(sorted(current, key=lambda ww: ww["x0"]))
    return lines

def line_text(line: List[dict]) -> str:
    return " ".join(w["text"] for w in line).strip()

# ------------------ Filesystem helpers ------------------
def find_latest_pdf(basedir: Path) -> Optional[Path]:
    pdfs = sorted(basedir.glob("*.pdf"), key=lambda p: p.stat().st_mtime, reverse=True)
    return pdfs[0] if pdfs else None

def resolve_pdf_path(user_arg: Optional[str]) -> Path:
    promos_dir = (Path(__file__).parent / ".." / "bottlemart_promos").resolve()
    if user_arg:
        pdf_path = (promos_dir / user_arg).resolve()
        if not pdf_path.exists():
            raise FileNotFoundError(f"PDF not found at: {pdf_path}")
        return pdf_path
    latest = find_latest_pdf(promos_dir)
    if latest is None:
        raise FileNotFoundError(f"No PDF found in: {promos_dir}")
    return latest.resolve()

# ------------------ Auto-calibration ------------------
def autocalibrate_code_x1(pdf_path: Path) -> Tuple[float, float]:
    """Find median x1 of first numeric token per line; return min/max around it."""
    xs = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words(extra_attrs=["x0","x1","top","bottom","size"]) or []
            lines = group_words_into_lines(words)
            for ln in lines:
                if not ln: continue
                first = ln[0]["text"].strip()
                if CODE_AT_START_RE.match(first):
                    xs.append(ln[0]["x1"])
    if not xs:
        return (MIN_CODE_X1_DEFAULT, MAX_CODE_X1_DEFAULT)
    xs.sort()
    med = xs[len(xs)//2]
    return (med - AUTO_MARGIN, med + AUTO_MARGIN)

# ------------------ Row extraction ------------------
def extract_row_from_line(
    line: List[dict],
    min_code_x1: float,
    max_code_x1: float
) -> Optional[Tuple[str, str, float]]:
    if not line:
        return None
    first = line[0]
    if not CODE_AT_START_RE.match(first["text"].strip()):
        return None
    if not (min_code_x1 <= first["x1"] <= max_code_x1):
        return None

    price_tokens = [w for w in line if PRICE_RE.match(w["text"].replace(",", ""))]
    if not price_tokens:
        return None
    retail_token = price_tokens[-1]
    retail = float(retail_token["text"].replace("$", "").replace(",", ""))

    first_price_x0 = price_tokens[0]["x0"]
    name_tokens = [w for w in line[1:] if w["x0"] < first_price_x0]
    name = " ".join(w["text"] for w in name_tokens).strip()
    name = re.sub(r"\s{2,}", " ", name)

    code = first["text"].strip()
    if not code or not name:
        return None
    return code, name, retail

def extract_rows_in_pages(pdf_path: Path, page_start: int, page_end_inclusive: int,
                          min_code_x1: float, max_code_x1: float) -> pd.DataFrame:
    """Extract rows (Code, Product, Retail) for a page range [start..end]."""
    rows = []
    with pdfplumber.open(pdf_path) as pdf:
        page_start = max(0, page_start)
        page_end_inclusive = min(page_end_inclusive, len(pdf.pages)-1)
        for p in range(page_start, page_end_inclusive + 1):
            page = pdf.pages[p]
            words = page.extract_words(extra_attrs=["x0","x1","top","bottom","size"]) or []
            lines = group_words_into_lines(words, y_tol=2.2)
            for ln in lines:
                parsed = extract_row_from_line(ln, min_code_x1, max_code_x1)
                if parsed:
                    rows.append((p, *parsed))
    df = pd.DataFrame(rows, columns=["Page","Code","Product","Retail Price Inc GST ($)"])
    return df

# ------------------ Page stats ------------------
def print_page_stats(pdf_path: Path, min_code_x1: float, max_code_x1: float, sample_k: int = 3):
    with pdfplumber.open(pdf_path) as pdf:
        n = len(pdf.pages)
        print(f"[INFO] Document has {n} pages (0-based indices).")
        for p in range(n):
            page = pdf.pages[p]
            words = page.extract_words(extra_attrs=["x0","x1","top","bottom","size"]) or []
            lines = group_words_into_lines(words, y_tol=2.2)
            rows = []
            for ln in lines:
                pr = extract_row_from_line(ln, min_code_x1, max_code_x1)
                if pr:
                    rows.append(pr)
            print(f"[PAGE {p}] product-like rows: {len(rows)}")
            for ex in rows[:sample_k]:
                print(f"   - Code: {ex[0]} | Name: {ex[1][:60]} | Retail: {ex[2]}")

# ------------------ Category ranges parsing ------------------
def parse_category_ranges(arg: str) -> Dict[str, Tuple[int,int]]:
    """
    Parse --category-ranges string like:
      "ALM BEER:6-8|CUB BEER:9-10|LION BEER:11-12|ALM CIDER:13-13|SPIRITS - SINGLE SELL:14-15|RTDS - SINGLE SELL:16-16|WINE - SINGLE SELL:17-18|SPARKLING WINE:19-20"
    Returns dict {category: (start,end)} with 0-based page indices expected from the user input.
    """
    out = {}
    if not arg:
        return out
    parts = [p.strip() for p in arg.split("|") if p.strip()]
    for part in parts:
        if ":" not in part or "-" not in part:
            raise ValueError(f"Invalid --category-ranges chunk: {part}")
        cat, rng = part.split(":", 1)
        a, b = rng.split("-", 1)
        start = int(a.strip())
        end = int(b.strip())
        if end < start:
            raise ValueError(f"Invalid range for {cat}: {start}-{end}")
        out[cat.strip()] = (start, end)
    return out

# ------------------ MAIN ------------------
def main():
    parser = argparse.ArgumentParser(description="Extract Bottlemart categories and save ONE Excel next to the PDF.")
    parser.add_argument("--pdf", type=str, default=None, help="PDF filename inside ../bottlemart_promos (if omitted, newest PDF is used).")
    parser.add_argument("--print-page-stats", action="store_true", help="Print product-like row stats per page to help decide ranges.")
    parser.add_argument("--category-ranges", type=str, default=None,
                        help="Manual page ranges per category. Example: \"ALM BEER:6-8|CUB BEER:9-10|...|SPARKLING WINE:19-20\"")
    parser.add_argument("--no-autocalib", action="store_true", help="Disable auto-calibration of code x1 bounds.")
    parser.add_argument("--min-code-x1", type=float, default=MIN_CODE_X1_DEFAULT, help="Manual min x1 for code column (if --no-autocalib).")
    parser.add_argument("--max-code-x1", type=float, default=MAX_CODE_X1_DEFAULT, help="Manual max x1 for code column (if --no-autocalib).")
    args = parser.parse_args()

    t0 = time.time()
    pdf_path = resolve_pdf_path(args.pdf)
    pdf_dir = pdf_path.parent
    print(f"[INFO] Using PDF: {pdf_path}")

    # Auto-calibrate code x1 thresholds
    if args.no_autocalib:
        min_x1, max_x1 = args.min_code_x1, args.max_code_x1
        print(f"[INFO] Auto-calibration disabled. Using code x1 bounds: [{min_x1}, {max_x1}]")
    else:
        min_x1, max_x1 = autocalibrate_code_x1(pdf_path)
        print(f"[INFO] Auto-calibrated code x1 bounds: [{min_x1:.1f}, {max_x1:.1f}]")

    # Page stats mode
    if args.print_page_stats:
        print_page_stats(pdf_path, min_x1, max_x1, sample_k=3)
        # No exit: allow combining with extraction if user also provided ranges.

    # Parse manual ranges (if any)
    ranges = parse_category_ranges(args.category_ranges) if args.category_ranges else {}

    all_rows = []
    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)

    # Iterate categories in order; if a range is provided, use it; else warn & skip
    for cat in CATEGORY_ORDER:
        if cat not in ranges:
            print(f"\n[WARN] No page range provided for category '{cat}'. Skipping this category.")
            continue
        start, end = ranges[cat]
        if start < 0 or end >= total_pages:
            print(f"[WARN] Range {start}-{end} for '{cat}' is out of document bounds (0..{total_pages-1}). Skipping.")
            continue
        print(f"\n[INFO] === Extracting '{cat}' in pages {start}..{end} ===")
        df_cat = extract_rows_in_pages(pdf_path, start, end, min_x1, max_x1)
        if df_cat.empty:
            print(f"[WARN] No rows found in pages {start}..{end} for '{cat}'.")
            continue
        df_cat.insert(0, "Category", cat)
        print(f"[INFO] Rows extracted: {len(df_cat)}")
        all_rows.append(df_cat)

    if not all_rows:
        print("[ERROR] No products extracted for any category. Use --print-page-stats to identify ranges and pass --category-ranges.")
        return

    # Concat & tidy
    df_all = pd.concat(all_rows, ignore_index=True)
    before = len(df_all)
    df_all = df_all.drop_duplicates(subset=["Category","Page","Code","Product","Retail Price Inc GST ($)"]).reset_index(drop=True)
    after = len(df_all)
    print(f"\n[INFO] ALL_PRODUCTS rows after de-dup: {after} (from {before}).")

    # Save ONE Excel next to the PDF
    out_xlsx = (pdf_dir / "bottlemart_products_all.xlsx").resolve()
    print(f"[INFO] Writing Excel to: {out_xlsx}")
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        df_all.to_excel(writer, sheet_name="ALL_PRODUCTS", index=False)

    dt = time.time() - t0
    print(f"[SUCCESS] Done. File saved to: {out_xlsx}")
    print(f"[INFO] Elapsed time: {dt:.2f}s")


if __name__ == "__main__":
    main()
