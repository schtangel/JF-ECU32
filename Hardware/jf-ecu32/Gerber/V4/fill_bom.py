#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Fill 'Mfr. Part #' (MPN) and 'Manufacturer' in a merged BOM by scraping LCSC product pages.

Usage:
  pip install pandas requests beautifulsoup4 lxml openpyxl tqdm
  python fill_mpn_from_lcsc.py merged_bom.xlsx
  python fill_mpn_from_lcsc.py merged_bom.csv

Output:
  <input>_with_mpn.xlsx
  <input>_with_mpn.csv

Notes:
- Builds LCSC Link as: https://www.lcsc.com/product-detail/C1590.html
- Lots of console logging to verify it really fetches pages and parses fields.
"""

import sys
import re
import time
import json
from pathlib import Path

import pandas as pd
import requests
from bs4 import BeautifulSoup
from tqdm import tqdm


def is_empty_cell(v) -> bool:
    """Treat NaN/None/'nan'/'none'/'-' etc as empty."""
    if v is None:
        return True
    try:
        if pd.isna(v):
            return True
    except Exception:
        pass
    s = str(v).strip().lower()
    return s in ("", "nan", "none", "null", "-", "n/a", "na")


def lcsc_product_url(lcsc_pn: str) -> str:
    lcsc_pn = str(lcsc_pn).strip()
    return f"https://www.lcsc.com/product-detail/{lcsc_pn}.html"


def _tree_search_for_fields(data):
    """
    Heuristic search through JSON for keys that might contain manufacturer / mpn.
    LCSC structures change; this is purposely broad.
    """
    found = {"manufacturer": None, "mpn": None}

    # candidate key sets (lowercased)
    mfr_keys = {"manufacturer", "mfr", "brandname", "brand", "manufacturername"}
    mpn_keys = {
        "manufacturerpartnumber",
        "mfrpartnumber",
        "mfr_part_number",
        "mfrpartno",
        "mpn",
        "partnumber",
        "manufacturerpn",
    }

    stack = [data]
    while stack:
        obj = stack.pop()
        if isinstance(obj, dict):
            for k, v in obj.items():
                lk = str(k).lower()

                # keep first decent-looking strings
                if found["manufacturer"] is None and lk in mfr_keys:
                    if isinstance(v, str) and v.strip() and len(v.strip()) < 120:
                        found["manufacturer"] = v.strip()

                if found["mpn"] is None and lk in mpn_keys:
                    if isinstance(v, str) and v.strip() and len(v.strip()) < 200:
                        found["mpn"] = v.strip()

                stack.append(v)
        elif isinstance(obj, list):
            stack.extend(obj)

        if found["manufacturer"] and found["mpn"]:
            break

    return found


def extract_from_next_data(html: str):
    """
    Try to parse Next.js JSON data embedded in the page.
    """
    # Next.js __NEXT_DATA__
    m = re.search(
        r'<script id="__NEXT_DATA__" type="application/json">\s*(\{.*?\})\s*</script>',
        html,
        re.S,
    )
    if not m:
        return {"manufacturer": None, "mpn": None, "source": None}

    try:
        data = json.loads(m.group(1))
    except Exception:
        return {"manufacturer": None, "mpn": None, "source": None}

    found = _tree_search_for_fields(data)
    if found["manufacturer"] or found["mpn"]:
        found["source"] = "__NEXT_DATA__"
        return found

    return {"manufacturer": None, "mpn": None, "source": None}


def extract_from_html_text(html: str):
    """
    Fallback: look for MPN/manufacturer labels in visible text.
    """
    soup = BeautifulSoup(html, "lxml")
    text = soup.get_text("\n", strip=True)

    mpn = None
    manufacturer = None

    mpn_patterns = [
        r"(Manufacturer\s*Part\s*Number)\s*[:：]?\s*([^\n]+)",
        r"(Mfr\.?\s*Part\s*#)\s*[:：]?\s*([^\n]+)",
        r"(Mfr\.?\s*Part\s*No\.?)\s*[:：]?\s*([^\n]+)",
        r"(MPN)\s*[:：]?\s*([^\n]+)",
    ]
    for pat in mpn_patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            cand = m.group(2).strip()
            cand = re.split(r"\s{2,}|\n", cand)[0].strip()
            if cand and len(cand) < 200:
                mpn = cand
                break

    mfr_patterns = [
        r"(Manufacturer)\s*[:：]?\s*([^\n]+)",
        r"(Brand)\s*[:：]?\s*([^\n]+)",
    ]
    for pat in mfr_patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            cand = m.group(2).strip()
            cand = re.split(r"\s{2,}|\n", cand)[0].strip()
            if cand and len(cand) < 120:
                manufacturer = cand
                break

    return {"manufacturer": manufacturer, "mpn": mpn, "source": "html_text" if (manufacturer or mpn) else None}


def looks_like_antibot(html: str) -> bool:
    s = html.lower()
    # crude heuristics; adjust if needed
    keywords = [
        "captcha",
        "please verify",
        "security check",
        "are you a human",
        "cloudflare",
        "bot detection",
        "incapsula",
        "access denied",
    ]
    return any(k in s for k in keywords)


def fetch_lcsc_fields(session: requests.Session, lcsc_pn: str, timeout=25, debug=False):
    url = lcsc_product_url(lcsc_pn)

    t0 = time.time()
    r = session.get(url, timeout=timeout, allow_redirects=True)
    dt = time.time() - t0

    if debug:
        print(f"[HTTP] PN={lcsc_pn} status={r.status_code} time={dt:.2f}s final_url={r.url} bytes={len(r.text)}")

    r.raise_for_status()

    if looks_like_antibot(r.text):
        raise RuntimeError("Anti-bot/captcha page detected")

    # A) JSON
    a = extract_from_next_data(r.text)
    if a.get("manufacturer") or a.get("mpn"):
        if debug:
            print(f"[PARSE] PN={lcsc_pn} source={a.get('source')} manufacturer={a.get('manufacturer')} mpn={a.get('mpn')}")
        return {"manufacturer": a.get("manufacturer"), "mpn": a.get("mpn"), "source": a.get("source")}, r.url

    # B) HTML text
    b = extract_from_html_text(r.text)
    if b.get("manufacturer") or b.get("mpn"):
        if debug:
            print(f"[PARSE] PN={lcsc_pn} source={b.get('source')} manufacturer={b.get('manufacturer')} mpn={b.get('mpn')}")
        return {"manufacturer": b.get("manufacturer"), "mpn": b.get("mpn"), "source": b.get("source")}, r.url

    if debug:
        print(f"[PARSE] PN={lcsc_pn} no fields found (page parsed, but keys/labels not detected)")
    return {"manufacturer": None, "mpn": None, "source": None}, r.url


def main():
    if len(sys.argv) < 2:
        print("Usage: python fill_mpn_from_lcsc.py <input.xlsx|input.csv>")
        sys.exit(1)

    in_path = Path(sys.argv[1])
    if not in_path.exists():
        raise FileNotFoundError(in_path)

    # --- load
    if in_path.suffix.lower() in (".xlsx", ".xls"):
        df = pd.read_excel(in_path)
    else:
        df = pd.read_csv(in_path)

    # --- ensure columns
    if "LCSC Part #" not in df.columns:
        raise ValueError("Expected column 'LCSC Part #'")

    if "Mfr. Part #" not in df.columns:
        df["Mfr. Part #"] = ""

    if "Manufacturer" not in df.columns:
        df["Manufacturer"] = ""

    # Always rebuild correct link format (as requested)
    df["LCSC Link"] = df["LCSC Part #"].apply(
        lambda x: lcsc_product_url(x) if not is_empty_cell(x) else ""
    )

    # --- stats
    total_rows = len(df)
    unique_pn = sorted({str(x).strip() for x in df["LCSC Part #"] if not is_empty_cell(x)})
    empty_mpn_rows = sum(is_empty_cell(v) for v in df["Mfr. Part #"])
    filled_mpn_rows = total_rows - empty_mpn_rows

    print("=== INPUT SUMMARY ===")
    print(f"File: {in_path}")
    print(f"Rows: {total_rows}")
    print(f"Unique LCSC PN: {len(unique_pn)}")
    print(f"Mfr. Part # empty rows: {empty_mpn_rows} (filled: {filled_mpn_rows})")
    print("First 10 unique PN:", unique_pn[:10])
    print("=====================")

    # --- session
    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                      "AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/120.0.0.0 Safari/537.36",
        "Accept-Language": "en-US,en;q=0.9",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Connection": "keep-alive",
    })

    # --- cache
    cache = {}  # pn -> fields dict

    # Debug options
    DEBUG_EVERY = 1          # print per-row details
    SLEEP_SEC = 0.60         # be polite, reduce antibot chance
    MAX_ERRORS_BEFORE_STOP = 30

    errors = 0
    parsed_ok = 0
    parsed_empty = 0
    skipped_already_filled = 0
    skipped_no_pn = 0
    cached_hits = 0

    print("=== START FETCH ===")
    for i in tqdm(range(total_rows), desc="Fetching LCSC"):
        pn_raw = df.at[i, "LCSC Part #"]
        if is_empty_cell(pn_raw):
            skipped_no_pn += 1
            if DEBUG_EVERY and (i % DEBUG_EVERY == 0):
                print(f"[SKIP] row={i} no LCSC PN")
            continue

        pn = str(pn_raw).strip()

        # skip if already filled (but handle NaN correctly)
        if not is_empty_cell(df.at[i, "Mfr. Part #"]):
            skipped_already_filled += 1
            if DEBUG_EVERY and (i % DEBUG_EVERY == 0):
                print(f"[SKIP] row={i} PN={pn} already has Mfr. Part #='{df.at[i, 'Mfr. Part #']}'")
            continue

        if pn in cache:
            fields = cache[pn]
            cached_hits += 1
            if DEBUG_EVERY and (i % DEBUG_EVERY == 0):
                print(f"[CACHE] row={i} PN={pn} -> manufacturer={fields.get('manufacturer')} mpn={fields.get('mpn')} source={fields.get('source')}")
        else:
            try:
                fields, final_url = fetch_lcsc_fields(session, pn, timeout=25, debug=True)
                fields["final_url"] = final_url
                cache[pn] = fields
            except Exception as e:
                errors += 1
                print(f"[ERROR] row={i} PN={pn} err={type(e).__name__}: {e}")
                cache[pn] = {"manufacturer": None, "mpn": None, "source": None, "error": str(e)}
                if errors >= MAX_ERRORS_BEFORE_STOP:
                    print(f"[STOP] too many errors ({errors}). Stopping.")
                    break
            time.sleep(SLEEP_SEC)

        # write back
        mpn_val = cache[pn].get("mpn")
        mfr_val = cache[pn].get("manufacturer")

        if mpn_val:
            df.at[i, "Mfr. Part #"] = mpn_val
        if mfr_val:
            df.at[i, "Manufacturer"] = mfr_val

        if mpn_val or mfr_val:
            parsed_ok += 1
            if DEBUG_EVERY and (i % DEBUG_EVERY == 0):
                print(f"[WRITE] row={i} PN={pn} Manufacturer='{mfr_val}' MPN='{mpn_val}' source={cache[pn].get('source')}")
        else:
            parsed_empty += 1
            if DEBUG_EVERY and (i % DEBUG_EVERY == 0):
                print(f"[NOFIELD] row={i} PN={pn} (fetched but did not find manufacturer/mpn)")

    print("=== FETCH DONE ===")
    print(f"parsed_ok rows: {parsed_ok}")
    print(f"no fields found rows: {parsed_empty}")
    print(f"skipped (already filled): {skipped_already_filled}")
    print(f"skipped (no PN): {skipped_no_pn}")
    print(f"cache hits: {cached_hits}")
    print(f"errors: {errors}")
    print("==================")

    # --- save
    out_xlsx = in_path.with_name(in_path.stem + "_with_mpn.xlsx")
    out_csv = in_path.with_name(in_path.stem + "_with_mpn.csv")
    df.to_excel(out_xlsx, index=False)
    df.to_csv(out_csv, index=False)

    print(f"Saved: {out_xlsx}")
    print(f"Saved: {out_csv}")
    print("Done.")


if __name__ == "__main__":
    main()
