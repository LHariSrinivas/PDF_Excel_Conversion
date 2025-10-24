"""
Patched PDF -> Excel extractor for SLDC Gujarat PDFs
- Stops at 'TOTAL' for each section.
- Skips rows containing unwanted phrases.
- Fixes page-to-page header mismatch for Wind section.
- Produces compact Wind sheet with columns:
  ['Sr No','Owner','DISCOM','Under REC','Installed Capacity (MW)','Share Active (MWh)','Share Reactive (Mvarh)']
"""

import os
import re
import pdfplumber
import pandas as pd
from winotify import Notification, audio

# ---------- USER CONFIG ----------
INPUT_FOLDER = r"D:/Projects/SLDC Gujarat Web Scraping + Excel Conversion/downloads"
OUTPUT_FOLDER = r"D:/Projects/SLDC Gujarat Web Scraping + Excel Conversion/excel_conversion"
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

DEBUG = False  # Set True to print debug traces (alignment decisions etc.)

# Unwanted phrases to ignore anywhere in a row (case-insensitive)
UNWANTED_PHRASES = [
    "Period Considered for the month",
    "Active Energy Received From",
    "Reactive Energy Supplied to",
    "GUJARAT ENERGY TRANSMISSION CORPORATION LIMITED"
]
# Lowercase for quick checks
UNWANTED_LOWER = [p.upper() for p in UNWANTED_PHRASES]
# ---------- END CONFIG ----------

# ---------- Utilities ----------
def extract_date_from_filename(base_name):
    m = re.search(r"_(\d{4})_([A-Za-z]{3})", base_name)
    if m:
        year, mon = m.group(1), m.group(2).upper()
    else:
        m2 = re.search(r"_([A-Za-z]{3})_(\d{4})", base_name)
        if m2:
            mon, year = m2.group(1).upper(), m2.group(2)
        else:
            return ""
    month_map = {"JAN":"01","FEB":"02","MAR":"03","APR":"04","MAY":"05","JUN":"06",
                 "JUL":"07","AUG":"08","SEP":"09","OCT":"10","NOV":"11","DEC":"12"}
    return f"01-{month_map.get(mon,'01')}-{year}"

def clean_empty_columns(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    return df.loc[:, ~((df.columns == "") & (df.replace('', pd.NA).isna().all()))]

def row_contains_unwanted(joined_upper: str) -> bool:
    """Return True if the row contains any unwanted phrase (case-insensitive)."""
    for p in UNWANTED_LOWER:
        if p in joined_upper:
            return True
    return False

# ---------- Canonical header discovery ----------
def find_canonical_wind_header_from_tables(pdf) -> tuple[list, int]:
    for page_num, page in enumerate(pdf.pages, start=1):
        for table in page.extract_tables():
            if not table:
                continue
            for row in table:
                if not row:
                    continue
                joined = " ".join([str(x).upper() for x in row if x])
                if "SR NO" in joined or "NAME OF WIND" in joined or "SHARE OF WIND" in joined:
                    return [cell if cell else "" for cell in row], page_num
    return None, None

def find_canonical_wind_header_from_text(pdf) -> tuple[list, int, str]:
    for page_num, page in enumerate(pdf.pages, start=1):
        text = page.extract_text() or ""
        for line in text.splitlines():
            UL = line.upper()
            if "SR NO" in UL or "NAME OF WIND" in UL or "NAME OF WIND FARM OWNER" in UL:
                parts = re.split(r"\s{2,}", line.strip())
                parts = [p.strip() for p in parts if p.strip()]
                return parts, page_num, line
    return None, None, None

# ---------- Compact-row logic (no fixed DISCOM list) ----------
def normalize_token(tok):
    if tok is None:
        return ""
    return str(tok).strip()

def looks_numeric(s: str) -> bool:
    if not s:
        return False
    s2 = str(s).replace(",", "").strip()
    return bool(re.fullmatch(r"[+-]?\d+(\.\d+)?", s2))

def looks_discom(s: str) -> bool:
    if not s:
        return False
    t = re.sub(r"\s+","", str(s)).upper()
    return bool(re.fullmatch(r"[A-Z]{2,6}", t)) or "VCL" in t

def looks_owner_name(s: str) -> bool:
    if not s:
        return False
    t = str(s).strip()
    has_letters = bool(re.search(r"[A-Za-z]", t))
    is_short_acro = bool(re.fullmatch(r"[A-Z0-9]{1,6}", re.sub(r"\s+","",t)))
    return has_letters and not is_short_acro and len(t) > 6

def compact_row(row):
    tokens = [normalize_token(c) for c in row if c is not None]
    tokens_nonempty = [t for t in tokens if t != ""]

    sr = owner = discom = under_rec = installed = active = reactive = ""

    # Sr No
    for t in tokens:
        if re.fullmatch(r"\d{1,4}", str(t).strip()):
            sr = t
            break
    if sr == "" and tokens_nonempty:
        sr = tokens_nonempty[0]

    # Owner
    owner_candidates = []
    for t in tokens:
        if t == "" or t == sr:
            continue
        if "(" in str(t) and ")" in str(t):
            owner_candidates.append(t)
        elif looks_owner_name(t):
            owner_candidates.append(t)
    if owner_candidates:
        owner = owner_candidates[0]
    else:
        for t in tokens:
            if t == "" or t == sr:
                continue
            if not looks_discom(t) and not looks_numeric(t):
                owner = t
                break

    # DISCOM
    for t in tokens:
        if t == "" or t == sr or t == owner:
            continue
        if looks_discom(t):
            discom = t
            break

    # Under REC
    for t in tokens:
        if t and ("REC" in str(t).upper() or "UNDER" in str(t).upper()):
            under_rec = t
            break

    # Numerics
    numerics = []
    for t in tokens:
        if t == "" or t == sr:
            continue
        if looks_numeric(t):
            numerics.append(t)
    if len(numerics) >= 1:
        installed = numerics[0]
    if len(numerics) >= 2:
        active = numerics[1]
    if len(numerics) >= 3:
        reactive = numerics[2]

    return [sr, owner, discom, under_rec, installed, active, reactive]

# ---------- Core extraction flow (with TOTAL stopping and unwanted skipping) ----------
def extract_sections(pdf_path, base_name=""):
    wind_compacted = []
    solar_raw = []

    with pdfplumber.open(pdf_path) as pdf:
        # canonical header (optional use)
        can_header, can_page = find_canonical_wind_header_from_tables(pdf)
        if not can_header:
            text_header, text_page, rawline = find_canonical_wind_header_from_text(pdf)
            if text_header:
                can_header = text_header

        if DEBUG:
            print("Canonical wind header found:", bool(can_header), "on page", can_page or text_page if not can_header else can_page)

        current_section = None

        # Walk pages & tables
        for page_num, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables()
            if DEBUG:
                print(f"Page {page_num} tables: {len(tables)}")
            for table in tables:
                if not table:
                    continue
                for row in table:
                    if row is None:
                        continue
                    clean_row = [cell if cell else "" for cell in row]
                    if all(cell == "" for cell in clean_row):
                        continue
                    joined_upper = " ".join([str(x).upper() for x in clean_row])

                    # If row contains any unwanted phrase, skip it immediately
                    if row_contains_unwanted(joined_upper):
                        if DEBUG:
                            print(f"Skipping unwanted row (page {page_num}): {joined_upper[:80]}...")
                        continue

                    # Section detection
                    if "SHARE OF WIND" in joined_upper or "NAME OF WIND FARM OWNER" in joined_upper or "SHARE OF WIND FARM OWNER" in joined_upper:
                        current_section = "wind"
                        if DEBUG:
                            print(f"Detected WIND section on page {page_num}")
                        continue
                    if "SHARE OF SOLAR" in joined_upper or "NAME OF SOLAR" in joined_upper or "SHARE OF SOLAR GENERATOR" in joined_upper:
                        current_section = "solar"
                        if DEBUG:
                            print(f"Detected SOLAR section on page {page_num}")
                        continue

                    # Stop when TOTAL encountered: stop collecting for this section (per your request)
                    if "TOTAL" in joined_upper:
                        if DEBUG:
                            print(f"Encountered TOTAL on page {page_num}. Stopping current section capture.")
                        current_section = None
                        continue

                    # Capture rows
                    if current_section == "wind":
                        comp = compact_row(clean_row)
                        # optional SEPC filter kept compatible (use base_name)
                        if "SEPC" in base_name.upper():
                            top = " ".join([str(x).upper() for x in comp])
                            if "CLEAN MAX" in top or "CLEANMAX" in top:
                                wind_compacted.append({"page": page_num, "raw": clean_row, "compact": comp})
                        else:
                            wind_compacted.append({"page": page_num, "raw": clean_row, "compact": comp})

                    elif current_section == "solar":
                        # skip rows flagged unwanted already above
                        solar_raw.append({"page": page_num, "raw": clean_row})

    return can_header, wind_compacted, solar_raw

# ---------- Main driver ----------
def main():
    for filename in os.listdir(INPUT_FOLDER):
        if not filename.lower().endswith(".pdf"):
            continue
        pdf_path = os.path.join(INPUT_FOLDER, filename)
        base_name = os.path.splitext(filename)[0]
        excel_path = os.path.join(OUTPUT_FOLDER, f"{base_name}.xlsx")

        print("Processing:", filename)
        canonical_header, wind_compacted, solar_raw = extract_sections(pdf_path, base_name=base_name)

        if not wind_compacted and not solar_raw:
            print(f"❌ No Wind/Solar data found in: {filename}")
            continue

        # Build wind DataFrame
        wind_rows = [r["compact"] for r in wind_compacted]
        if wind_rows:
            df_wind = pd.DataFrame(wind_rows, columns=[
                "Sr No", "Owner", "DISCOM", "Under REC",
                "Installed Capacity (MW)", "Share Active (MWh)", "Share Reactive (Mvarh)"
            ])
            date_str = extract_date_from_filename(base_name)
            df_wind.insert(1, "Date", date_str)

            # Filter: keep rows that look like real data (Sr No numeric + at least one numeric column)
            def valid_sr(x):
                return bool(re.fullmatch(r"\d{1,3}", str(x).strip()))
            def has_numeric_row(r):
                return any(re.fullmatch(r"[+-]?\d+(\.\d+)?", str(v).replace(",","").strip()) for v in [r["Installed Capacity (MW)"], r["Share Active (MWh)"], r["Share Reactive (Mvarh)"]])
            df_wind = df_wind[df_wind["Sr No"].apply(valid_sr) & df_wind.apply(has_numeric_row, axis=1)].reset_index(drop=True)
        else:
            df_wind = pd.DataFrame()

        # Build solar DataFrame (raw rows preserved), pad to uniform width and insert date
        if solar_raw:
            maxlen = max(len(r["raw"]) for r in solar_raw)
            solar_rows_padded = [r["raw"] + [""]*(maxlen - len(r["raw"])) for r in solar_raw]
            if canonical_header and len(canonical_header) == maxlen:
                colnames_solar = canonical_header
            else:
                colnames_solar = [f"COL_{i}" for i in range(maxlen)]
            df_solar = pd.DataFrame(solar_rows_padded, columns=colnames_solar)
            date_str = extract_date_from_filename(base_name)
            df_solar.insert(1, "Date", date_str)
        else:
            df_solar = pd.DataFrame()

        # Clean and write Excel
        df_wind = clean_empty_columns(df_wind)
        df_solar = clean_empty_columns(df_solar)

        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            if not df_wind.empty:
                df_wind.to_excel(writer, sheet_name="Wind Energy", index=False)
            if not df_solar.empty:
                df_solar.to_excel(writer, sheet_name="Solar Energy", index=False)

        print(f"✅ Saved Excel for → {filename} -> {excel_path}")

    # Toast (Windows)
    try:
        toast = Notification(
            app_id="SLDC Gujarat Data Extraction",
            title="PDF to Excel Conversion Complete",
            msg="Wind & Solar data successfully extracted and saved to Excel.",
            duration="short"
        )
        toast.set_audio(audio.Default, loop=False)
        toast.show()
    except Exception:
        pass

if __name__ == "__main__":
    main()


