import os
import pandas as pd
from collections import defaultdict
from winotify import Notification, audio
import re

# 📁 Input/Output paths
input_folder = "D:/Projects/SLDC Gujarat Web Scraping + Excel Conversion/excel_conversion/"
# output_folder = "D:/Projects/SLDC Gujarat Web Scraping + Excel Conversion/testing/all_combined_files/"
output_folder = "D:/ONEDRIVE 2026/OneDrive - CMES/SLDCGuj all Combined Excel"
os.makedirs(output_folder, exist_ok=True)

# 📅 Month order to sort files
MONTH_INDEX = {
    "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6,
    "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12
    }

ENERGY_NAMES = [
        "HETENERGY(BHILDI-HYBRID)",
        "66KVYASHASWA(HYBRID)",
        "SANATHAL(HEM_URJA_HYBRID)",
        "MOTA_DEVLIYA(HETENERGY_HYBRID)",
        "66KVCLEANMAXPIPARADI(HYBRID)",
        "SEPC(HYBRID)",
        "66_KV_MOTA_KHIJADIYA(SALPIPALIYA_WF)",
        "66_KV_MOTA_KHIJADIYA(SALPIPALIYA_HYBRID)",
        "DHARAGAR(GNESL)",
        "66 KV GHELDA(GNESL)",
        "220KV_NAGPUR(OP_WIND)HYBRID",
        "220_KV_NAGPUR(CLEANMAX VAYU)",
        "66KV BHOJPARA (CLEANMAX)"
    ]

# =========================================================
# 🔑 HEADER CANONICALIZATION
# Folds spacing-variant column names into ONE canonical column.
# e.g. 'Share in Reactive Energy (Mvarh)' vs 'Share inReactiveEnergy(Mvarh)'
# =========================================================

# Force specific spellings to win, regardless of file order.
# Key = column name with ALL whitespace removed + uppercased.
# Canonical display names to pin (keys are the AGGRESSIVELY-squashed forms)
FORCE_CANONICAL = {
    "SHAREINREACTIVEENERGYMVARH":      "Share in Reactive Energy (Mvarh)",
    "SHAREINACTIVEENERGYIMPORTMWH":    "Share in Active Energy (Import) (Mwh)",
    "SHAREINACTIVEENERGYMWH":          "Share in Active Energy (Mwh)",
    "NAMEOFWINDFARMOWNER":             "Name of Wind Farm Owner",
    "SOLARENTITYNAME":                 "Solar Entity Name",
    "DISCOMALLOCATION":                "DISCOM Allocation",
    "UNDERRECMECHANISM":               "UNDER REC Mechanism",
    "INSTALLEDCAPACITYMW":             "Installed Capacity (MW)",
}

def _squash(name):
    """
    Aggressively normalize a header for MATCHING:
    remove ALL whitespace, parentheses, dots, and other punctuation,
    then uppercase. This collapses cosmetic PDF-extraction variants
    (e.g. '(Import Mwh)' vs '(Import)(Mwh)', trailing '.', etc.)
    into a single key.
    """
    s = re.sub(r"[\s().,\-_]+", "", str(name))   # strip spaces & punctuation
    return s.upper()

def canonicalize_columns(df):
    """
    Merge columns that are identical except for whitespace differences.
    The FIRST-seen spelling of each squashed key is kept (or the FORCE_CANONICAL
    override, if defined), and any later variant's values are folded into it —
    filling only where the target is currently blank/NaN.
    """
    if df.empty:
        return df

    canonical = {}                       # squashed_key -> canonical column name kept
    merged = pd.DataFrame(index=df.index)

    for col in df.columns:
        key = _squash(col)

        if key not in canonical:
            # First time we see this logical column.
            # Use the forced spelling if we have one, else keep this file's spelling.
            keep_name = FORCE_CANONICAL.get(key, col)
            canonical[key] = keep_name
            merged[keep_name] = df[col]
        else:
            # A whitespace variant of a column we already have -> fold values in.
            target = canonical[key]
            target_series = merged[target]
            variant_series = df[col].reindex(merged.index)

            # Keep existing non-empty values; fill blanks/NaN from the variant.
            is_filled = target_series.notna() & (target_series.astype(str).str.strip() != "")
            merged[target] = target_series.where(is_filled, variant_series)

    return merged

# --- Grouping Logic ---
# (Site_Name)_(YYYY)_(MON).xlsx ; non-greedy (.+?) to handle underscores in site name
file_pattern = re.compile(r"(.+?)_(\d{4})_([A-Z]{3})\.xlsx", re.I)

energy_sites = defaultdict(list)
for file in os.listdir(input_folder):
    if not file.lower().endswith(".xlsx"):
        continue

    match = file_pattern.match(file)
    if match:
        site_name = match.group(1)
        year = int(match.group(2))
        month_abbr = match.group(3).upper()
        month_index = MONTH_INDEX.get(month_abbr, 99)

        sort_key = (year, month_index)
        energy_sites[site_name].append((file, sort_key))
    else:
        # Fallback for MON_YYYY pattern
        file_pattern_alt = re.compile(r"(.+?)_([A-Z]{3})_(\d{4})\.xlsx", re.I)
        match_alt = file_pattern_alt.match(file)
        if match_alt:
            site_name = match_alt.group(1)
            year = int(match_alt.group(3))
            month_abbr = match_alt.group(2).upper()
            month_index = MONTH_INDEX.get(month_abbr, 99)

            sort_key = (year, month_index)
            energy_sites[site_name].append((file, sort_key))
        else:
            print(f"⚠️ File '{file}' did not match pattern, skipping.")

# 🔁 Merge files for each energy site
for site_name, file_data_list in energy_sites.items():
    print(f"\n🔧 Merging for site: {site_name}")

    # Sort chronologically by (year, month_index) — ensures JAN (old spelling) is seen first
    files_sorted_tuples = sorted(file_data_list, key=lambda item: item[1])

    wind_data_all = []
    solar_data_all = []

    for file_tuple in files_sorted_tuples:
        file = file_tuple[0]
        path = os.path.join(input_folder, file)
        print(f"   📄 Reading: {file}")

        try:
            excel_files = pd.ExcelFile(path, engine="openpyxl")
            wind_df = pd.read_excel(path, sheet_name="Wind Energy", dtype=str) if "Wind Energy" in excel_files.sheet_names else pd.DataFrame()
            solar_df = pd.read_excel(path, sheet_name="Solar Energy", dtype=str) if "Solar Energy" in excel_files.sheet_names else pd.DataFrame()

            # 🛡️ Ensure Date column exists
            if not wind_df.empty and "Date" not in wind_df.columns:
                print(f"   ⚠️ Skipped wind — 'Date' missing in {file}")
                wind_df = pd.DataFrame()

            if not solar_df.empty and "Date" not in solar_df.columns:
                print(f"   ⚠️ Skipped solar — 'Date' missing in {file}")
                solar_df = pd.DataFrame()

            # 🔑 Canonicalize headers per-file BEFORE collecting
            wind_df = canonicalize_columns(wind_df)
            solar_df = canonicalize_columns(solar_df)

            wind_data_all.append(wind_df)
            solar_data_all.append(solar_df)

        except Exception as e:
            print(f"   ❌ Error in {file}: {e}")

    combined_path = os.path.join(output_folder, f"{site_name}_combined.xlsx")
    with pd.ExcelWriter(combined_path, engine="openpyxl") as writer:
        # ---------- WIND ----------
        if any(not df.empty for df in wind_data_all):
            wind_merged = pd.concat([df for df in wind_data_all if not df.empty], ignore_index=True)
            # 🔑 Canonicalize ONCE MORE on the combined frame, in case variants
            # appeared across different files.
            wind_merged = canonicalize_columns(wind_merged)
            wind_merged["Sr No"] = range(1, len(wind_merged) + 1)
            wind_merged["Energy Name"] = site_name
            print(f"   Wind columns: {list(wind_merged.columns)}")
            wind_merged.to_excel(writer, sheet_name="Wind Energy", index=False)
        else:
            pd.DataFrame(columns=["Sr No", "Date"]).to_excel(writer, sheet_name="Wind Energy", index=False)

        # ---------- SOLAR ----------
        if any(not df.empty for df in solar_data_all):
            solar_merged = pd.concat([df for df in solar_data_all if not df.empty], ignore_index=True)
            solar_merged = canonicalize_columns(solar_merged)
            solar_merged["Sr No"] = range(1, len(solar_merged) + 1)
            solar_merged["Energy Name"] = site_name
            print(f"   Solar columns: {list(solar_merged.columns)}")
            solar_merged.to_excel(writer, sheet_name="Solar Energy", index=False)
        else:
            pd.DataFrame(columns=["Sr No", "Date"]).to_excel(writer, sheet_name="Solar Energy", index=False)

    print(f"✅ Combined Excel created: {combined_path}")

toast = Notification(
    app_id="SLDC Gujarat Data",
    title="Excel Merging",
    msg="All Excel Files have been Merged Successfully.",
    duration="long"
)
toast.set_audio(audio.Default, loop=False)
toast.show()