import os
import pandas as pd
from collections import defaultdict
from winotify import Notification, audio
import re

# 📁 Input/Output paths
input_folder = "D:/Projects/SLDC Gujarat Web Scraping + Excel Conversion/excel_conversion"
output_folder = "D:/Projects/SLDC Gujarat Web Scraping + Excel Conversion/all_combined_excel_files"
os.makedirs(output_folder, exist_ok=True)

# 📅 Month order to sort files
MONTH_INDEX = {
    "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6,
    "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12
    }

def get_month_index(filename):
    # Find first month in the filename
    for month_abbr, idx in MONTH_INDEX.items():
        if month_abbr in filename.upper():
            return idx
    return 999

# Extract site name dynamically
def extract_site_name(filename):
    filename = filename.replace(".xlsx", "")
    parts = filename.split("_")
    clean_parts = []
    for part in parts:
        if part.upper() in MONTH_INDEX:  # stop at month
            break
        if re.fullmatch(r"[a-fA-F0-9]{1,12}", part):  # skip random hash
            continue
        clean_parts.append(part)
    return "_".join(clean_parts)

# Group files by site name
energy_sites = defaultdict(list)
for file in os.listdir(input_folder):
    if file.lower().endswith(".xlsx"):
        site_name = extract_site_name(file)
        energy_sites[site_name].append(file)

# 🔁 Merge files for each energy site 
for site_name, files in energy_sites.items():
    print(f"\n🔧 Merging for site: {site_name}")
    files_sorted = sorted(files, key=get_month_index)

    wind_data_all = []
    solar_data_all = []

    for file in files_sorted:
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

            wind_data_all.append(wind_df)
            solar_data_all.append(solar_df)

        except Exception as e:
            print(f"   ❌ Error in {file}: {e}")

    combined_path = os.path.join(output_folder, f"{site_name}_combined.xlsx")
    with pd.ExcelWriter(combined_path, engine="openpyxl") as writer:
        if any(not df.empty for df in wind_data_all):
            wind_merged = pd.concat([df for df in wind_data_all if not df.empty], ignore_index=True)
            wind_merged["Sr No"] = range(1, len(wind_merged) + 1)
            wind_merged.to_excel(writer, sheet_name="Wind Energy", index=False)
        else:
            pd.DataFrame(columns=["Sr No", "Date"]).to_excel(writer, sheet_name="Wind Energy", index=False)

        if any(not df.empty for df in solar_data_all):
            solar_merged = pd.concat([df for df in solar_data_all if not df.empty], ignore_index=True)
            solar_merged["Sr No"] = range(1, len(solar_merged) + 1)
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