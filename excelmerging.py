import os
import pandas as pd
from collections import defaultdict

# 📁 Input/Output paths
input_folder = "D:/SLDC Gujarat Web Scraping + Excel Conversion/excel_conversion"
output_folder = "all_combined_excel_files"
os.makedirs(output_folder, exist_ok=True)

# 📅 Month order to sort files
MONTH_INDEX = {
    "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6,
    "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12
}

# 📦 Get month index from filename
def get_month_index(filename):
    for month_abbr, idx in MONTH_INDEX.items():
        if month_abbr in filename.upper():
            return idx
    return 999

# 📁 Group files by energy site
energy_sites = defaultdict(list)
for file in os.listdir(input_folder):
    if file.lower().endswith(".xlsx"):
        energy_name = file.split("_")[0]
        energy_sites[energy_name].append(file)

# 🔁 Merge files for each energy site
for energy_name, files in energy_sites.items():
    print(f"\n🔧 Merging for site: {energy_name}")
    files_sorted = sorted(files, key=get_month_index)

    wind_data_all = []
    solar_data_all = []

    for file in files_sorted:
        path = os.path.join(input_folder, file)
        print(f"   📄 Reading: {file}")
        try:
            wind_df = pd.read_excel(path, sheet_name="Wind Energy", dtype=str, engine="openpyxl")
            solar_df = pd.read_excel(path, sheet_name="Solar Energy", dtype=str, engine="openpyxl")

            # 🛡️ Ensure Date column exists
            if "Date" not in wind_df.columns or "Date" not in solar_df.columns:
                print(f"   ⚠️ Skipped — 'Date' column missing in {file}")
                continue

            # # 🧹 Drop fully empty columns (including headers!)
            # wind_df = wind_df.dropna(axis=1, how='all')
            # solar_df = solar_df.dropna(axis=1, how='all')

            wind_data_all.append(wind_df)
            solar_data_all.append(solar_df)

        except Exception as e:
            print(f"   ❌ Error in {file}: {e}")

    if wind_data_all or solar_data_all:
        combined_path = os.path.join(output_folder, f"{energy_name}_combined.xlsx")
        with pd.ExcelWriter(combined_path, engine="openpyxl") as writer:
            if wind_data_all:
                wind_merged = pd.concat(wind_data_all, ignore_index=True)
                wind_merged["Sr No"] = range(1, len(wind_merged) + 1)
                wind_merged.to_excel(writer, sheet_name="Wind Energy", index=False)

            if solar_data_all:
                solar_merged = pd.concat(solar_data_all, ignore_index=True)
                solar_merged["Sr No"] = range(1, len(solar_merged) + 1)
                solar_merged.to_excel(writer, sheet_name="Solar Energy", index=False)

        print(f"✅ Combined Excel created: {combined_path}")
    else:
        print(f"⚠️ No valid data found for {energy_name}")