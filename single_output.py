import os
import pandas as pd
from collections import defaultdict
from winotify import Notification, audio
import re

# 📁 Input/Output paths
input_folder = "D:/Projects/SLDC Gujarat Web Scraping + Excel Conversion/excel_conversion/"
output_folder = "D:/Projects/SLDC Gujarat Web Scraping + Excel Conversion/all_combined_excel_files"
os.makedirs(output_folder, exist_ok=True)

# 📅 Month order to sort files
MONTH_INDEX = {
    "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6,
    "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12
}

file_pattern     = re.compile(r"(.+?)_(\d{4})_([A-Z]{3})\.xlsx", re.I)
file_pattern_alt = re.compile(r"(.+?)_([A-Z]{3})_(\d{4})\.xlsx", re.I)

energy_sites = defaultdict(list)

for file in os.listdir(input_folder):
    if not file.lower().endswith(".xlsx"):
        continue

    match = file_pattern.match(file)
    if match:
        site_name  = match.group(1)
        year       = int(match.group(2))
        month_abbr = match.group(3).upper()
    else:
        match = file_pattern_alt.match(file)
        if match:
            site_name  = match.group(1)
            year       = int(match.group(3))
            month_abbr = match.group(2).upper()
        else:
            print(f"⚠️ File '{file}' did not match any pattern, skipping.")
            continue

    month_index = MONTH_INDEX.get(month_abbr, 99)
    energy_sites[site_name].append((file, (year, month_index)))

# 🗂️ Master lists to collect data across ALL sites
all_wind_dfs  = []
all_solar_dfs = []

# 🔁 Merge files for each energy site
for site_name, file_data_list in energy_sites.items():
    print(f"\n🔧 Merging for site: {site_name}")

    files_sorted = sorted(file_data_list, key=lambda item: item[1])

    wind_data_all  = []
    solar_data_all = []

    for file, _ in files_sorted:
        path = os.path.join(input_folder, file)
        print(f"   📄 Reading: {file}")

        try:
            xls      = pd.ExcelFile(path, engine="openpyxl")
            wind_df  = pd.read_excel(xls, sheet_name="Wind Energy",  dtype=str) if "Wind Energy"  in xls.sheet_names else pd.DataFrame()
            solar_df = pd.read_excel(xls, sheet_name="Solar Energy", dtype=str) if "Solar Energy" in xls.sheet_names else pd.DataFrame()

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

    # ✅ Write per-site combined file (unchanged, for your team)
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

    print(f"✅ Per-site file created: {combined_path}")

    if any(not df.empty for df in wind_data_all):
        print(f"   🗂️ Wind columns for {site_name}: {wind_merged.columns.tolist()}")
    if any(not df.empty for df in solar_data_all):
        print(f"   🗂️ Solar columns for {site_name}: {solar_merged.columns.tolist()}")

    def normalize_columns(df):
        df = df.copy()
        df.columns = [
            re.sub(r'\s+', ' ', col)   # collapse all whitespace/newlines to single space
               .strip()                 # strip leading/trailing spaces
               .rstrip('.')             # remove trailing dots
            for col in df.columns
        ]
        return df    

    # 🗂️ Also feed into master lists with site tag
    for df in wind_data_all:
        if not df.empty:
            df = normalize_columns(df)
            df.insert(1, "Site Name", site_name)
            all_wind_dfs.append(df)

    for df in solar_data_all:
        if not df.empty:
            df = normalize_columns(df)
            df.insert(1, "Site Name", site_name)
            all_solar_dfs.append(df)

# 💾 Write the two master output files
output_map = {
    "Wind Energy":  (all_wind_dfs,  "All_Sites_Wind_Energy.xlsx"),
    "Solar Energy": (all_solar_dfs, "All_Sites_Solar_Energy.xlsx"),
}

for sheet_label, (dfs, filename) in output_map.items():
    out_path = os.path.join(output_folder, filename)

    if dfs:
        # Use column order from first valid df as the master template
        master_columns = dfs[0].columns.tolist()

        # Reindex all dfs to match master column order, filling missing with NaN
        aligned_dfs = [df.reindex(columns=master_columns) for df in dfs]

        merged = pd.concat(aligned_dfs, ignore_index=True)
        merged["Sr No"] = range(1, len(merged) + 1)
        merged.to_excel(out_path, sheet_name=sheet_label, index=False)
        print(f"\n✅ Master {sheet_label} → {out_path}  ({len(merged)} rows)")
    else:
        pd.DataFrame(columns=["Sr No", "Site Name", "Date"]).to_excel(
            out_path, sheet_name=sheet_label, index=False
        )
        print(f"\n⚠️ No {sheet_label} data found. Empty file written.")

# 🔔 Toast notification
toast = Notification(
    app_id="SLDC Gujarat Data",
    title="Excel Merging Complete",
    msg="Per-site files + Wind & Solar master files created successfully.",
    duration="long"
)
toast.set_audio(audio.Default, loop=False)
toast.show()