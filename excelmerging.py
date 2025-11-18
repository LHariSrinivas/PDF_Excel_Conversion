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

# --- DELETED get_month_index ---
# --- DELETED extract_site_name ---

# --- NEW Grouping Logic ---
# This regex captures (Site_Name)_(YYYY)_(MON).xlsx
# It's non-greedy (.+?) to handle underscores in the site name
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
        month_index = MONTH_INDEX.get(month_abbr, 99) # Get 1-12 index
        
        # Store the file and its sort key (year, month_index)
        sort_key = (year, month_index)
        energy_sites[site_name].append( (file, sort_key) )
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
            energy_sites[site_name].append( (file, sort_key) )
        else:
            print(f"⚠️ File '{file}' did not match pattern, skipping.")

# 🔁 Merge files for each energy site 
for site_name, file_data_list in energy_sites.items():
    print(f"\n🔧 Merging for site: {site_name}")
    
    # Sort the list based on the tuple (year, month_index)
    # This is the fix for Bug #1
    files_sorted_tuples = sorted(file_data_list, key=lambda item: item[1])

    wind_data_all = []
    solar_data_all = []

    # Loop through the sorted tuples
    for file_tuple in files_sorted_tuples:
        file = file_tuple[0] # Get the filename from the tuple
        path = os.path.join(input_folder, file)
        print(f"   📄 Reading: {file}")
        
        try:
            excel_files = pd.ExcelFile(path, engine="openpyxl")
            # Use dtype=str to prevent pandas from breaking data
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