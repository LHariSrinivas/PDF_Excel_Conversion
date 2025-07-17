import os
import pandas as pd

input_folder = "excel_conversion"
output_file = "downloads/cleanmax_final_combined.xlsx"

# Month ordering to preserve natural month-wise merging
MONTH_INDEX = {
    "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6,
    "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12
}

# Expected columns
columns = [
    "Sr No", "Date", "", "Name of Wind Farm Owner", "DISCOM\nAllocation",
    "UNDER\nREC\nMechanism", "Installed\nCapacity (MW)",
    "Share in\nActive Energy\n(Mwh)", "Share in Reactive\nEnergy (Mvarh)"
]

def get_month_index(filename):
    for month_abbr, idx in MONTH_INDEX.items():
        if month_abbr in filename.upper():
            return idx
    return 999  # If no match, push to end

# Gather all Excel files and sort by month index
files = [
    f for f in os.listdir(input_folder) if f.lower().endswith(".xlsx")
]
files = sorted(files, key=get_month_index)

# Merged data containers
all_wind_data = []
all_solar_data = []

# Process each file
for filename in files:
    filepath = os.path.join(input_folder, filename)
    print(f"📂 Adding: {filename}")

    try:
        wind_df = pd.read_excel(filepath, sheet_name="Wind Energy", dtype=str, engine="openpyxl")
        solar_df = pd.read_excel(filepath, sheet_name="Solar Energy", dtype=str, engine="openpyxl")

        if "Date" not in wind_df.columns or "Date" not in solar_df.columns:
            print(f"⚠️ Skipped {filename} — Missing 'Date' column")
            continue

        # Preserve original order and structure
        wind_df = wind_df[[col for col in columns if col in wind_df.columns]]
        solar_df = solar_df[[col for col in columns if col in solar_df.columns]]

        all_wind_data.append(wind_df)
        all_solar_data.append(solar_df)

    except Exception as e:
        print(f"❌ Error with '{filename}' → {type(e).__name__}: {e}")

# Export merged file
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    if all_wind_data:
        wind_merged = pd.concat(all_wind_data, ignore_index=True)
        wind_merged["Sr No"] = range(1, len(wind_merged) + 1)
        wind_merged = wind_merged.astype(str)
        wind_merged.to_excel(writer, sheet_name="Wind Energy", index=False)

    if all_solar_data:
        solar_merged = pd.concat(all_solar_data, ignore_index=True)
        solar_merged["Sr No"] = range(1, len(solar_merged) + 1)
        solar_merged = solar_merged.astype(str)
        solar_merged.to_excel(writer, sheet_name="Solar Energy", index=False)

print(f"\n✅ Merged file created in monthly order: {output_file}")