import os
import pdfplumber
import pandas as pd
from winotify import Notification, audio
import re
from datetime import datetime

# Set folder paths
input_folder = "downloads"
output_folder = "excel_conversion"
os.makedirs(output_folder, exist_ok=True)

# Expected column headers (including empty column)
columns = [
    "Sr No", "", "Name of Wind Farm Owner", "DISCOM\nAllocation",
    "UNDER\nREC\nMechanism", "Installed\nCapacity (MW)",
    "Share in\nActive Energy\n(Mwh)", "Share in Reactive\nEnergy (Mvarh)"
]

def extract_cleanmax_from_pdf(pdf_path):
    wind_data, solar_data = [], []
    current_section = None

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    row_clean = [cell.strip() if cell else "" for cell in row]

                    # Detect section headers
                    if any("SHARE OF WIND FARM OWNER" in cell for cell in row_clean):
                        current_section = "wind"
                        continue
                    elif any("SHARE OF SOLAR GENERATOR" in cell for cell in row_clean):
                        current_section = "solar"
                        continue
                    elif any("CONSIDERATIONS FOR ISSUING" in cell for cell in row_clean):
                        current_section = None
                        continue

                    # Skip irrelevant rows
                    if all(cell == "" for cell in row_clean):
                        continue
                    if any("TOTAL" in cell.upper() for cell in row_clean if cell):
                        continue

                    # Filter CLEANMAX rows
                    if len(row_clean) >= 3 and "CLEANMAX" in row_clean[2].upper():
                        row_clean = row_clean[:8] + [""] * (8 - len(row_clean))
                        if current_section == "wind":
                            wind_data.append(row_clean)
                        elif current_section == "solar":
                            solar_data.append(row_clean)

    return wind_data, solar_data

# Loop through all PDFs in the input directory
for filename in os.listdir(input_folder):
    if filename.lower().endswith(".pdf"):
        pdf_path = os.path.join(input_folder, filename)
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        excel_path = os.path.join(output_folder, f"{base_name}.xlsx")

        # Check if Excel file already exists
        if os.path.exists(excel_path):
            print(f"⚠️  Skipping '{filename}' → Excel already exists: {excel_path}")
            continue

        # Extract data
        wind_rows, solar_rows = extract_cleanmax_from_pdf(pdf_path)

        # Skip files with no data
        if not wind_rows and not solar_rows:
            print(f"⚠️  Skipping '{filename}' → No CLEANMAX data found.")
            continue

        df_wind = pd.DataFrame(wind_rows, columns=columns)
        df_solar = pd.DataFrame(solar_rows, columns=columns)

        # Extract date from filename, e.g., "HETENERGYBHILDI-HYBRID_2025_JAN_xyz.pdf"
        match = re.search(r"_(\d{4})_([A-Z]{3})_", base_name)
        if match:
            year = match.group(1)
            month_abbr = match.group(2).upper()
            month_map = {
                "JAN": "01", "FEB": "02", "MAR": "03", "APR": "04",
                "MAY": "05", "JUN": "06", "JUL": "07", "AUG": "08",
                "SEP": "09", "OCT": "10", "NOV": "11", "DEC": "12"
            }
            if month_abbr in month_map:
                month = month_map[month_abbr]
                date_str = f"01-{month}-{year}"
            else:
                date_str = ""
        else:
            date_str = ""

        # Insert the 'Date' column after "Sr No"
        if date_str:
            df_wind.insert(1, "Date", date_str)
            df_solar.insert(1, "Date", date_str)
   
        with pd.ExcelWriter(excel_path) as writer:
            df_wind.to_excel(writer, sheet_name="Wind Energy", index=False)
            df_solar.to_excel(writer, sheet_name="Solar Energy", index=False)

        print(f"✅ Saved Excel for '{filename}' → {excel_path}")

toast = Notification(
    app_id="SLDC Gujarat Data",
    title="PDF to EXCEL Conversion",
    msg="All files converted successfully!",
    duration="long"  # Stays ~25 sec then moves to Action Center
)
toast.set_audio(audio.Default, loop=False)
toast.show()