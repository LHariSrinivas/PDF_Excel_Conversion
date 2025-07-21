import os
import pdfplumber
import pandas as pd
from winotify import Notification, audio
import re

# Set folder paths
input_folder = "D:/SLDC Gujarat Web Scraping + Excel Conversion/downloads"
output_folder = "D:/SLDC Gujarat Web Scraping + Excel Conversion/excel_conversion"
os.makedirs(output_folder, exist_ok=True)

def extract_date_from_filename(base_name):
    # Try first pattern: _YYYY_MON_
    match = re.search(r"_(\d{4})_([A-Z]{3})_", base_name)
    if match:
        year = match.group(1)
        mon = match.group(2).upper()
    else:
        # Try alternate pattern: _MON_YYYY
        match = re.search(r"_([A-Z]{3})_(\d{4})", base_name)
        if match:
            mon = match.group(1).upper()
            year = match.group(2)
        else:
            return ""

    # Month mapping
    month_map = {
        "JAN": "01", "FEB": "02", "MAR": "03", "APR": "04",
        "MAY": "05", "JUN": "06", "JUL": "07", "AUG": "08",
        "SEP": "09", "OCT": "10", "NOV": "11", "DEC": "12"
    }

    return f"01-{month_map[mon]}-{year}" if mon in month_map else ""

def extract_sections(pdf_path):
    wind_rows, solar_rows = [], []
    wind_header, solar_header = None, None
    wind_capture = solar_capture = False

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    clean_row = [cell.strip() if cell else "" for cell in row]

                    # Skip completely empty rows
                    if all(cell == "" for cell in clean_row):
                        continue

                    # Reset capture flags on section headers
                    if any("SHARE OF WIND FARM OWNER" in cell.upper() for cell in clean_row):
                        wind_capture = True
                        solar_capture = False
                        continue
                    elif any("SHARE OF SOLAR GENERATOR" in cell.upper() for cell in clean_row):
                        wind_capture = False
                        solar_capture = True
                        continue

                    # Stop capturing on TOTAL or other unrelated sections
                    if any("TOTAL" in cell.upper() or "CONSIDERATION" in cell.upper() 
                           or "PERIOD CONSIDERED" in cell.upper() or "CERTIFICATE" 
                           in cell.upper() for cell in clean_row):
                        wind_capture = False
                        solar_capture = False
                        continue

                    # Wind section logic
                    if wind_capture:
                        if not wind_header:
                            wind_header = clean_row
                        else:
                            if len(clean_row) == len(wind_header):
                                wind_rows.append(clean_row)
                        continue

                    # Solar section logic
                    if solar_capture:
                        if not solar_header:
                            solar_header = clean_row
                        else:
                            if len(clean_row) == len(solar_header):
                                solar_rows.append(clean_row)

    return wind_header, wind_rows, solar_header, solar_rows

# --- MAIN LOOP ---
for filename in os.listdir(input_folder):
    if not filename.lower().endswith(".pdf"):
        continue

    pdf_path = os.path.join(input_folder, filename)
    base_name = os.path.splitext(filename)[0]
    excel_path = os.path.join(output_folder, f"{base_name}.xlsx")

    # if os.path.exists(excel_path):
    #     print(f"⚠️ Skipping already existing: {filename}")
    #     continue

    wind_header, wind_rows, solar_header, solar_rows = extract_sections(pdf_path)
    if not wind_rows and not solar_rows:
        print(f"❌ No data found in: {filename}")
        continue

    date_str = extract_date_from_filename(base_name)

    # Create DataFrames with headers
    df_wind = pd.DataFrame(wind_rows, columns=wind_header) if wind_header else pd.DataFrame()
    df_solar = pd.DataFrame(solar_rows, columns=solar_header) if solar_header else pd.DataFrame()

    # Remove columns where both header is empty and all values are empty
    def remove_completely_empty_columns(df):
        return df.loc[:, ~((df.columns == "") & (df.isna() | df.eq("")).all())]

    df_wind = remove_completely_empty_columns(df_wind)
    df_solar = remove_completely_empty_columns(df_solar)


    # Insert Date
    if not df_wind.empty:
        df_wind.insert(1, "Date", date_str)
    if not df_solar.empty:
        df_solar.insert(1, "Date", date_str)
    
    # # Remove entirely empty columns
    # df_wind = df_wind.loc[:, ~(df_wind.replace('', pd.NA).isna().all())]
    # df_solar = df_solar.loc[:, ~(df_solar.replace('', pd.NA).isna().all())]

    # Save Excel
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        if not df_wind.empty:
            df_wind.to_excel(writer, sheet_name="Wind Energy", index=False)
        if not df_solar.empty:
            df_solar.to_excel(writer, sheet_name="Solar Energy", index=False)

    print(f"✅ Extracted and saved → {excel_path}")

# Notification
toast = Notification(
    app_id="SLDC Gujarat Data",
    title="PDF to Excel Conversion",
    msg="Only Sections B (Wind) and D (Solar) were extracted from all PDFs.",
    duration="long"
)
toast.set_audio(audio.Default, loop=False)
toast.show()
