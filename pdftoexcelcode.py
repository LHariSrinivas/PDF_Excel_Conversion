import os
import pdfplumber
import pandas as pd
from winotify import Notification, audio
import re

# Set folder paths
input_folder = "D:/Projects/SLDC Gujarat Web Scraping + Excel Conversion/downloads"
output_folder = "D:/Projects/SLDC Gujarat Web Scraping + Excel Conversion/excel_conversion"
os.makedirs(output_folder, exist_ok=True)

def extract_date_from_filename(base_name):
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

    month_map = {
        "JAN": "01", "FEB": "02", "MAR": "03", "APR": "04",
        "MAY": "05", "JUN": "06", "JUL": "07", "AUG": "08",
        "SEP": "09", "OCT": "10", "NOV": "11", "DEC": "12"
    }
    return f"01-{month_map[mon]}-{year}" if mon in month_map else ""

def clean_empty_columns(df):
    return df.loc[:, ~((df.columns == "") & (df.replace('', pd.NA).isna().all()))]

def extract_sections(pdf_path):
    wind_rows, solar_rows = [], []
    wind_header, solar_header = None, None
    current_section = None  # None, 'wind', 'solar'
    wind_done = solar_done = False

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    clean_row = [cell.strip() if cell else "" for cell in row]

                    if all(cell == "" for cell in clean_row):
                        continue

                    # Detect start of sections
                    if any("SHARE OF WIND FARM OWNER" in cell.upper() for cell in clean_row):
                        current_section = "wind"
                        continue
                    elif any("SHARE OF SOLAR GENERATOR" in cell.upper() for cell in clean_row):
                        current_section = "solar"
                        continue

                    # Stop capturing when TOTAL appears
                    if any("TOTAL" in cell.upper() for cell in clean_row):
                        if current_section == "wind":
                            wind_done = True
                        elif current_section == "solar":
                            solar_done = True
                        current_section = None
                        continue

                    # Capture data rows
                    if current_section == "wind" and not wind_done:
                        if not wind_header:
                            wind_header = clean_row
                        elif len(clean_row) == len(wind_header):
                            row_text = " ".join(clean_row).upper()
                            if "SEPC" in base_name:
                                if "CLEAN MAX" in row_text or "CLEANMAX" in row_text:
                                    wind_rows.append(clean_row)
                            else:
                                wind_rows.append(clean_row)

                    elif current_section == "solar" and not solar_done:
                        if not solar_header:
                            solar_header = clean_row
                        elif len(clean_row) == len(solar_header):
                            row_text = " ".join(clean_row).upper()
                            if "SEPC" in base_name:
                                if "CLEAN MAX" in row_text or "CLEANMAX" in row_text:
                                    solar_rows.append(clean_row)                                    
                            else:
                                solar_rows.append(clean_row)

    return wind_header, wind_rows, solar_header, solar_rows

# --- MAIN LOOP ---
for filename in os.listdir(input_folder):
    if not filename.lower().endswith(".pdf"):
        continue

    pdf_path = os.path.join(input_folder, filename)
    base_name = os.path.splitext(filename)[0]
    excel_path = os.path.join(output_folder, f"{base_name}.xlsx")

    wind_header, wind_rows, solar_header, solar_rows = extract_sections(pdf_path)
    if not wind_rows and not solar_rows:
        print(f"❌ No Wind/Solar data in: {filename}")
        continue

    date_str = extract_date_from_filename(base_name)

    df_wind = pd.DataFrame(wind_rows, columns=wind_header) if wind_header else pd.DataFrame()
    df_solar = pd.DataFrame(solar_rows, columns=solar_header) if solar_header else pd.DataFrame()

    df_wind = clean_empty_columns(df_wind)
    df_solar = clean_empty_columns(df_solar)

    if not df_wind.empty:
        df_wind.insert(1, "Date", date_str)
        df_wind.rename(columns={"SSr No": "Sr No"}, inplace=True)

        # for col in df_wind.columns:
        #     if col in "Sr No":
        #         df_wind[col] = range(1, len(df_wind)+1)

        if "Sr No" in df_wind.columns:
            df_wind["Sr No"] = range(1, len(df_wind)+1)

    if not df_solar.empty:
        df_solar.insert(1, "Date", date_str)
        df_solar.rename(columns={"SSr No": "Sr No"}, inplace=True)

        # for col in df_solar.columns:
        #     if col in "Sr No":
        #         df_solar[col] = range(1, len(df_solar)+1)
        if "Sr No" in df_solar.columns:
            df_solar["Sr No"] = range(1, len(df_solar)+1)

    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        if not df_wind.empty:
            df_wind.to_excel(writer, sheet_name="Wind Energy", index=False)
        if not df_solar.empty:
            df_solar.to_excel(writer, sheet_name="Solar Energy", index=False)

    print(f"✅ Saved Excel for → {filename}")

# ✅ Toast Notification
toast = Notification(
    app_id="SLDC Gujarat Data",
    title="PDF to Excel Conversion",
    msg="Wind & Solar section data (across pages) saved successfully.",
    duration="long"
)
toast.set_audio(audio.Default, loop=False)
toast.show()