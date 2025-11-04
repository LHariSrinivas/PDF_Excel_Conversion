import os
import pdfplumber
import pandas as pd
from winotify import Notification, audio
import re

# Set folder paths
input_folder = "D:/Projects/SLDC Gujarat Web Scraping + Excel Conversion/downloads"
output_folder = "D:/Projects/SLDC Gujarat Web Scraping + Excel Conversion/sample_testing/pdftoexcel"
os.makedirs(output_folder, exist_ok=True) # Ensure output folder exists

# --- UNWANTED_TEXT list for filtering ---
UNWANTED_TEXT = ["Period Considered for the month", 
                 "Active Energy Received From",
                 "Reactive Energy Supplied to", 
                 "GUJARAT ENERGY TRANSMISSION CORPORATION LIMITED"]

# --- FIXED extract_date_from_filename ---
def extract_date_from_filename(base_name):
    # This regex is more general to find YEAR_MON or MON_YEAR
    match = re.search(r"(\d{4})_([A-Z]{3})", base_name.upper()) # No leading '_'
    if match:
        year = match.group(1)
        mon = match.group(2).upper()
    else:
        # Try alternate pattern: _MON_YYYY
        match = re.search(r"([A-Z]{3})_(\d{4})", base_name.upper())
        if match:
            mon = match.group(1).upper()
            year = match.group(2)
        else:
            print(f"DEBUG WARNING: No date found in filename: {base_name}")
            return "" # No date found

    month_map = {
        "JAN": "01", "FEB": "02", "MAR": "03", "APR": "04",
        "MAY": "05", "JUN": "06", "JUL": "07", "AUG": "08",
        "SEP": "09", "OCT": "10", "NOV": "11", "DEC": "12"
    }
    return f"01-{month_map[mon]}-{year}" if mon in month_map else ""

# --- FIXED clean_empty_columns ---
def clean_empty_columns(df):
    """
    Removes columns from a DataFrame that have a blank header (None or '')
    AND contain no data (all values are None, NA, or '').
    """
    if df.empty:
        return df

    # 1. Find columns where the header name is blank ('' or None)
    is_blank_col = (df.columns.astype(str).str.strip() == '') | (df.columns.isna())
    
    # 2. Find columns where all values are NA or blank strings
    # .replace handles both '' and None, .isna() catches all
    is_all_na_col = df.replace('', pd.NA).isna().all()
    
    # 3. We want to drop columns that are BOTH blank AND all-NA
    is_to_drop = is_blank_col & is_all_na_col
    
    # 4. Select columns that are NOT to be dropped
    return df.loc[:, ~is_to_drop]

def align_to_header(row, full_header):
    """
    (This is your provided function, unchanged)
    """
    header_len = len(full_header)

    data_values = [cell for cell in row]

    if len(row) == header_len:
        return row

    header_nonempty_indices = [i for i, h in enumerate(full_header) if h and h.strip()]
    nonempty_count = len(header_nonempty_indices)

    if len(data_values) == nonempty_count:
        aligned = [""] * header_len
        di = 0
        for idx in header_nonempty_indices:
            aligned[idx] = data_values[di]
            di += 1
        return aligned

    aligned_row = [""] * header_len
    data_index = 0
    for i in range(header_len):
        if full_header[i] and full_header[i].strip():
            if data_index < len(data_values):
                aligned_row[i] = data_values[data_index]
                data_index += 1
            else:
                aligned_row[i] = ""
        else:
            aligned_row[i] = ""

    if data_index < len(data_values) and len(data_values) > 0:
        print(f"DEBUG WARNING: Row has unmapped data: {data_values[data_index:]}. Returning None.")
        return None

    return aligned_row

def extract_sections(pdf_path):
    wind_rows, solar_rows = [], []
    wind_header, solar_header = None, None
    current_section = None
    total_count = 0
                     
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]

    # Helper to clean just for checking, since row data is now raw
    def clean_for_check(cell):
        if isinstance(cell, str):
            return cell.strip().upper()
        return ""

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            print(f"--- Processing page {page_num + 1} ---")

            for table in tables:
                if not table: continue
                
                for row in table:
                    if not row: continue 
                    
                    # FIX: Use un-stripped row data, as per your note
                    clean_row = [cell if cell else "" for cell in row]
                    
                    if all(cell == "" for cell in clean_row):
                        continue

                    # Create a stripped/upper list just for CHECKS
                    check_row_text = " ".join([clean_for_check(c) for c in clean_row if c])

                    # Section detection
                    if "SHARE OF WIND FARM OWNER" in check_row_text:
                        current_section = "wind"
                        print(f"DEBUG: Found wind section header on page {page_num + 1}.")
                        continue
                    elif "SHARE OF SOLAR GENERATOR" in check_row_text:
                        current_section = "solar"
                        print(f"DEBUG: Found solar section header on page {page_num + 1}.")
                        continue

                    # Stop when TOTAL encountered
                    if "TOTAL" in check_row_text and current_section is not None:
                        total_count += 1
                        print(f"DEBUG: Found 'TOTAL' row (Count: {total_count}) on page {page_num + 1}.")
                        current_section = None
                        continue

                    # Capture headers
                    if current_section == "wind" and not wind_header and "SR NO" in check_row_text and "WIND FARM OWNER" in check_row_text:
                        wind_header = [cell if cell else "" for cell in row]
                        print(f"DEBUG: Captured wind header. Length: {len(wind_header)}")
                        continue
                    elif current_section == "solar" and not solar_header and "SOLAR ENTITY NAME" in check_row_text:
                        solar_header = [cell if cell else "" for cell in row]
                        print(f"DEBUG: Captured solar header. Length: {len(solar_header)}")
                        continue
                    
                    # --- Unwanted text check: This is DELIBERATELY SKIPPED here ---
                    # We will filter the final DataFrame, which is safer.
                    
                    # --- Data Capture Logic (with Page 2 Wind Fix) ---
                    norm_row = None
                    first_cell_val = clean_for_check(clean_row[0])

                    if current_section == "wind" and wind_header:
                        # --- Special handling for Page 2+ Wind Continuation ---
                        if page_num > 0 and first_cell_val.isdigit() and len(clean_row) < len(wind_header):
                            print(f"DEBUG: Applying Page 2+ wind row logic for Sr No '{first_cell_val}'")
                            try:
                                # Manually extract the 7 data values from their known positions
                                data_list = [
                                    clean_row[0], # Sr No (at index 0)
                                    clean_row[2], # Name (at index 2)
                                    clean_row[3], # DISCOM (at index 3)
                                    clean_row[4], # REC (at index 4)
                                    clean_row[5], # Capacity (at index 5)
                                    clean_row[6], # Active (at index 6)
                                    clean_row[7]  # Reactive (at index 7)
                                ]
                                
                                norm_row = [""] * len(wind_header)
                                header_nonempty_indices = [i for i, h in enumerate(wind_header) if h and h.strip()]
                                
                                if len(data_list) == len(header_nonempty_indices):
                                    for idx, val in zip(header_nonempty_indices, data_list):
                                        norm_row[idx] = val
                                else:
                                    print(f"DEBUG WARNING: Page 2 data map mismatch. Data({len(data_list)}) vs Header({len(header_nonempty_indices)})")
                                    norm_row = None
                                    
                            except IndexError:
                                print(f"DEBUG WARNING: Page 2 wind row IndexError. Row: {clean_row}")
                                norm_row = None
                        
                        # --- ELSE: Use the original align_to_header for Page 1 and other data ---
                        elif first_cell_val.isdigit(): # Only process data rows
                            norm_row = align_to_header(clean_row, wind_header)
                        
                        # --- Append logic ---
                        if norm_row and len(norm_row) == len(wind_header):
                            row_text = " ".join([clean_for_check(c) for c in norm_row if c])
                            if "SEPC" in base_name.upper():
                                if "CLEAN MAX" in row_text or "CLEANMAX" in row_text:
                                    wind_rows.append(norm_row)
                            else:
                                wind_rows.append(norm_row)

                    elif current_section == "solar" and solar_header:
                        if not first_cell_val.isdigit(): # Skip non-data rows
                            continue
                            
                        norm_row = align_to_header(clean_row, solar_header)
                        if norm_row and len(norm_row) == len(solar_header):
                            row_text = " ".join([clean_for_check(c) for c in norm_row if c])
                            if "SEPC" in base_name.upper():
                                if "CLEAN MAX" in row_text or "CLEANMAX" in row_text:
                                    solar_rows.append(norm_row)
                            else:
                                solar_rows.append(norm_row)
    
    # Final check: Remove any rows that are just header remnants
    if wind_header:
        wind_header_text = " ".join([clean_for_check(c) for c in wind_header if c])
        wind_rows = [r for r in wind_rows if " ".join([clean_for_check(c) for c in r if c]) != wind_header_text]
    if solar_header:
        solar_header_text = " ".join([clean_for_check(c) for c in solar_header if c])
        solar_rows = [r for r in solar_rows if " ".join([clean_for_check(c) for c in r if c]) != solar_header_text]

    print(f"--- Final Count for {base_name}: {len(wind_rows)} wind, {len(solar_rows)} solar ---")
    return wind_header, wind_rows, solar_header, solar_rows

# --- MAIN LOOP ---

for filename in os.listdir(input_folder):
    if not filename.lower().endswith(".pdf"):
        continue

    pdf_path = os.path.join(input_folder, filename)
    base_name = os.path.splitext(filename)[0]
    excel_path = os.path.join(output_folder, f"{base_name}.xlsx")
    
    print(f"\n--- Processing {filename} ---")
    wind_header, wind_rows, solar_header, solar_rows = extract_sections(pdf_path)
    
    if not wind_rows and not solar_rows:
        print(f"❌ No Wind/Solar data in: {filename}")
        continue

    date_str = extract_date_from_filename(base_name)

    df_wind = pd.DataFrame(wind_rows, columns=wind_header) if wind_header else pd.DataFrame()
    df_solar = pd.DataFrame(solar_rows, columns=solar_header) if solar_header else pd.DataFrame()

    df_wind = clean_empty_columns(df_wind)
    df_solar = clean_empty_columns(df_solar)

    # --- NEW: Filter unwanted text rows from the DataFrames ---
    pat = '|'.join(UNWANTED_TEXT)

    if not df_wind.empty:
        # Find the "Name" column (it might have newlines from the header)
        wind_name_col = [col for col in df_wind.columns if "WIND FARM OWNER" in str(col).upper()]
        if wind_name_col:
            # Filter rows where the "Name" column contains any unwanted text
            df_wind = df_wind[~df_wind[wind_name_col[0]].astype(str).str.contains(pat, case=False, na=False)]
        else:
            print(f"DEBUG WARNING: Could not find 'Name of Wind Farm Owner' column in {filename} to filter.")

    if not df_solar.empty:
        # Find the "Name" column
        solar_name_col = [col for col in df_solar.columns if "SOLAR ENTITY NAME" in str(col).upper()]
        if solar_name_col:
            # Filter rows where the "Name" column contains any unwanted text
            df_solar = df_solar[~df_solar[solar_name_col[0]].astype(str).str.contains(pat, case=False, na=False)]
        else:
            print(f"DEBUG WARNING: Could not find 'Solar Entity Name' column in {filename} to filter.")
    # --- END NEW FILTER ---

    df_wind = df_wind.replace(r'^\s*$', pd.NA, regex=True)
    df_wind = df_wind.dropna(thresh=2).reset_index(drop=True)
    
    # Do the same for solar
    df_solar = df_solar.replace(r'^\s*$', pd.NA, regex=True)
    df_solar = df_solar.dropna(thresh=2).reset_index(drop=True)

    if not df_wind.empty:
        df_wind.insert(1, "Date", date_str)
        df_wind.rename(columns={"SSr No": "Sr No"}, inplace=True)
        if "Sr No" in df_wind.columns:
            df_wind["Sr No"] = range(1, len(df_wind)+1)

    if not df_solar.empty:
        df_solar.insert(1, "Date", date_str)
        df_solar.rename(columns={"SSr No": "Sr No"}, inplace=True)
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
    app_id="SLDC Gujarat Data Extraction",
    title="PDF to Excel Conversion Complete",
    msg="Wind & Solar data successfully extracted and saved to Excel.",
    duration="short"
)
toast.set_audio(audio.Default, loop=False)
toast.show()