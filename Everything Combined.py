from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from winotify import Notification, audio
import pdfplumber
import pandas as pd
import re
import os
import requests
from datetime import datetime
from collections import defaultdict

def pdf_extraction():
    # --- Configuration ---
    ENERGY_NAMES = ["HETENERGY(BHILDI-HYBRID)",
                    "66KVYASHASWA(HYBRID)",
                    "SANATHAL(HEM_URJA_HYBRID)",
                    "MOTA_DEVLIYA(HETENERGY_HYBRID)",
                    "66KVCLEANMAXPIPARADI(HYBRID)",
                    "SEPC(HYBRID)",
                    "66_KV_MOTA_KHIJADIYA(SALPIPALIYA_WF)",
                    "66_KV_MOTA_KHIJADIYA(SALPIPALIYA_HYBRID)",
                    "DHARAGAR(GNESL)",
                    "66 KV GHELDA(GNESL)",
                    "220KV_NAGPUR(OP_WIND)HYBRID"
                    ]
    YEAR = "2025"
    MONTH_INDEX = {
        "JAN": "1", "FEB": "2", "MAR": "3", "APR": "4",
        "MAY": "5", "JUN": "6", "JUL": "7", "AUG": "8",
        "SEP": "9", "OCT": "10", "NOV": "11", "DEC": "12"
    }
    BASE_URL = "https://www.sldcguj.com/Energy_Block_New.php"
    DOWNLOAD_DIR = "D:/Projects/SLDC Gujarat Web Scraping + Excel Conversion/downloads"
    ICON_PATH = "C:/Users/Hari.Srinivas/Downloads/images.png"

    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    # Setup Chrome
    options = Options()
    options.binary_location = r"C:\Program Files\Google\Chrome\Application\chrome.exe"  # 👈 paste your path here
    options.add_argument("--headless")
    options.add_argument("--window-size=1200,800")
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    wait = WebDriverWait(driver, 20)

    # --- Tracker Summary ---
    downloaded = []
    already_present = []
    no_pdf = []
    skipped_future = []

    try:
        for ENERGY_NAME in ENERGY_NAMES:
            print(f"\n🔄 Processing ENERGY: {ENERGY_NAME}")
            driver.get(BASE_URL)
            wait.until(EC.presence_of_element_located((By.ID, "energy_name")))

            try:
                Select(driver.find_element(By.ID, "energy_name")).select_by_visible_text(ENERGY_NAME)
            except:
                print(f"❌ ENERGY_NAME not found in dropdown: {ENERGY_NAME}")
                continue

            Select(driver.find_element(By.ID, "year")).select_by_visible_text(YEAR)

            for month_name, month_num in MONTH_INDEX.items():
                print(f"📅 {ENERGY_NAME} → {month_name} {YEAR}")
                now = datetime.now()
                month_date = datetime(int(YEAR), int(month_num), 1)
                if month_date > now:
                    print(f"⏩ Skipped {month_name} – Future month.")
                    skipped_future.append(f"{ENERGY_NAME}-{month_name}")
                    continue

                # --- PRE-CHECK: skip if file already exists in DOWNLOAD_DIR ---
                # Expected filename patterns that the downloader uses:
                # 1) ENERGY_NAME_YEAR_MONTH_suffix.pdf
                # 2) base_pdf_name_MONTH_YEAR.pdf
                # We'll look for any file containing the site name and the month/year
                site_key = ENERGY_NAME.replace(" ", "_").replace("(", "").replace(")", "")
                expected_month_year = f"_{month_name}_{YEAR}"
                already_found = False
                for existing in os.listdir(DOWNLOAD_DIR):
                    if not existing.lower().endswith('.pdf'):
                        continue
                    existing_up = existing.upper()
                    # normalize: check if site_key (upper) in filename and month and year present
                    if site_key.upper() in existing_up and month_name.upper() in existing_up and YEAR in existing_up:
                        print(f"✔️ Already exists on disk, skipping scrape: {existing}")
                        already_present.append(f"{ENERGY_NAME}-{month_name}")
                        already_found = True
                        break
                if already_found:
                    continue

                # Refresh dropdown each loop
                Select(driver.find_element(By.ID, "month")).select_by_visible_text(month_name)
                submit_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@type='submit']")))
                driver.execute_script("arguments[0].click();", submit_btn)

                try:
                    pdf_links = wait.until(
                        EC.presence_of_all_elements_located(
                            (By.XPATH, f"//a[contains(@href, '{ENERGY_NAME}') and contains(@href, '.pdf')]")
                        )
                    )
                    
                    if not pdf_links:
                        print(f"❌ No PDFs found for {ENERGY_NAME} → {month_name}")
                        no_pdf.append(f"{ENERGY_NAME}-{month_name}")
                        continue

                    # Try to select most recent version
                    pdf_info = []
                    for link in pdf_links:
                        href = link.get_attribute("href")
                        if "-" in href and href.endswith(".pdf"):
                            suffix = href.split("-")[-1].replace(".pdf", "")
                        else:
                            suffix = ""
                        pdf_info.append((suffix, href))

                    if any(suffix for suffix, _ in pdf_info):
                        # Case: There is a suffix — sort and use it
                        pdf_info.sort(reverse=True, key=lambda x: x[0])
                        selected_suffix, selected_href = pdf_info[0]
                        filename = f"{ENERGY_NAME}_{YEAR}_{month_name}_{selected_suffix}.pdf"
                    else:
                        # Case: No suffix — use the filename from the URL
                        selected_href = pdf_info[-1][1]
                        base_pdf_name = os.path.basename(selected_href).split("?")[0].replace(".pdf", "")
                        filename = f"{base_pdf_name}_{month_name}_{YEAR}.pdf"

                    # Sanitize file name
                    filename = filename.replace(" ", "_").replace("(", "").replace(")", "")
                    file_path = os.path.join(DOWNLOAD_DIR, filename)

                    if os.path.exists(file_path):
                        print(f"✔️ Already exists: {filename}")
                        already_present.append(f"{ENERGY_NAME}-{month_name}")
                        continue

                    print(f"📥 Downloading: {selected_href}")
                    response = requests.get(selected_href)
                    with open(file_path, "wb") as f:
                        f.write(response.content)

                    print(f"✅ Saved: {file_path}")
                    downloaded.append(f"{ENERGY_NAME}-{month_name}")

                except TimeoutException:
                    print(f"❌ Timeout: No PDFs for {ENERGY_NAME} → {month_name}")
                    no_pdf.append(f"{ENERGY_NAME}-{month_name}")

    finally:
        driver.quit()
    from collections import defaultdict

    # --- Grouped Summary by ENERGY_NAME ---
    status_map = defaultdict(lambda: defaultdict(list))

    for entry in downloaded:
        energy, month = entry.rsplit("-", 1)
        status_map[energy]["✅ Downloaded"].append(month)

    for entry in already_present:
        energy, month = entry.rsplit("-", 1)
        status_map[energy]["✔️ Existing"].append(month)

    for entry in no_pdf:
        energy, month = entry.rsplit("-", 1)
        status_map[energy]["❌ Not Found"].append(month)

    for entry in skipped_future:
        energy, month = entry.rsplit("-", 1)
        status_map[energy]["⏩ Skipped"].append(month)

    # --- Format Notification ---
    lines = []
    for energy, statuses in status_map.items():
        lines.append(f"📌 {energy}")
        for status, months in statuses.items():
            months_str = ", ".join(months)
            lines.append(f"  {status}: {months_str}")

    summary_msg = "\n".join(lines) or "No files processed."

    # --- Toast Notification ---
    toast = Notification(
        app_id="SLDC Gujarat Multi-Energy",
        title="🔔 SLDC Download Summary",
        msg=summary_msg,
        duration="long",
        icon=ICON_PATH if os.path.exists(ICON_PATH) else None
    )
    toast.set_audio(audio.Default, loop=False)
    toast.show()

def excel_conversion():
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

def excel_merging():
    input_folder = "D:/Projects/SLDC Gujarat Web Scraping + Excel Conversion/excel_conversion"
    # output_folder = "D:/Projects/SLDC Gujarat Web Scraping + Excel Conversion/all_combined_excel_files"

    output_folder = "D:/OneDrive - CMES/SLDCGuj all Combined Excel"
    os.makedirs(output_folder, exist_ok=True)

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

    energy_sites = defaultdict(list)
    for file in os.listdir(input_folder):
        if file.lower().endswith(".xlsx"):
            site_name = extract_site_name(file)
            energy_sites[site_name].append(file)

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
        duration="short"
    )
    toast.set_audio(audio.Default, loop=False)
    toast.show()

if __name__ == "__main__":
    pdf_extraction()
    excel_conversion()
    excel_merging()