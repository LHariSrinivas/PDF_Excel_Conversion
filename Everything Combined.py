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
import time
import shutil
from datetime import datetime
from collections import defaultdict

def pdf_extraction():
    # ================= CONFIG =================

    ENERGY_NAMES = [
            "HETENERGY(BHILDI-HYBRID)",
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

    DOWNLOAD_DIR = r"D:\Projects\SLDC Gujarat Web Scraping + Excel Conversion\downloads"
    ICON_PATH = r"C:\Users\Hari.Srinivas\Downloads\images.png"
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    # ================= CHROME =================

    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,
    }

    options = Options()
    options.add_experimental_option("prefs", prefs)
    options.add_argument("--headless=new")
    options.add_argument("--window-size=1400,900")
    options.add_experimental_option("excludeSwitches", ["enable-logging"])

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )
    wait = WebDriverWait(driver, 30)

    # ================= HELPERS =================

    def sanitize_name(s):
        for ch in r'<>:"/\\|?*':
            s = s.replace(ch, "")
        return s.replace(" ", "_").replace("(", "").replace(")", "")

    def list_pdfs():
        return [f for f in os.listdir(DOWNLOAD_DIR) if f.lower().endswith(".pdf")]

    def wait_for_download(before_files, timeout=30):
        end = time.time() + timeout
        while time.time() < end:
            current = set(list_pdfs())
            new_files = current - before_files
            if new_files:
                return max(
                    new_files,
                    key=lambda f: os.path.getmtime(os.path.join(DOWNLOAD_DIR, f))
                )
            time.sleep(0.5)
        return None

    def is_valid_pdf(path):
        try:
            with open(path, "rb") as f:
                return f.read(5) == b"%PDF-"
        except:
            return False

    def normalize(s):
        return s.replace(" ", "").upper()

    downloaded = []
    already_present = []
    no_pdf = []
    skipped_future = []

    # ================= MAIN =================

    try:
        for ENERGY_NAME in ENERGY_NAMES:
            print(f"\n🔄 Processing ENERGY: {ENERGY_NAME}")

            driver.get(BASE_URL)
            time.sleep(1)

            # ---------- ENERGY DROPDOWN ----------
            energy_select = wait.until(
                EC.presence_of_element_located((By.ID, "energy_name"))
            )

            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});",
                energy_select
            )

            # Wait until options are populated
            # wait.until(
            #     lambda d: len(
            #         Select(d.find_element(By.ID, "energy_name")).options
            #     ) > 1
            # )

            select_energy = Select(energy_select)

            # print("🔍 ENERGY options visible to Selenium:")
            # for opt in select_energy.options:
            #     print("   →", repr(opt.text))

            target_norm = normalize(ENERGY_NAME)
            matched = False

            for opt in select_energy.options:
                if normalize(opt.text) == target_norm:
                    driver.execute_script("arguments[0].selected = true;", opt)
                    driver.execute_script(
                        "arguments[0].dispatchEvent(new Event('change'));",
                        energy_select
                    )
                    print(f"✅ Selected ENERGY: {opt.text}")
                    matched = True
                    break

            if not matched:
                print(f"❌ ENERGY not found: {ENERGY_NAME}")
                continue

            # ---------- YEAR ----------
            Select(driver.find_element(By.ID, "year")).select_by_visible_text(YEAR)

            # ---------- MONTH LOOP ----------
            for month_name, month_num in MONTH_INDEX.items():
                if datetime(int(YEAR), int(month_num), 1) > datetime.now():
                    print(f"⏩ Skipping future month: {month_name}")
                    continue

                print(f"📅 {ENERGY_NAME} → {month_name} {YEAR}")

                safe_energy = sanitize_name(ENERGY_NAME)
                target_filename = f"{safe_energy}_{YEAR}_{month_name}.pdf"
                target_path = os.path.join(DOWNLOAD_DIR, target_filename)

                if os.path.exists(target_path):
                    print(f"✔️ Already exists: {target_filename}")
                    continue

                Select(driver.find_element(By.ID, "month")).select_by_visible_text(month_name)
                submit_btn = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//button[@type='submit']"))
                )
                driver.execute_script("arguments[0].click();", submit_btn)

                try:
                    wait.until(
                        EC.presence_of_all_elements_located(
                            (By.XPATH, "//a[contains(@href, '.pdf')]")
                        )
                    )

                    all_pdf_links = driver.find_elements(
                        By.XPATH, "//a[contains(@href, '.pdf')]"
                    )

                    energy_norm = ENERGY_NAME.replace(" ", "").upper()

                    filtered_links = []
                    for link in all_pdf_links:
                        href = link.get_attribute("href") or ""
                        href_norm = href.replace(" ", "").upper()

                        # ✅ keep only PDFs that belong to this ENERGY
                        if energy_norm in href_norm:
                            filtered_links.append(link)

                    if not filtered_links:
                        print("❌ No ENERGY-specific PDF links found on page")
                        print("   PDFs seen on page:")
                        for l in all_pdf_links:
                            print("   →", l.get_attribute("href"))
                        continue

                    # ✅ NOW pick the last ENERGY-specific PDF
                    selected_link = filtered_links[-1]

                    print("⬇️ Clicking LAST PDF on page:")
                    print("   ", selected_link.get_attribute("href"))

                    before_files = set(list_pdfs())

                    # 🔴 IMPORTANT: click link (do NOT driver.get)
                    driver.execute_script("arguments[0].click();", selected_link)

                    new_file = wait_for_download(before_files)

                    if not new_file:
                        print("❌ PDF download did not complete")
                        continue

                    new_path = os.path.join(DOWNLOAD_DIR, new_file)

                    if not is_valid_pdf(new_path):
                        print("⚠️ Downloaded file is not a valid PDF")
                        continue

                    os.replace(new_path, target_path)
                    print(f"✅ Saved as: {target_filename}")

                except TimeoutException:
                    print("❌ Timeout waiting for PDF links")

    finally:
        driver.quit()

    print("\n🎯 DONE — all energies processed.")

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

    lines = []
    for energy, statuses in status_map.items():
        lines.append(f"📌 {energy}")
        for status, months in statuses.items():
            lines.append(f"  {status}: {', '.join(months)}")

    summary_msg = "\n".join(lines) if lines else "No files processed."

    # ---------------- Notification ----------------
    toast = Notification(
        app_id="SLDC Gujarat Multi-Energy",
        title="🔔 SLDC PDF Download Summary",
        msg=summary_msg,
        duration="long",
        icon=ICON_PATH if os.path.exists(ICON_PATH) else None
    )
    toast.set_audio(audio.Default, loop=False)
    toast.show()

    print("\n📢 Done. Summary:\n")
    print(summary_msg)

def excel_conversion():
    # Set folder paths
    input_folder = "D:/Projects/SLDC Gujarat Web Scraping + Excel Conversion/downloads"
    output_folder = "D:/Projects/SLDC Gujarat Web Scraping + Excel Conversion/excel_conversion"
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

if __name__ == "__main__":
    pdf_extraction()
    excel_conversion()
    excel_merging()