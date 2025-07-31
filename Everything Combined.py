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
                "SEPC(HYBRID)"
                ]
    YEAR = "2025"
    MONTH_INDEX = {
        "JAN": "1", "FEB": "2", "MAR": "3", "APR": "4",
        "MAY": "5", "JUN": "6", "JUL": "7", "AUG": "8",
        "SEP": "9", "OCT": "10", "NOV": "11", "DEC": "12"
    }
    BASE_URL = "https://www.sldcguj.com/Energy_Block_New.php"
    DOWNLOAD_DIR = "D:/SLDC Gujarat Web Scraping + Excel Conversion/downloads"
    ICON_PATH = "C:/Users/Hari.Srinivas/Downloads/images.png"

    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    # Setup Chrome
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--window-size=1200,800")
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    wait = WebDriverWait(driver, 10)

    # --- Tracker Summary ---
    downloaded = []
    # already_present = [] Commenting this part out to tackle new revision if it occurs.
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

                    # if os.path.exists(file_path):
                    #     print(f"✔️ Already exists: {filename}")
                    #     already_present.append(f"{ENERGY_NAME}-{month_name}")
                    #     continue

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

    # --- Grouped Summary by ENERGY_NAME ---
    status_map = defaultdict(lambda: defaultdict(list))

    for entry in downloaded:
        energy, month = entry.rsplit("-", 1)
        status_map[energy]["✅ Downloaded"].append(month)

    # for entry in already_present:
    #     energy, month = entry.rsplit("-", 1)
    #     status_map[energy]["✔️ Existing"].append(month)

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
        duration="short",
        icon=ICON_PATH if os.path.exists(ICON_PATH) else None
    )
    toast.set_audio(audio.Default, loop=False)
    toast.show()

def excel_conversion():

    input_folder = "D:/SLDC Gujarat Web Scraping + Excel Conversion/downloads"
    output_folder = "D:/SLDC Gujarat Web Scraping + Excel Conversion/excel_conversion"
    os.makedirs(output_folder, exist_ok=True)

    def extract_date_from_filename(base_name):
        match = re.search(r"_(\d{4})_([A-Z]{3})_", base_name)
        if match:
            year = match.group(1)
            mon = match.group(2).upper()
        else:
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
                                wind_rows.append(clean_row)
                        elif current_section == "solar" and not solar_done:
                            if not solar_header:
                                solar_header = clean_row
                            elif len(clean_row) == len(solar_header):
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

            if "Sr No" in df_wind.columns:
                df_wind["Sr No"] = range(1, len(df_wind)+1)

        if not df_solar.empty:
            df_solar.insert(1, "Date", date_str)

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
        msg="Wind & Solar data (across pages) saved successfully.",
        duration="short"
    )
    toast.set_audio(audio.Default, loop=False)
    toast.show()

def excel_merging():
    input_folder = "D:/SLDC Gujarat Web Scraping + Excel Conversion/excel_conversion"
    # output_folder = "D:/SLDC Gujarat Web Scraping + Excel Conversion/all_combined_excel_files"

    output_folder = "C:/Users/Hari.Srinivas/OneDrive - CMES/SLDCGuj all Combined Excel"
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