from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from winotify import Notification, audio
import pandas as pd
import os
import requests
from datetime import datetime
import pdfplumber

def pdf_extraction():
    # --- Configuration ---
    ENERGY_NAME = "HETENERGY(BHILDI-HYBRID)"
    YEAR = "2025"
    MONTH_INDEX = {
        "JAN": "1", "FEB": "2", "MAR": "3", "APR": "4",
        "MAY": "5", "JUN": "6", "JUL": "7", "AUG": "8",
        "SEP": "9", "OCT": "10", "NOV": "11", "DEC": "12"
    }
    BASE_URL = "https://www.sldcguj.com/Energy_Block_New.php"
    DOWNLOAD_DIR = "D:/SLDC Gujarat Web Scraping + Excel Conversion/downloads"
    ICON_PATH = "C:/Users/Hari.Srinivas/Downloads/images.png"  # Change to your valid .ico path

    # Ensure download directory exists
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    # --- Setup Browser ---
    options = Options()
    options.add_argument("--headless")  # Uncomment for headless run
    options.add_argument("--window-size=1200,800")
    options.add_argument("--disable-logging")
    options.add_experimental_option("excludeSwitches", ["enable-logging"])

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    wait = WebDriverWait(driver, 20)

    # --- Summary Trackers ---
    downloaded = []
    already_present = []
    no_pdf = []
    skipped_future = []

    # --- Main Execution ---
    try:
        driver.get(BASE_URL)
        wait.until(EC.presence_of_element_located((By.ID, "energy_name")))
        Select(driver.find_element(By.ID, "energy_name")).select_by_visible_text(ENERGY_NAME)
        Select(driver.find_element(By.ID, "year")).select_by_visible_text(YEAR)

        for month_name, month_num in MONTH_INDEX.items():
            print(f"\n📅 Processing {month_name} {YEAR}...")
            now = datetime.now()
            month_date = datetime(int(YEAR), int(month_num), 1)
            if month_date > now:
                print(f"⏩ Skipped {month_name} – Future month.")
                skipped_future.append(month_name)
                continue

            # Reselect elements before each submission
            Select(driver.find_element(By.ID, "month")).select_by_visible_text(month_name)
            submit_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@type='submit' and contains(@class, 'btn-primary')]")))
            driver.execute_script("arguments[0].click();", submit_btn)

            try:
                # Wait for all matching PDFs to load
                pdf_links = wait.until(
                    EC.presence_of_all_elements_located(
                        (By.XPATH, f"//a[contains(@href, '{ENERGY_NAME}') and contains(@href, '.pdf')]")
                    )
                )

                if not pdf_links:
                    print(f"❌ No PDFs found for {month_name}")
                    no_pdf.append(month_name)
                    continue

                found_new_pdf = False

                for index, link in enumerate(pdf_links, start=1):
                    pdf_href = link.get_attribute("href")

                    # Generate unique filename
                    if "-" in pdf_href:
                        suffix = pdf_href.split("-")[-1].replace(".pdf", "")
                    else:
                        suffix = f"v{index}"
                    filename = f"{ENERGY_NAME}_{YEAR}_{month_name}_{suffix}.pdf"
                    filename = filename.replace(" ", "_").replace("(", "").replace(")", "")
                    file_path = os.path.join(DOWNLOAD_DIR, filename)

                    if os.path.exists(file_path):
                        print(f"✔️ Already exists: {filename}")
                        continue

                    print(f"📥 Downloading: {pdf_href}")
                    response = requests.get(pdf_href)
                    with open(file_path, "wb") as f:
                        f.write(response.content)
                    print(f"✅ Saved: {file_path}")
                    found_new_pdf = True

                if found_new_pdf:
                    downloaded.append(month_name)
                else:
                    already_present.append(month_name)

            except TimeoutException:
                print(f"❌ Timeout: No PDFs found for {month_name}")
                no_pdf.append(month_name)

    finally:
        driver.quit()

    # --- Notification Summary ---
    summary_lines = []
    if downloaded:
        summary_lines.append(f"📥 New: {', '.join(downloaded)}")
    if already_present:
        summary_lines.append(f"✔️ Existing: {', '.join(already_present)}")
    if no_pdf:
        summary_lines.append(f"❌ Not Available: {', '.join(no_pdf)}")
    if skipped_future:
        summary_lines.append(f"⏩ Skipped: {', '.join(skipped_future)}")

    notification_message = "\n".join(summary_lines) or "No actions taken."

    toast = Notification(
        app_id="SLDC Gujarat Script",
        title="📊 SLDC Gujarat Summary",
        msg=notification_message,
        duration="long",  # Stays ~25 sec then moves to Action Center
        icon=ICON_PATH if os.path.exists(ICON_PATH) else None
    )

    toast.set_audio(audio.Default, loop=False)
    toast.show()

def excel_conversion():
    # Set folder paths
    input_folder = "D:/SLDC Gujarat Web Scraping + Excel Conversion/downloads"
    output_folder = "D:/SLDC Gujarat Web Scraping + Excel Conversion/excel_conversion"
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

if __name__ == "__main__":
    pdf_extraction()
    excel_conversion()