from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from winotify import Notification, audio

import os
import requests
from datetime import datetime

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

            # Attempt to sort by suffix (timestamp or ID), fallback to last link
            pdf_info = []
            for link in pdf_links:
                href = link.get_attribute("href")
                if "-" in href and href.endswith(".pdf"):
                    suffix = href.split("-")[-1].replace(".pdf", "")
                else:
                    suffix = ""  # No suffix
                pdf_info.append((suffix, href))

            if any(suffix for suffix, _ in pdf_info):
                # If at least one suffix is present, sort by suffix
                pdf_info.sort(reverse=True, key=lambda x: x[0])
                selected_suffix, selected_href = pdf_info[0]
            else:
                # No suffixes — fallback to last link
                selected_suffix, selected_href = "vLast", pdf_info[-1][1]

            # Generate filename
            filename = f"{ENERGY_NAME}_{YEAR}_{month_name}_{selected_suffix}.pdf"
            filename = filename.replace(" ", "_").replace("(", "").replace(")", "")
            file_path = os.path.join(DOWNLOAD_DIR, filename)

            if os.path.exists(file_path):
                print(f"✔️ Already exists: {filename}")
                already_present.append(month_name)
                continue

            print(f"📥 Downloading latest: {selected_href}")
            response = requests.get(selected_href)
            with open(file_path, "wb") as f:
                f.write(response.content)
            print(f"✅ Saved: {file_path}")
            downloaded.append(month_name)


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