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
ENERGY_NAMES = ["HETENERGY(BHILDI-HYBRID)",
            "66KVYASHASWA(HYBRID)",
            "SANATHAL(HEM_URJA_HYBRID)",
            "MOTA_DEVLIYA(HETENERGY_HYBRID)",
            "66KVCLEANMAXPIPARADI(HYBRID)",
            "SEPC(HYBRID)",
            "66_KV_MOTA_KHIJADIYA(SALPIPALIYA_WF)",
            "66_KV_MOTA_KHIJADIYA(SALPIPALIYA_HYBRID)"
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