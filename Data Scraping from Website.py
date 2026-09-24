from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from winotify import Notification, audio
from collections import defaultdict
from datetime import datetime
import os
import shutil
import time

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