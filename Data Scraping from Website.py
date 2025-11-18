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
import os
import shutil
import time
from datetime import datetime
import os

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

# set your folder
DOWNLOAD_DIR = r"D:\Projects\SLDC Gujarat Web Scraping + Excel Conversion\downloads"
ICON_PATH = r"C:\Users\Hari.Srinivas\Downloads\images.png"
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# ---------------- Chrome Setup ----------------
chrome_path = shutil.which("chrome") or shutil.which("google-chrome")
if not chrome_path:
    chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"  # fallback

prefs = {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "plugins.always_open_pdf_externally": True,
    "profile.default_content_setting_values.automatic_downloads": 1
}

options = Options()
options.binary_location = chrome_path
options.add_experimental_option("prefs", prefs)
options.add_argument("--headless=new")
options.add_argument("--window-size=1200,800")
options.add_experimental_option("excludeSwitches", ["enable-logging"])

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver, 20)

# ---------------- Helpers ----------------
def sanitize_name(s: str) -> str:
    # Keep letters, numbers, -, _, remove others and collapse spaces
    bad = r'<>:"/\\|?*'
    for ch in bad:
        s = s.replace(ch, "")
    s = s.replace(" ", "_")
    s = s.replace("(", "").replace(")", "")
    return s

def list_pdf_files(folder):
    return [f for f in os.listdir(folder) if f.lower().endswith(".pdf")]

def list_partial_files(folder):
    # capture Chrome partial download extension .crdownload
    return [f for f in os.listdir(folder) if f.lower().endswith(".crdownload")]

def wait_for_new_file(folder, before_set, timeout=20, poll=0.5):
    """
    Wait up to timeout seconds for a new .pdf file to appear in folder that wasn't in before_set.
    Returns the new filename (basename) or None.
    """
    end = time.time() + timeout
    while time.time() < end:
        current = set(list_pdf_files(folder))
        new = current - before_set
        # ensure no .crdownload present (download still in progress)
        if new and not list_partial_files(folder):
            # return the newest by modified time
            new_files = list(new)
            new_files.sort(key=lambda n: os.path.getmtime(os.path.join(folder, n)), reverse=True)
            return new_files[0]
        time.sleep(poll)
    return None

def is_valid_pdf(path):
    try:
        with open(path, "rb") as f:
            header = f.read(5)
        return header == b"%PDF-"
    except Exception:
        return False

# ---------------- Trackers ----------------
downloaded = []
already_present = []
no_pdf = []
skipped_future = []

# ---------------- Main ----------------
try:
    for ENERGY_NAME in ENERGY_NAMES:
        normalized_energy_name = ENERGY_NAME.replace(" ", "").upper()
        print(f"\n🔄 Processing ENERGY: {ENERGY_NAME}")
        driver.get(BASE_URL)
        # wait for dropdown (robust)
        try:
            wait.until(EC.presence_of_element_located((By.ID, "energy_name")))
        except TimeoutException:
            print("❌ energy_name dropdown not found; skipping this ENERGY.")
            continue

        # select energy; if not found skip
        try:
            Select(driver.find_element(By.ID, "energy_name")).select_by_visible_text(ENERGY_NAME)
        except Exception:
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

            # prepare target filename (deterministic)
            safe_energy = sanitize_name(ENERGY_NAME)
            target_filename = f"{safe_energy}_{YEAR}_{month_name}.pdf"
            target_path = os.path.join(DOWNLOAD_DIR, target_filename)

            # If already exists, skip (no overwrite)
            if os.path.exists(target_path):
                print(f"✔️ Already exists (skipping): {target_filename}")
                already_present.append(f"{ENERGY_NAME}-{month_name}")
                # still select month in UI to keep consistent state, then continue
                Select(driver.find_element(By.ID, "month")).select_by_visible_text(month_name)
                submit_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@type='submit']")))
                driver.execute_script("arguments[0].click();", submit_btn)
                time.sleep(0.5)
                continue

            # Select month and submit
            Select(driver.find_element(By.ID, "month")).select_by_visible_text(month_name)
            submit_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@type='submit']")))
            driver.execute_script("arguments[0].click();", submit_btn)

            try:
                # Wait for PDF links to become present on the page (any .pdf link)
                wait.until(EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf')]")))
                # collect links and filter by normalized energy name presence in href
                all_links = driver.find_elements(By.XPATH, "//a[contains(@href, '.pdf')]")
                pdf_links = []
                for l in all_links:
                    href = l.get_attribute("href") or ""
                    if normalized_energy_name in href.replace(" ", "").upper():
                        pdf_links.append(href)

                if not pdf_links:
                    print(f"❌ No PDFs found for {ENERGY_NAME} → {month_name}")
                    no_pdf.append(f"{ENERGY_NAME}-{month_name}")
                    continue

                # choose latest: prefer suffix timestamps; if none, choose last link in page order
                pdf_info = []
                for href in pdf_links:
                    if "-" in href and href.lower().endswith(".pdf"):
                        suffix = href.split("-")[-1].replace(".pdf", "")
                    else:
                        suffix = ""
                    pdf_info.append((suffix, href))

                # if any suffix present, sort by suffix desc; else keep original order and choose last
                if any(suf for suf, _ in pdf_info):
                    pdf_info.sort(reverse=True, key=lambda x: x[0])
                    selected_href = pdf_info[0][1]
                else:
                    selected_href = pdf_info[-1][1]

                # BEFORE downloading: snapshot existing pdf files
                before_files = set(list_pdf_files(DOWNLOAD_DIR))

                # Trigger the browser to GET the PDF (Chrome will save it)
                # Use driver.get so same session/cookies used
                print(f"📥 Triggering browser download for: {selected_href}")
                driver.get(selected_href)

                # Wait for new file to appear
                new_file = wait_for_new_file(DOWNLOAD_DIR, before_files, timeout=25, poll=0.7)
                if not new_file:
                    print(f"❌ No new downloaded file detected for {ENERGY_NAME}-{month_name}")
                    no_pdf.append(f"{ENERGY_NAME}-{month_name}")
                    continue

                new_path = os.path.join(DOWNLOAD_DIR, new_file)
                # Ensure the file is fully written (no partial)
                # double-check there is no .crdownload now
                time.sleep(0.5)

                # Validate it's a real PDF
                if not is_valid_pdf(new_path):
                    print(f"⚠️ Downloaded file is not a valid PDF: {new_file}")
                    # remove invalid file to keep folder clean
                    try:
                        os.remove(new_path)
                    except Exception:
                        pass
                    no_pdf.append(f"{ENERGY_NAME}-{month_name}")
                    continue

                # Rename/move to deterministic target filename
                if os.path.exists(target_path):
                    # race condition: target exists now, skip rename and mark already present
                    print(f"✔️ Target already exists after download: {target_filename}. Removing newly downloaded file.")
                    try:
                        os.remove(new_path)
                    except Exception:
                        pass
                    already_present.append(f"{ENERGY_NAME}-{month_name}")
                else:
                    os.replace(new_path, target_path)
                    print(f"✅ Saved and renamed to: {target_filename}")
                    downloaded.append(f"{ENERGY_NAME}-{month_name}")

            except TimeoutException:
                print(f"❌ Timeout waiting for PDF links for {ENERGY_NAME} → {month_name}")
                no_pdf.append(f"{ENERGY_NAME}-{month_name}")

finally:
    driver.quit()

# ---------------- Summary ----------------
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
