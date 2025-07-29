# 📊 SLDC Gujarat Energy Data Pipeline

This project automates the full pipeline of **scraping**, **extracting**, **converting**, and **merging** monthly energy PDF reports (Wind & Solar) from the [SLDC Gujarat website](https://www.sldcguj.com). The processed data is saved in structured Excel files, categorized by energy site and month.

---

## ⚛️ Features

- Automated browser interaction with Selenium with silent background execution.
- Dynamic downloading of monthly PDFs per energy site.
- Only downloads latest PDF in case multiple PDF's are present.
- PDF data extraction using `pdfplumber`.
- Excel conversion of extracted Wind & Solar data.
- Merging Excel files for each energy site into single master files.
- Notifications with status summaries for every phase.

---

## 📁 Directory Structure

```
project-root/
├── downloads/                     # Raw downloaded PDFs
├── excel_conversion/             # Excel files converted from PDFs
├── all_combined_excel_files/    # Merged Excel files (one per site)
├── main_script.py                # Your main script file
└── README.md
```

---

## ⚙️ Prerequisites

Make sure the following Python packages are installed:

```bash
pip install selenium webdriver-manager winotify pdfplumber pandas openpyxl requests
```

Also, ensure **Google Chrome** and a stable **Chrome Driver Manager** (version should be same as your current Chrome version) is installed on your system.

---

## 🧠 How It Works

The program consists of **three main phases**:

### 1️⃣ `pdf_extraction()`

**Purpose**: Downloads monthly PDFs from SLDC Gujarat for a list of energy sites.

#### Logic:

- Uses Selenium with Chrome in headless mode (doesn't open Chrome while script is running) to interact with the website.
- Selects `Energy Name`, `Year`, and `Month` from dropdowns.
- Locates and downloads the latest `.pdf` file per energy site per month.
- Skips months in the future.
- Saves PDFs in the `downloads/` folder.
- Sends a desktop notification with a summary:
  - ✅ Downloaded
  - ❌ Not Found
  - ⏩ Skipped (future months)

#### Customizable:

- Energy sites can be added/removed in the `ENERGY_NAMES` list.
- `YEAR` and `DOWNLOAD_DIR` are configurable.
- Uses `winotify` for toast notifications.

---

### 2️⃣ `excel_conversion()`

**Purpose**: Extracts **Wind** and **Solar** generation data from downloaded PDFs and converts them into Excel format.

#### Logic:

- Opens each PDF using `pdfplumber`.
- Detects table sections using keywords:
  - `"SHARE OF WIND FARM OWNER"`
  - `"SHARE OF SOLAR GENERATOR"`
- Extracts headers and row data up to the "TOTAL" marker.
- Cleans empty columns.
- Adds a `Date` column extracted from the filename.
- Saves Excel files in the `excel_conversion/` folder with two sheets:
  - **Wind Energy**
  - **Solar Energy**
- Notifies user on completion.

---

### 3️⃣ `excel_merging()`

**Purpose**: Combines individual monthly Excel files into one master Excel file per energy site.

#### Logic:

- Groups Excel files by energy site name.
- Sorts them by month (based on filename).
- Reads and appends Wind and Solar data across months.
- Adds a `Sr No` column for row indexing.
- Saves combined files in a configured output folder.
- Final notification alerts when all merging is complete.

---

## 🔔 Notifications

The script uses `winotify` to display Windows toast notifications during:

- PDF download completion
- Excel conversion completion
- Final merge completion

To include a custom icon in the notification, set `ICON_PATH` in `pdf_extraction()` (Optional)

---

## 📝 Configuration

Inside the script, these are the key paths you might want to customize:

```python
DOWNLOAD_DIR = "downloads/pdfdirectory"
EXCEL_OUTPUT_DIR = "downloads/excelconversion"
MERGE_OUTPUT_DIR = "downloads/allmerged_directory"
ICON_PATH = "C:/Users/<YourName>/Downloads/images.png"
```

Update these to match your system paths before running.

---

## ▶️ Running the Script

Just run the script using:

```bash
python Everything Combined.py
```

All three functions (`pdf_extraction`, `excel_conversion`, `excel_merging`) will execute sequentially.

---

## 🧹 Dependencies

- `selenium` – Browser automation
- `webdriver-manager` – Automatically manage chromedriver
- `pdfplumber` – Extract text and tables from PDFs
- `pandas` – Data manipulation and Excel conversion
- `openpyxl` – Write to Excel files
- `winotify` – Windows toast notifications
- `requests` – PDF download

requirements.txt file has been given for all the dependencies.

---

## 📬 License

This project is for educational and internal automation purposes.
