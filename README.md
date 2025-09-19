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
- Cleans empty columns
- Extracts headers dynamically and captures data rows until `"TOTAL"` is encountered.
- Handles  **multi-page continuity** :

* Wind data starting on one page and flowing into the next is merged seamlessly.
* Solar data starting mid-page and continuing on following pages is aligned properly.
* Fixes **ragged tables** where empty columns collapse by realigning rows against the header using the `align_to_header()` function.
* Normalizes inconsistent headers (e.g., `"SSr No"` → `"Sr No"`).
* Removes completely empty columns.
* Inserts a `Date` column (derived from filename).
* Saves one Excel per PDF inside `excel_conversion/`, with two sheets:

  * **Wind Energy**
  * **Solar Energy**
* Notifies user on completion.

#### How align_to_header() function works?

One big challenge with the PDFs is that  **tables are not consistent across pages** :

* On the first page, the table has the **full header structure** with many empty columns (e.g., `"Sr No"`, 3 blank cols, `"Solar Entity Name"`, `"DISCOM Allocation"`, `"Under REC Mechanism"`,`"Installed Capacity"`, `"Active Energy"`, `"Reactive Energy"`, " ").
* On later pages, `pdfplumber`  **collapses blank columns** , producing  **shorter rows** .
* Example:
  * Page 1 (full header length = 13 cols):
    ['1', '', '', '', 'HARSHA ENGINEERS (SOLAR)', 'UGVCL', '', '', '0.675', '132.954', '', '12.528', '']
  * Page 2 (collapsed → only 7 cols):
    ['3', 'LAXON STEELS LIMITED (SOLAR)', 'UGVCL', '', '0.675', '131.959', '11.603']

Without correction, these wouldn’t line up with the header and merging would fail.


#### How alignment works:

The `align_to_header(row, header, section)` function **realigns rows** to match the full header:

1. **Case 1: Row length = Header length**

   → Perfect match, keep as is.
2. **Case 2: Row shorter than Header length**

   → Expand row into the right header columns:

   * Always map `row[0]` → `"Sr No"`
   * If section =  **Wind** :

     * Owner name → index 4
     * DISCOM → index 5
     * Under REC → index 6
     * Installed Capacity → index 8
     * Active Energy → index 9
     * Reactive Energy → index 11
   * If section =  **Solar** :

     Same mapping but `"Solar Entity Name"` instead of Wind Owner.
   * All other positions are filled with `""`.

   ✅ This ensures a collapsed 7-column row expands into a 13-column row that aligns with the header.
3. **Case 3: Row longer than Header length**

   → Truncate to header size.

#### Why it matters:

* This keeps **all rows consistent** with the header, even across pages with inconsistent layouts.
* Prevents unrelated text (like "Period Considered" or certificates) from being captured.
* Ensures merged Excel files are  **structured, analyzable, and uniform** .

When we run `align_to_header(row, header, section)`, every row is “stretched” to the  **same length as the header** .

* Empty/blank positions are explicitly filled with `""`.
* So even if the PDF page had only 7 columns, we rebuild it into a 13-column row aligned exactly with the header.

This guarantees that the **column structure is consistent** across all pages.

The `clean_empty_columns` drops any columns where:

* The header is blank (`""`), **and**
* All the cells in that column are empty/NaN.

So, columns that existed in the PDF just for spacing or formatting (blank headers, no data under them) get  **removed automatically** .

✅ That’s why in the  **final Excel output you don’t see the blank columns** .

They exist during extraction for proper alignment, but are cleaned away before saving.

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
