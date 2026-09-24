import pdfplumber
import os
import re

# ================= CONFIG =================
TARGET = r"D:\Projects\SLDC Gujarat Web Scraping + Excel Conversion\downloads"
NAME_FILTER = "220KV_NAGPUROP_WINDHYBRID_2026"          # e.g. "NAGPUR" or "AUG"; "" = all files
MAX_CELL_LEN = 40
# =========================================


def preview(cell):
    if cell is None:
        return "."                       # marks a None/blank cell
    s = str(cell).replace("\n", "<NL>")  # show line breaks inside cells
    return s[:MAX_CELL_LEN]


def inspect_pdf(pdf_path):
    base = os.path.basename(pdf_path)
    print("\n" + "=" * 90)
    print(f"FILE: {base}")
    print("=" * 90)

    with pdfplumber.open(pdf_path) as pdf:
        print(f"Total pages: {len(pdf.pages)}")

        for page_num, page in enumerate(pdf.pages, start=1):
            print("\n" + "-" * 90)
            print(f"### PAGE {page_num}  (size: {round(page.width)} x {round(page.height)})")
            print("-" * 90)

            tables = page.extract_tables()
            if not tables:
                print("   [!] No tables detected on this page.")
                txt = (page.extract_text() or "").strip().splitlines()
                for line in txt[:8]:
                    print("   text:", line[:80])
                continue

            print(f"   Tables detected: {len(tables)}")

            for t_idx, table in enumerate(tables):
                print(f"\n   +-- TABLE {t_idx}  ({len(table)} rows) --------------------")
                for r_idx, row in enumerate(table):
                    if row is None:
                        print(f"   | row {r_idx:>2}: <None>")
                        continue
                    first = str(row[0]).strip() if row and row[0] else ""
                    tag = "DATA" if first[:1].isdigit() else "    "
                    cells = " | ".join(preview(c) for c in row)
                    print(f"   | row {r_idx:>2} [{tag}] len={len(row):>2}: {cells}")
                print("   +----------------------------------------------------")


def main():
    if os.path.isfile(TARGET):
        inspect_pdf(TARGET)
    elif os.path.isdir(TARGET):
        pdfs = [f for f in os.listdir(TARGET) if f.lower().endswith(".pdf")]
        if NAME_FILTER:
            pdfs = [f for f in pdfs if NAME_FILTER.upper() in f.upper()]
        if not pdfs:
            print("No matching PDFs found.")
            return
        print(f"Found {len(pdfs)} PDF(s) to inspect.")
        for f in sorted(pdfs):
            inspect_pdf(os.path.join(TARGET, f))
    else:
        print(f"TARGET path does not exist: {TARGET}")


if __name__ == "__main__":
    main()