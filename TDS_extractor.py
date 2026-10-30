"""
extract_challans.py

Extracts details from Income Tax "Challan Receipt" PDFs in a folder and
writes them to a single Excel file, with one row per TDS/TCS section
entry (a challan with multiple sections produces multiple rows).

Usage:
    python extract_challans.py <folder_with_pdfs> [output.xlsx]

If no output path is given, it writes "challan_summary.xlsx" inside the
input folder.
"""

import os
import re
import sys
import glob

import pdfplumber
from openpyxl import Workbook


def clean_amount(value):
    """Turn '₹ 1,17,324' -> 117324 (int). Returns 0 for blanks/None."""
    if value is None:
        return 0
    value = str(value).replace("₹", "").replace(",", "").strip()
    if value == "" or value == "-":
        return 0
    try:
        return int(float(value))
    except ValueError:
        return 0


def extract_field(text, label):
    """Grab the value after 'Label :' on its own line."""
    pattern = rf"{re.escape(label)}\s*:\s*(.+)"
    m = re.search(pattern, text)
    return m.group(1).strip() if m else ""


def clean_description(description):
    """Collapse the multi-line description text into one tidy string."""
    text = (description or "").replace("\n", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def extract_section_ref(description):
    """
    Pull just the section reference and Table clause out of the
    description, e.g. '393(1) [Table: Sl. No. 6(iii).D(b)]'.
    """
    text = clean_description(description)
    m = re.search(r"(\d{2,4}\(\d+\)\s*\[Table:[^\]]*\])", text)
    return m.group(1) if m else text


def extract_breakup_field(text, label):
    """
    Pull a value from the 'Tax Breakup Details' section, where lines look
    like 'B Surcharge ₹ 0' or 'F Fee under section 234E ₹ 0'.
    """
    pattern = rf"{re.escape(label)}\s*₹\s*([\d,]+)"
    m = re.search(pattern, text)
    return clean_amount(m.group(1)) if m else 0


def extract_cin(text):
    """
    CIN (Challan Identification Number) is unique per challan - used to
    detect duplicate PDFs (e.g. the same challan downloaded/saved twice
    under different filenames).
    """
    return extract_field(text, "CIN")


def process_pdf(filepath):
    """Return a list of row-dicts, one per section line in the challan."""
    rows = []

    with pdfplumber.open(filepath) as pdf:
        full_text = "\n".join(p.extract_text() or "" for p in pdf.pages)

        challan_no = extract_field(full_text, "Challan No")
        tender_date = extract_field(full_text, "Tender Date")
        bsr_code = extract_field(full_text, "BSR code")
        cin = extract_cin(full_text)

        # Challan-level totals from the "Tax Breakup Details" section
        # (these apply to the whole challan, not to an individual section row)
        surcharge_total = extract_breakup_field(full_text, "B Surcharge")
        cess_total = extract_breakup_field(full_text, "C Cess")
        interest_total = extract_breakup_field(full_text, "D Interest")
        penalty_total = extract_breakup_field(full_text, "E Penalty")
        fee_234e_total = extract_breakup_field(full_text, "F Fee under section 234E")

        # --- Gather every raw table row, across ALL pages, in reading order.
        # A section-wise row can be split across a page break: the row with
        # the S.No/Code/Amount ends up on one page, and the tail end of its
        # (long) description wraps onto the table's continuation on the next
        # page with a blank S.No. We must NOT treat that continuation as its
        # own row - it belongs to the last row that had a real S.No.
        raw_rows = []
        idx_desc = idx_code = idx_total = None
        for page in pdf.pages:
            for table in page.extract_tables():
                if not table:
                    continue
                header = [(c or "").replace("\n", " ").strip() for c in table[0]]
                if "Code" not in header or "Description" not in header:
                    continue  # not the section-wise table

                idx_desc = header.index("Description")
                idx_code = header.index("Code")
                idx_total = len(header) - 1  # last column is "Total (a+b+c)"

                raw_rows.extend(table[1:])

        if idx_desc is None:
            return rows  # no section-wise table found in this PDF at all

        # --- Stitch continuation rows onto the previous row with a real S.No.
        current = None
        for row in raw_rows:
            if not row:
                continue

            sno = (row[0] or "").strip()
            code = (row[idx_code] or "").strip() if row[idx_code] else ""

            if sno.isdigit():
                # Starts a brand-new section row - flush whatever we had.
                if current is not None:
                    rows.append(current)

                current = {
                    "Challan Number": challan_no,
                    "Challan Date": tender_date,
                    "Bsr Code": bsr_code,
                    "Ita 2025 Section Codes": code,
                    "Section": clean_description(row[idx_desc]),
                    "Sum of Tds": clean_amount(row[idx_total]),
                    "Surcharge": surcharge_total,
                    "Cess": cess_total,
                    "Interest": interest_total,
                    "Penalty": penalty_total,
                    "Fee under section 234E": fee_234e_total,
                    "Source File": os.path.basename(filepath),
                    "_CIN": cin,
                }
            else:
                # No S.No -> this is either the "Total" summary row, a blank
                # spacer row, or a continuation of the previous row's
                # wrapped description. Skip the former two, merge the latter.
                desc_piece = clean_description(row[idx_desc]) if row[idx_desc] else ""
                is_total_row = "total" in desc_piece.lower() and not code

                if current is not None and desc_piece and not is_total_row:
                    current["Section"] = (current["Section"] + " " + desc_piece).strip()
                    # In case the code/amount themselves landed on the
                    # continuation line instead of the first line.
                    if code and not current["Ita 2025 Section Codes"]:
                        current["Ita 2025 Section Codes"] = code
                    if row[idx_total]:
                        amt = clean_amount(row[idx_total])
                        if amt:
                            current["Sum of Tds"] = amt

        if current is not None:
            rows.append(current)

    # Now that full descriptions are assembled, trim each one down to just
    # the section reference + Table clause.
    for r in rows:
        r["Section"] = extract_section_ref(r["Section"])

    return rows


COLUMNS = [
    "Challan Number", "Challan Date", "Bsr Code",
    "Ita 2025 Section Codes", "Section", "Sum of Tds",
    "Surcharge", "Cess", "Interest", "Penalty", "Fee under section 234E",
    "Source File",
]


def main():
    # Default to the current folder (the folder the script is run from)
    # if no folder is given on the command line.
    folder = sys.argv[1] if len(sys.argv) > 1 else os.getcwd()
    output_path = sys.argv[2] if len(sys.argv) > 2 else os.path.join(folder, "challan_summary.xlsx")

    print(f"Scanning folder: {folder}")
    pdf_files = sorted(glob.glob(os.path.join(folder, "*.pdf")))
    if not pdf_files:
        print(f"No PDF files found in {folder}")
        return

    all_rows = []
    seen_cins = {}  # CIN -> filename that first produced it, for duplicate detection

    for f in pdf_files:
        try:
            rows = process_pdf(f)
            if not rows:
                print(f"  Warning: no section rows found in {os.path.basename(f)}")
                continue

            cin = rows[0].get("_CIN", "")
            if cin and cin in seen_cins:
                print(f"  Skipping duplicate challan (CIN {cin}) in "
                      f"{os.path.basename(f)} - already extracted from {seen_cins[cin]}")
                continue

            if cin:
                seen_cins[cin] = os.path.basename(f)

            all_rows.extend(rows)
            print(f"  OK: {os.path.basename(f)} -> {len(rows)} row(s)")
        except Exception as e:
            print(f"  Error processing {os.path.basename(f)}: {e}")

    if not all_rows:
        print("No data extracted from any PDF.")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Challans"
    ws.append(COLUMNS)
    for row in all_rows:
        ws.append([row.get(col, "") for col in COLUMNS])

    wb.save(output_path)
    print(f"\nDone. Wrote {len(all_rows)} rows from {len(pdf_files)} PDF(s) to:\n{output_path}")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        import traceback
        print("\nSomething went wrong:\n")
        traceback.print_exc()
    finally:
        # Keep the window open when double-clicked from Windows Explorer so
        # the messages/errors above don't vanish the instant the script ends.
        # (Running from an already-open cmd/PowerShell window is unaffected
        # apart from this one extra prompt.)
        input("\nPress Enter to close this window...")
