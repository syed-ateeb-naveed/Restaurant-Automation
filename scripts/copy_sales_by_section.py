"""
copy_sales_by_section.py
------------------------
Copies the "Sales By Section" columns (Bill Count, Net Sales, Remaining,
Skip, DD, Uber) from a Weekly Sales Summary file into the Weekly Sales
Report (main) file.

How the target row is found (same logic as touchbistro.py):
  1. Search column A for TARGET_WEEK_LABEL — if found, write there directly.
  2. If not found, search for PREV_WEEK_LABEL and insert a new row just below it,
     writing TARGET_WEEK_LABEL into column A of that new row.

After writing, a blank row (with formatting copied from the filled row) is
ensured immediately below — so any data sitting further down is never
displaced as new weeks accumulate.

Usage:
    python copy_sales_by_section.py <summary_file> [--target <report_file>]

Examples:
    python copy_sales_by_section.py "Weekly Sales summary (week 5).xlsx"
    python copy_sales_by_section.py "Weekly Sales summary (week 5).xlsx" --target "Weekly Sales Report (main).xlsx"

If --target is omitted, defaults to "Weekly Sales Report (main).xlsx" in the
same folder as the summary file.
"""

import sys
import os
from copy import copy
import openpyxl


# ════════════════════════════════════════════════════════════════
# CONFIG  —  update these before each run
# ════════════════════════════════════════════════════════════════

MASTER_XLSX = r"..\Weekly Reports\Weekly Sales summary (week 5).xlsx"

# Label of the previous week — used to locate the insertion point if the
# target row for this week doesn't exist yet.
PREV_WEEK_LABEL   = "May 18 - May 24"

# Label for the week being filled in.
# If this row already exists in column A the script writes into it directly.
# If it doesn't exist, it creates the row below PREV_WEEK_LABEL.
TARGET_WEEK_LABEL = "May 25 - May 31"


# ════════════════════════════════════════════════════════════════
# Sheet name mapping:  source sheet name  →  target sheet name
# ════════════════════════════════════════════════════════════════

SHEET_MAP = {
    "Highway":           "Highway",
    "Kababwala":         "Kababwala",
    "Pizza K Heartland": "Pizza K Heartland",
    "Pizza K Eglinton":  "Pizza K Eglinton",
    "Karachi Food Court":"Karachi Food Court",
    "Markham":           "Markham",
    "Jane":              "Jane",
    "Ajax":              "Ajax",
    "Queen St.":         "PKQueen",
    "Kababwala - Queen": "KKWQueen",
    "Lebovic":           "Lebovic",
}

# Source columns (1-based): A=date, B=Bill Count, C=Net Sales, D=Remaining,
#                            E=Skip, F=DD, G=Uber
SRC_DATE_COL   = 1
SRC_VALUES_COL = 2   # first of the 6 Sales By Section values (B through G)
NUM_COLS       = 6


# ════════════════════════════════════════════════════════════════
# Helpers
# ════════════════════════════════════════════════════════════════

def find_bill_count_col(ws):
    """Return the column number of the 'Bill Count' header in a target sheet."""
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and "bill count" in str(cell.value).lower():
                return cell.column
    return None


def find_latest_src_row(ws):
    """
    Return the last row in the source sheet that has actual data in the
    Sales By Section columns (C–G). Returns (row_number, date_string).
    """
    last_row = last_date = None
    for r in range(3, ws.max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(3, 8)]
        if any(v is not None for v in vals):
            last_row = r
            last_date = ws.cell(row=r, column=SRC_DATE_COL).value
    return last_row, last_date


def find_target_row(ws):
    """
    Return the row to write into (mirrors touchbistro.py exactly).

    Searches column A for TARGET_WEEK_LABEL first.
    Falls back to PREV_WEEK_LABEL and inserts a new row just below it.
    Raises RuntimeError if neither label is found.
    """
    prev_row = None
    for cell in ws["A"]:
        val = str(cell.value).strip() if cell.value is not None else ""
        if val == TARGET_WEEK_LABEL:
            return cell.row
        if val == PREV_WEEK_LABEL:
            prev_row = cell.row
            break  # use the FIRST occurrence, not the last

    if prev_row is None:
        raise RuntimeError(
            f"Neither '{TARGET_WEEK_LABEL}' nor '{PREV_WEEK_LABEL}' found "
            f"in column A of sheet '{ws.title}'. Check CONFIG labels."
        )

    target_row = prev_row + 1
    ws.insert_rows(target_row)

    # Copy formatting from the PREV_WEEK_LABEL row into the new row
    for col in range(1, ws.max_column + 1):
        src_cell = ws.cell(row=prev_row, column=col)
        tgt_cell = ws.cell(row=target_row, column=col)
        if src_cell.has_style:
            tgt_cell.font         = copy(src_cell.font)
            tgt_cell.border       = copy(src_cell.border)
            tgt_cell.fill         = copy(src_cell.fill)
            tgt_cell.alignment    = copy(src_cell.alignment)
            tgt_cell.protection   = copy(src_cell.protection)
            tgt_cell.number_format = src_cell.number_format

    ws.cell(row=target_row, column=1, value=TARGET_WEEK_LABEL)
    print(f"   ℹ  Created row {target_row} for '{TARGET_WEEK_LABEL}' below '{PREV_WEEK_LABEL}' (formatting copied).")
    return target_row


def copy_row_style(ws, src_row, tgt_row):
    """Copy cell formatting from src_row to tgt_row (values untouched)."""
    for col in range(1, ws.max_column + 1):
        src_cell = ws.cell(row=src_row, column=col)
        tgt_cell = ws.cell(row=tgt_row, column=col)
        if src_cell.has_style:
            tgt_cell.font         = copy(src_cell.font)
            tgt_cell.border       = copy(src_cell.border)
            tgt_cell.fill         = copy(src_cell.fill)
            tgt_cell.alignment    = copy(src_cell.alignment)
            tgt_cell.protection   = copy(src_cell.protection)
            tgt_cell.number_format = src_cell.number_format


def ensure_blank_row_below(ws, row):
    """
    Make sure the row immediately below `row` is blank.
    - Already blank → just re-applies formatting from `row`.
    - Has data     → inserts a new blank row and copies formatting.
    Returns True if a row was inserted, False otherwise.
    """
    next_row = row + 1
    has_data = any(
        ws.cell(row=next_row, column=c).value is not None
        for c in range(1, ws.max_column + 1)
    )
    if has_data:
        ws.insert_rows(next_row)
        copy_row_style(ws, row, next_row)
        return True
    else:
        copy_row_style(ws, row, next_row)
        return False


# ════════════════════════════════════════════════════════════════
# Main logic
# ════════════════════════════════════════════════════════════════

def copy_section(src_file, tgt_file):
    print(f"\nSource : {os.path.basename(src_file)}")
    print(f"Target : {os.path.basename(tgt_file)}")
    print(f"Week   : '{TARGET_WEEK_LABEL}'  (prev: '{PREV_WEEK_LABEL}')")
    print("-" * 60)

    wb_src = openpyxl.load_workbook(src_file)
    wb_tgt = openpyxl.load_workbook(tgt_file)

    any_error = False

    for src_name, tgt_name in SHEET_MAP.items():
        # --- source data ---
        if src_name not in wb_src.sheetnames:
            print(f"  [SKIP] '{src_name}' not found in source")
            continue
        ws_src = wb_src[src_name]
        src_row, _ = find_latest_src_row(ws_src)
        if src_row is None:
            print(f"  [SKIP] '{src_name}': no data found in source")
            continue
        src_vals = [ws_src.cell(row=src_row, column=SRC_VALUES_COL + i).value
                    for i in range(NUM_COLS)]

        # --- target sheet ---
        if tgt_name not in wb_tgt.sheetnames:
            print(f"  [SKIP] '{tgt_name}' not found in target")
            continue
        ws_tgt = wb_tgt[tgt_name]

        bc_col = find_bill_count_col(ws_tgt)
        if bc_col is None:
            print(f"  [ERROR] '{tgt_name}': no 'Bill Count' column found")
            any_error = True
            continue

        # Find or create the target row using PREV/TARGET labels
        try:
            tgt_row = find_target_row(ws_tgt)
        except RuntimeError as exc:
            print(f"  [ERROR] {exc}")
            any_error = True
            continue

        # Write the 6 values
        old_vals = [ws_tgt.cell(row=tgt_row, column=bc_col + i).value for i in range(NUM_COLS)]
        for i, val in enumerate(src_vals):
            ws_tgt.cell(row=tgt_row, column=bc_col + i).value = val

        # Ensure blank buffer row below
        inserted = ensure_blank_row_below(ws_tgt, tgt_row)
        blank_note = "[blank row inserted below]" if inserted else "[blank row already present]"

        changed = old_vals != src_vals
        tag = "UPDATED" if changed else "OK (no change)"
        print(f"  [{tag}] {tgt_name} | row {tgt_row} | "
              f"BC={src_vals[0]}, Net={src_vals[1]}, Rem={src_vals[2]}, "
              f"Skip={src_vals[3]}, DD={src_vals[4]}, Uber={src_vals[5]}")
        if changed:
            print(f"           was: {old_vals}")
        print(f"           {blank_note}")

    try:
        wb_tgt.save(tgt_file)
    except PermissionError:
        print("\n⚠  SAVE FAILED: the target file is open in Excel.")
        print("   Close it and run the script again.")
        sys.exit(1)

    print("-" * 60)
    print(f"Saved → {tgt_file}")
    if any_error:
        print("⚠  Some sheets had errors — check the output above.")
    else:
        print("✓  All done.")


def main():
    args = sys.argv[1:]
    if not args:
        print("Usage: python copy_sales_by_section.py <summary_file> [--target <report_file>]")
        sys.exit(1)

    src_file = args[0]
    tgt_file = None

    if "--target" in args:
        idx = args.index("--target")
        tgt_file = args[idx + 1]

    # Resolve paths relative to the script's own directory
    script_dir = os.path.dirname(os.path.abspath(__file__))
    weekly_reports_dir = os.path.join(os.path.dirname(script_dir), "Weekly Reports")

    if not os.path.isabs(src_file):
        src_file = os.path.join(weekly_reports_dir, src_file)
    if tgt_file is None:
        tgt_file = os.path.join(weekly_reports_dir, "Weekly Sales Report (main).xlsx")
    elif not os.path.isabs(tgt_file):
        tgt_file = os.path.join(weekly_reports_dir, tgt_file)

    if not os.path.exists(src_file):
        print(f"Error: source file not found: {src_file}")
        sys.exit(1)
    if not os.path.exists(tgt_file):
        print(f"Error: target file not found: {tgt_file}")
        sys.exit(1)

    copy_section(src_file, tgt_file)


if __name__ == "__main__":
    main()
