"""Functions removed from active files during the cleanup (2026-11-16).

Kept here only for reference. Nothing imports this file. Not runnable on its own.
"""
import os


# ------------------------------------------------------------
# From excel_utils.py - old single-sheet result helpers (columns Status / Message / LOG).
# Replaced by workbook.py (save_sheet, update_source_columns, temp file + replace).
# save_progress was only used by archive/production_version_fix.py.
# ------------------------------------------------------------
def save_excel(df, file_path):
    df.to_excel(file_path, index=False)


def prepare_result_columns(df):
    if "Status" not in df.columns:
        df["Status"] = ""

    if "Message" not in df.columns:
        df["Message"] = ""

    return df


def log_result(df, idx, status, message):
    df.loc[idx, "Status"] = status
    df.loc[idx, "Message"] = message


def add_log(df, idx, message):
    print(message)
    df.at[idx, "LOG"] += f"{message}\n"


def save_progress(input_file, df):
    # NOTE: overwrites the whole file with ONE sheet (other sheets are lost)
    temp_file = input_file.replace(".xlsx", "_temp.xlsx")
    df.to_excel(temp_file, index=False, engine="openpyxl")
    os.replace(temp_file, input_file)


# ------------------------------------------------------------
# From reader.py - debug print of the BOMs read from Excel (never called).
# ------------------------------------------------------------
def print_boms(boms):
    for i, bom in enumerate(boms, 1):
        print(f"[{i}] {bom.action:<6} parent={bom.parent} alt={bom.alt_bom} "
            f"base={bom.base_qty:g} rows={bom.excel_rows[0]}-{bom.excel_rows[-1]}")
        for item in bom.items:
            print(f"      row {item.excel_row:<4} {item.item_category}  {item.component:<10} "
                f"qty {item.quantity:>9g}  {item.unit:<3} scrap {item.scrap:g}")
