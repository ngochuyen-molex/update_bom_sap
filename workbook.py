"""Shared Excel helper for all phases: one workbook, the results of each phase next to
the source or after it.

    CREATE_BOM.xlsx
        Sheet1               source + phase 1 result columns   create_main.py
                             (Created Alt BOM, STATUS, MESSAGE, Updated At per component row)
        Routing Result       phase 2  routing_main.py
        Prod Version Result  phase 3  production_version.py
"""
import os
import shutil
import time
from datetime import datetime
import pandas as pd

# One workbook for the chain create BOM -> assign components -> production version
CHAIN_FILE = r"C:\Users\nguyeb3\Downloads\CREATE_BOM.xlsx"

ROUTING_SHEET = "Routing Result"
PROD_VERSION_SHEET = "Prod Version Result"

# Sheet order after the first (source) sheet; other sheets go to the end.
# "Result" / "Preview" are the sheets of the altbom file (md_assign_component.xlsx).
SHEET_ORDER = [ROUTING_SHEET, "Result", "Preview", PROD_VERSION_SHEET]

# Column names shared by the phases (source sheet + phase 1 result columns)
PARENT_COL = "Parent Material"
ALT_BOM_COL = "Alternative BOM"
CREATED_ALT_BOM_COL = "Created Alt BOM"
STATUS_COL = "STATUS"
MESSAGE_COL = "MESSAGE"
UPDATED_AT_COL = "Updated At"

CONTINUE_STATUSES = ("PENDING", "FAILED")   # processed by "continue"


def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def cell(row, column):
    value = row.get(column)
    return "" if pd.isna(value) else str(value).strip()


def has_sheet(path, name):
    with pd.ExcelFile(path) as xl:      # close the file again, or Windows blocks the next save
        return name in xl.sheet_names


def read_sheet(path, name=0):
    # Sheet by name or index (0 = first / source sheet).
    # All values as text ("0004", "2" stay as they are), blanks as ""
    return pd.read_excel(path, sheet_name=name, dtype=str).fillna("")


def _remove_temp(temp):
    # Leftover temp copy of a failed save
    if os.path.exists(temp):
        try:
            os.remove(temp)
        except OSError:
            pass


def write_sheets(path, sheets):
    """Replaces only the given sheets ({name: DataFrame}).

    Temp copy + replace, so a crash can't corrupt the file. The first sheet is never
    touched; result sheets are kept in SHEET_ORDER.
    """
    temp = path.replace(".xlsx", "_temp.xlsx")
    try:
        shutil.copy2(path, temp)
        with pd.ExcelWriter(temp, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            for name, df in sheets.items():
                df.to_excel(writer, sheet_name=name, index=False)
            book = writer.book
            first, rest = book.worksheets[0], book.worksheets[1:]
            rest.sort(key=lambda ws: SHEET_ORDER.index(ws.title) if ws.title in SHEET_ORDER else len(SHEET_ORDER))
            book._sheets = [first] + rest
            book.active = 0
        os.replace(temp, path)
    finally:
        _remove_temp(temp)


def close_in_excel(path, attempts=3):
    """Saves (your edits) + closes `path` in the running Excel, so the script can write it.

    Saving first means Excel shows no "save changes?" pop-up and nothing you typed is lost.
    Returns True if the workbook was found and closed.
    """
    import pythoncom
    import win32com.client
    name = os.path.basename(path)
    target = os.path.normcase(os.path.abspath(path))
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except pythoncom.com_error:
        print(f"Excel is not running here, so {name} is locked by something else (another user / OneDrive).")
        return False
    for attempt in range(1, attempts + 1):
        try:
            for wb in excel.Workbooks:
                full = wb.FullName
                if os.path.normcase(full) == target or os.path.basename(full).lower() == name.lower():
                    wb.Save()
                    wb.Close(SaveChanges=False)
                    print(f"Saved your changes and closed {name} in Excel.")
                    return True
            print(f"{name} is not open in this Excel (maybe another user / OneDrive / a second Excel).")
            return False
        except pythoncom.com_error as e:
            # e.g. "call was rejected": Excel is busy, usually a cell is in edit mode
            print(f"Excel did not respond (attempt {attempt}/{attempts}): "
                  f"leave cell edit mode (press Enter or Esc in Excel). {e}")
            time.sleep(3)
    return False


def ask_when_locked(path):
    """The file is open/locked. Returns True = try the save again, False = stop."""
    while True:
        print(f"\n{os.path.basename(path)} is open in Excel, results can't be saved.")
        try:
            choice = input("  [s] save your edits + close it in Excel and continue\n"
                           "  [r] I closed it myself, retry\n"
                           "  [q] stop the run\n  > ").strip().lower()
        except EOFError:                # no keyboard (unattended run) -> stop
            choice = "q"
        if choice == "s":
            close_in_excel(path)
            return True
        if choice == "r":
            return True
        if choice == "q":
            return False


def _save_or_ask(path, write):
    # Runs write(); while the file is locked, asks what to do. True = saved.
    while True:
        try:
            write()
            return True
        except PermissionError:
            if not ask_when_locked(path):
                print(f"NOT saved: {path}")
                return False


def save_sheets(path, sheets):
    # write_sheets + the [s]/[r]/[q] prompt when the file is open. True = saved.
    return _save_or_ask(path, lambda: write_sheets(path, sheets))


def save_sheet(path, name, rows, columns):
    # Writes one result sheet. True = saved, False = you chose to stop
    return save_sheets(path, {name: pd.DataFrame(rows, columns=columns)})


def update_source_columns(path, values, columns):
    """Fills result columns in the FIRST sheet, cell by cell (other columns and the
    formatting stay as they are). Missing result columns are added at the end.

    values  {excel_row: {column: value}}   excel_row = row number in Excel (header = 1)
    True = saved, False = file locked and you chose to stop.
    """
    return _save_or_ask(path, lambda: _write_source_columns(path, values, columns))


def _write_source_columns(path, values, columns):
    from openpyxl import load_workbook
    temp = path.replace(".xlsx", "_temp.xlsx")
    try:
        shutil.copy2(path, temp)
        book = load_workbook(temp)
        ws = book.worksheets[0]
        headers = {str(c.value).strip(): c.column for c in ws[1] if c.value is not None}
        for name in columns:
            if name not in headers:
                headers[name] = ws.max_column + 1
                ws.cell(row=1, column=headers[name], value=name)
        for excel_row, row_values in values.items():
            for name, value in row_values.items():
                ws.cell(row=excel_row, column=headers[name], value=value)
        book.save(temp)
        book.close()
        os.replace(temp, path)
    finally:
        _remove_temp(temp)
