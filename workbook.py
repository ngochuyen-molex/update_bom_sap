"""Shared Excel helper for all phases: one workbook, the results of each phase next to
the source or after it.

    CREATE_BOM.xlsx
        Sheet1               source + phase 1 result columns   bom_create_main.py
                             (Created Alt BOM, STATUS, MESSAGE, Updated At per component row)
        Routing Result       phase 2  routing_main.py
        Prod Version Result  phase 3  production_version.py

Also has the input helpers of the standalone tools (load_excel, unmerge_and_fill),
formerly excel_utils.py.
"""
import copy
import os
import shutil
import time
from datetime import datetime
import openpyxl
import pandas as pd
from config import CHAIN_FILE      # re-exported: phases import it from here

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


def load_excel(file_path):
    # Old style read: numbers stay numbers, blanks stay NaN (bom.py checks rely on that).
    # Used by bom_update_main.py, update_baseqty.py, update_workcenter_efficiency.py.
    # New code: use read_sheet (keeps leading zeros). See TODO-3.
    return pd.read_excel(file_path)


def _copy_cell(src, dst):
    dst.value = src.value
    dst._style = copy.copy(src._style)


def unmerge_and_fill(input_file, output_file=None,
                     group_cols=("Parent Material", "Alternative BOM"),
                     fill_in_group_cols=("Base Qty after",),
                     component_col="Component after"):
    """
    Unmerge all merged cells and fill every row with its value (keeping format).
    Then fill blank group columns (e.g. Parent Material) from the row above,
    and fill blank fill_in_group_cols within the same Parent Material + Alt BOM.
    Saves to output_file (default: <input>_unmerged.xlsx). Input is not changed.
    """
    if output_file is None:
        output_file = input_file.replace(".xlsx", "_unmerged.xlsx")

    wb = openpyxl.load_workbook(input_file)
    ws = wb.active

    # 1. Unmerge and copy top-left value to every cell of the range
    ranges = list(ws.merged_cells.ranges)
    for r in ranges:
        top = ws.cell(r.min_row, r.min_col)
        value, style = top.value, top._style
        ws.unmerge_cells(str(r))
        for row in range(r.min_row, r.max_row + 1):
            for col in range(r.min_col, r.max_col + 1):
                cell = ws.cell(row, col)
                cell.value = value
                cell._style = copy.copy(style)
    print(f"Unmerged {len(ranges)} ranges")

    headers = {ws.cell(1, c).value: c for c in range(1, ws.max_column + 1)}
    for name in (*group_cols, *fill_in_group_cols, component_col):
        if name not in headers:
            raise ValueError(f"Column '{name}' not found in {input_file}")
    comp_c = headers[component_col]

    # 2. Fill blank group columns from the row above (only rows that have a component)
    for name in group_cols:
        c = headers[name]
        for row in range(3, ws.max_row + 1):
            cell = ws.cell(row, c)
            if cell.value in (None, "") and ws.cell(row, comp_c).value not in (None, ""):
                _copy_cell(ws.cell(row - 1, c), cell)
                print(f"Row {row}: filled '{name}' = {cell.value} from row above")

    # 3. Fill blank columns within the same group
    group_c = [headers[n] for n in group_cols]
    for name in fill_in_group_cols:
        c = headers[name]
        key, first = None, None
        for row in range(2, ws.max_row + 1):
            row_key = tuple(ws.cell(row, g).value for g in group_c)
            cell = ws.cell(row, c)
            if row_key != key:
                key, first = row_key, cell
            elif cell.value in (None, ""):
                _copy_cell(first, cell)

    # 4. Report rows still missing required values
    missing = [row for row in range(2, ws.max_row + 1)
               if ws.cell(row, comp_c).value not in (None, "")
               and any(ws.cell(row, headers[n]).value in (None, "")
                       for n in (*group_cols, *fill_in_group_cols))]
    if missing:
        print(f"WARNING: rows still missing values: {missing}")

    wb.save(output_file)
    print(f"Saved: {output_file}")
    return output_file


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
