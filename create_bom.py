import openpyxl
from connect import get_session, SAPConnectionError
from excel_utils import unmerge_and_fill
from reader import read_boms, validate_boms, print_boms
from bom import open_cs02, open_bom, find_component_row, set_text, press_enter, save_bom, BOM_TABLE_ID
from bom_create import open_cs01, open_bom_cs01, enter_items, get_created_alt_bom, format_quantity, QUANTITY_COL

INPUT_FILE = r"C:\Users\nguyeb3\Downloads\CHANGE_BOM.xlsx"

# DRY_RUN = read + validate + preview, no SAP
# TEST    = fill SAP screens for the first TEST_LIMIT BOMs, do NOT save
# RUN     = fill and save every BOM
MODE = "DRY_RUN"
TEST_LIMIT = 1

# Scrap and base quantity steps for CS01 are not recorded yet.
# Keep False so RUN does not save incomplete new BOMs.
CREATE_SAVE_READY = False

RESULT_COLUMNS = ("STATUS", "MESSAGE", "CREATED ALT BOM")


# ============================================================
# SAP FLOWS (one BOM each, always from a fresh transaction)
# ============================================================
def get_status_bar(session):
    sbar = session.findById("wnd[0]/sbar")
    return sbar.MessageType, sbar.Text.strip()

def check_no_error(session, step):
    msg_type, text = get_status_bar(session)
    if msg_type in ("E", "A"):
        raise RuntimeError(f"{step}: SAP error '{text}'")
    return text

def create_bom(session, bom, save):
    open_cs01(session)
    open_bom_cs01(session, bom.parent, bom.usage, bom.plant)
    check_no_error(session, "Open CS01")

    bom.created_alt_bom = get_created_alt_bom(session)
    enter_items(session, bom.items)
    press_enter(session)
    check_no_error(session, "Enter items")

    warning = ""
    if bom.created_alt_bom != bom.alt_bom:
        warning = f" WARNING: SAP alt BOM {bom.created_alt_bom}, Excel alt BOM {bom.alt_bom}."

    if not save:
        return f"Filled {len(bom.items)} items, not saved (TEST).{warning}"
    if not CREATE_SAVE_READY:
        raise RuntimeError("CS01 save disabled: scrap and base quantity steps are not implemented yet")

    save_bom(session)
    return check_no_error(session, "Save") + warning

def change_bom_quantity(session, bom, save):
    item = bom.items[0]
    open_cs02(session)
    open_bom(session, bom.parent, bom.alt_bom, bom.usage, bom.plant)
    check_no_error(session, "Open CS02")

    row_idx = find_component_row(session, item.component)
    set_text(session, f"{BOM_TABLE_ID}/{QUANTITY_COL.format(row=row_idx)}", format_quantity(item.quantity))
    press_enter(session)
    check_no_error(session, "Enter quantity")

    if not save:
        return f"Quantity {format_quantity(item.quantity)} entered for {item.component}, not saved (TEST)."

    save_bom(session)
    return check_no_error(session, "Save")


# ============================================================
# RESULT FILE
# ============================================================
def prepare_result_sheet(ws):
    headers = {ws.cell(1, c).value: c for c in range(1, ws.max_column + 1)}
    for name in RESULT_COLUMNS:
        if name not in headers:
            col = ws.max_column + 1
            ws.cell(1, col).value = name
            headers[name] = col
    return headers

def write_result(ws, headers, bom):
    for row in bom.excel_rows:
        ws.cell(row, headers["STATUS"]).value = bom.status
        ws.cell(row, headers["MESSAGE"]).value = bom.message
        ws.cell(row, headers["CREATED ALT BOM"]).value = bom.created_alt_bom

def save_result(wb, result_file):
    try:
        wb.save(result_file)
    except PermissionError:
        print(f"Cannot save {result_file}: close it in Excel. Progress will be saved on the next BOM.")


# ============================================================
# MAIN
# ============================================================
def main():
    unmerged_file = unmerge_and_fill(INPUT_FILE)
    boms = read_boms(unmerged_file)
    print(f"\n{len(boms)} BOMs: "
        f"{sum(b.action == 'CREATE' for b in boms)} CREATE, "
        f"{sum(b.action == 'CHANGE' for b in boms)} CHANGE\n")
    print_boms(boms)

    problems = validate_boms(boms)
    if problems:
        print("\nVALIDATION PROBLEMS - fix the Excel before running SAP:")
        for p in problems:
            print("  -", p)
        return
    print("\nValidation OK")

    if MODE == "DRY_RUN":
        print("DRY_RUN: SAP not touched. Set MODE = 'TEST' or 'RUN' to continue.")
        return

    try:
        session = get_session()
    except SAPConnectionError as e:
        print(f"SAP connection failed: {e}")
        return

    save = MODE == "RUN"
    if MODE == "TEST":
        boms = boms[:TEST_LIMIT]

    result_file = unmerged_file.replace(".xlsx", "_result.xlsx")
    wb = openpyxl.load_workbook(unmerged_file)
    ws = wb.active
    headers = prepare_result_sheet(ws)

    for i, bom in enumerate(boms, 1):
        print(f"\n===== [{i}/{len(boms)}] {bom.action} {bom.parent} alt {bom.alt_bom} =====")
        try:
            if bom.action == "CREATE":
                bom.message = create_bom(session, bom, save)
            else:
                bom.message = change_bom_quantity(session, bom, save)
            bom.status = "SUCCESS" if save else "TESTED"
        except Exception as e:
            bom.status, bom.message = "FAILED", str(e)
        print(f"{bom.status}: {bom.message}")

        write_result(ws, headers, bom)
        save_result(wb, result_file)

    print(f"\nDone. "
        f"{sum(b.status == 'SUCCESS' for b in boms)} SUCCESS, "
        f"{sum(b.status == 'TESTED' for b in boms)} TESTED, "
        f"{sum(b.status == 'FAILED' for b in boms)} FAILED. Results: {result_file}")


if __name__ == "__main__":
    main()
