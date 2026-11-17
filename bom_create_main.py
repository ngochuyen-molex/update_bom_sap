"""Phase 1 - create BOMs in CS01 from the source sheet.

Usage:
    python bom_create_main.py [check] [slow]

    (no option)  read Sheet1, create every BOM that is not SUCCESS yet
    check        read + validate Sheet1, write the result columns (PENDING / SKIPPED / FAILED), no SAP
    slow         wait after every SAP action

Results are written into Sheet1 itself, on every component row of the BOM, in the
columns Created Alt BOM / STATUS / MESSAGE / Updated At (added at the end if missing).
All other columns stay as they are.

A run always continues: BOMs whose rows are SUCCESS are never created again
(that would create a duplicate BOM in SAP).
"""
import sys
import time
from connect import get_session, slow_from_args, SAPConnectionError
from bom import save_bom, press_enter
from bom_create import open_cs01, open_bom_cs01, enter_items, get_created_alt_bom
from reader import read_boms, validate_boms
from utils import clean_component, clean_text
from workbook import (CHAIN_FILE, PARENT_COL, ALT_BOM_COL, CREATED_ALT_BOM_COL,
                      STATUS_COL, MESSAGE_COL, UPDATED_AT_COL,
                      update_source_columns, read_sheet, now, cell)

RESULT_COLUMNS = [CREATED_ALT_BOM_COL, STATUS_COL, MESSAGE_COL, UPDATED_AT_COL]


def previous_results(path):
    # {(material, alt BOM): (status, message, created alt BOM, updated at)} from Sheet1,
    # taken from the first component row of each BOM
    result = {}
    for _, row in read_sheet(path, 0).iterrows():
        key = (clean_component(cell(row, PARENT_COL)), clean_text(cell(row, ALT_BOM_COL)))
        if key[0] and key not in result:
            result[key] = (cell(row, STATUS_COL), cell(row, MESSAGE_COL),
                           cell(row, CREATED_ALT_BOM_COL), cell(row, UPDATED_AT_COL))
    return result


def build_worklist(path):
    boms = read_boms(path)          # first sheet = source
    problems = validate_boms(boms)
    previous = previous_results(path)
    stamps = []

    for bom in boms:
        status, message, created, updated = previous.get((bom.parent, bom.alt_bom), ("", "", "", ""))
        if status == "SUCCESS":
            # Already created: keep it, never create twice
            bom.status, bom.message, bom.created_alt_bom = "SUCCESS", message, created
            stamps.append(updated)
            continue
        bad = [p for p in problems if p.startswith(f"{bom.parent} alt {bom.alt_bom} ")]
        if bad:
            bom.status, bom.message = "FAILED", "; ".join(bad)
        elif bom.action != "CREATE":
            bom.status, bom.message = "SKIPPED", f"Action {bom.action}: not done by bom_create_main.py"
        else:
            bom.status, bom.message = "PENDING", ""
        stamps.append(None)
    return boms, stamps


def row_values(bom, updated_at=None):
    # Same values on every component row of the BOM
    values = {CREATED_ALT_BOM_COL: bom.created_alt_bom, STATUS_COL: bom.status,
              MESSAGE_COL: bom.message, UPDATED_AT_COL: updated_at or now()}
    return {item.excel_row: values for item in bom.items}


def save(path, boms, stamps):
    values = {}
    for i, bom in enumerate(boms):
        values.update(row_values(bom, stamps[i]))
    return update_source_columns(path, values, RESULT_COLUMNS)


def create_bom(session, bom):
    open_cs01(session)
    open_bom_cs01(session, material=bom.parent, bom_usage=bom.usage, plant=bom.plant)
    bom.created_alt_bom = get_created_alt_bom(session)
    enter_items(session, bom.items)
    press_enter(session)            # confirm item entries / warnings
    press_enter(session)
    time.sleep(2)
    save_bom(session)
    bom.status = "SUCCESS"
    bom.message = session.findById("wnd[0]/sbar").text
    if bom.created_alt_bom != bom.alt_bom:
        bom.message += f" WARNING: SAP alt BOM {bom.created_alt_bom}, Excel alt BOM {bom.alt_bom}."


def main():
    args = sys.argv[1:]
    path = CHAIN_FILE
    print(f"Phase 1 - create BOM, file: {path}")
    try:
        boms, stamps = build_worklist(path)
        todo = [i for i, b in enumerate(boms) if b.status == "PENDING"]
        counts = {s: sum(b.status == s for b in boms) for s in ("PENDING", "SUCCESS", "SKIPPED", "FAILED")}
        print(f"{len(boms)} BOMs: {counts}")
        for b in boms:
            if b.status == "FAILED":
                print(f"  FAILED {b.parent} alt {b.alt_bom}: {b.message}")
        if not save(path, boms, stamps) or "check" in args:
            return
        if not todo:
            print("Nothing to do.")
            return

        session = get_session(slow=slow_from_args())
        print(f"Connected to SAP system {session.Info.SystemName}")
        saved = True
        try:
            for i in todo:
                bom = boms[i]
                print("=" * 80)
                try:
                    create_bom(session, bom)
                except Exception as e:
                    bom.status, bom.message = "FAILED", str(e)
                print(f"{bom.parent} alt {bom.alt_bom}: {bom.status} {bom.message}")
                stamps[i] = now()
                saved = save(path, boms, stamps)    # after every BOM
                if not saved:
                    # never create more BOMs than are recorded: next run would create them twice
                    print(f"STOPPED. The result of {bom.parent} alt {bom.alt_bom} is NOT in Excel: "
                          f"it was {bom.status} in SAP - set its rows by hand before the next run.")
                    break
        finally:
            if saved and save(path, boms, stamps):
                print(f"Result saved to Sheet1 of {path}")

    except PermissionError:
        print(f"Cannot write {path}. Close it in Excel and run again.")
    except SAPConnectionError as e:
        print(f"SAP Connection Error: {e}")


if __name__ == "__main__":
    main()
