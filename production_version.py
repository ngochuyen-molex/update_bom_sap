"""Phase 3 - create production versions in MM02 (MRP 4 -> production versions).

Usage:
    python production_version.py [check] [slow]      routing list from sheet "Routing Result"
    python production_version.py continue [slow]     carry on with "Prod Version Result" (retry works too)
    python production_version.py fresh               rebuild from "Routing Result", keeps SUCCESS rows
    python production_version.py test [slow]         one hard-coded version, no Excel

One production version per routing x new alt BOM of phase 2; no Snowflake in this phase.
Deleted / trial routings are SKIPPED. A production version that is SUCCESS is never
created again. Text = routing description + " B" + alt BOM (max 40 characters).
"""
import sys
from datetime import date
import re
from connect import get_session, slow_from_args, SAPConnectionError
from bom import set_text, press_enter
from models import ProductionVersion
from routing import get_cell_text, get_status_text, is_trial
from utils import clean_component
from workbook import (CHAIN_FILE, ROUTING_SHEET, PROD_VERSION_SHEET, CONTINUE_STATUSES,
                      has_sheet, read_sheet, save_sheet, now, cell)

RESULT_COLUMNS = ["Material", "Group", "GrC", "Description", "Alt BOM", "Version Text",
                  "New Version", "STATUS", "MESSAGE", "Updated At"]

# Production version list popup (MM02 -> MRP 4 -> production versions)
VERSION_TABLE_ID = "wnd[1]/usr/tblSAPLMDIAT_MKAL_NEW"
VERSION_COL = 0             # ctxt[0,row] = version, e.g. "0001"
MAX_REAL_VERSION = 5000     # ignore special versions like 9998 / 9999

VALID_TO = "9999/12/31"
DATE_FORMAT = "%Y/%m/%d"    # same as the recording, e.g. 2026/11/03
MAX_TEXT_LENGTH = 40        # production version text field length

def version_text(description, alt_bom, material=""):
    # Production version text = routing description + " B" + alt BOM, e.g. "MF50-08 - ... B2".
    # Longer than 40 -> remove " - <material>" once, e.g.
    #   "MF15-13 - 5019400308 - L501940E 03 #03" -> "MF15-13 - L501940E 03 #03"
    # Still too long -> error, nothing is changed in SAP.
    description = description.strip()
    text = f"{description} B{alt_bom}"
    if len(text) > MAX_TEXT_LENGTH and material and material in description:
        short = re.sub(rf"\s*-\s*{re.escape(material)}", "", description, count=1).strip()
        text = f"{short} B{alt_bom}"
    if len(text) > MAX_TEXT_LENGTH:
        raise ValueError(f"Description longer than {MAX_TEXT_LENGTH} ({len(text)}): '{text}'")
    return text

# ============================================================
# NAVIGATION (MM02)
# ============================================================
def open_mm02(session):
    session.findById("wnd[0]/tbar[0]/okcd").text = "/NMM02"
    press_enter(session)
    session.findById("wnd[0]").maximize()

def open_production_versions(session, material):
    print(f"MM02 material={clean_component(material)}")
    set_text(session, "wnd[0]/usr/ctxt[0]", clean_component(material))
    press_enter(session)
    # Recording: Enter on two popups (select views, then org levels - plant defaults to 1901)
    for step in ("select views", "org levels"):
        popup = session.findById("wnd[1]", False)
        if popup is None:
            raise RuntimeError(f"Expected popup '{step}' not shown. SAP status: '{get_status_text(session)}'")
        popup.sendVKey(0)
    session.findById("wnd[0]/usr/sub/6/btn").press()
    if session.findById(VERSION_TABLE_ID, False) is None:
        raise RuntimeError(f"Production version list not shown. SAP status: '{get_status_text(session)}'")

# ============================================================
# VERSION LIST
# ============================================================
def read_versions(session):
    # All versions in the list (scrolls if longer than the screen), stops at the first blank row
    table = session.findById(VERSION_TABLE_ID)
    total_rows = int(table.RowCount)
    visible_rows = int(table.VisibleRowCount)
    versions = []

    position = 0
    while position < total_rows:
        session.findById(VERSION_TABLE_ID).VerticalScrollbar.Position = position
        first_row = int(session.findById(VERSION_TABLE_ID).VerticalScrollbar.Position)

        for row_idx in range(min(visible_rows, total_rows - first_row)):
            value = get_cell_text(session, VERSION_TABLE_ID, VERSION_COL, row_idx)
            if value == "":
                return versions
            versions.append(value)

        if first_row + visible_rows >= total_rows:
            break
        position = first_row + visible_rows
    return versions

def next_version(versions):
    # Max of the real versions (< 5000) + 1, 4 digits. No versions yet -> "0001"
    numbers = [int(v) for v in versions if v.isdigit() and int(v) < MAX_REAL_VERSION]
    return f"{max(numbers, default=0) + 1:04d}"

# ============================================================
# NEW VERSION DETAILS (wnd[2])
# ============================================================
def enter_new_version(session, pv: ProductionVersion):
    session.findById("wnd[1]/tbar[0]/btn[7]").press()     # New details
    if session.findById("wnd[2]", False) is None:
        raise RuntimeError(f"New details popup not shown. SAP status: '{get_status_text(session)}'")

    today = date.today().strftime(DATE_FORMAT)
    print(f"New version {pv.version}: '{pv.text}', {today} - {VALID_TO}, "
          f"routing {pv.group}/{pv.counter}, alt BOM {pv.alt_bom}, usage {pv.bom_usage}")

    set_text(session, "wnd[2]/usr/ctxt[0]", pv.version)
    set_text(session, "wnd[2]/usr/txt[0]", pv.text)
    set_text(session, "wnd[2]/usr/ctxt[1]", today)
    set_text(session, "wnd[2]/usr/ctxt[2]", VALID_TO)
    session.findById("wnd[2]/usr/cmb[1]").key = "N"
    set_text(session, "wnd[2]/usr/ctxt[3]", pv.group)
    set_text(session, "wnd[2]/usr/txt[4]", pv.counter)
    set_text(session, "wnd[2]/usr/ctxt[6]", pv.alt_bom)
    set_text(session, "wnd[2]/usr/ctxt[7]", pv.bom_usage)

    session.findById("wnd[2]/usr/btn").press()            # Check (validates group / GrC)

    warning = ""
    if session.findById("wnd[3]", False) is not None:
        # Warning popup: group / group counter not valid -> log it and close the popup
        warning = read_popup_text(session, "wnd[3]") or "Check warning (no text)"
        print(f"WARNING: {warning}")
        session.findById("wnd[3]/tbar[0]/btn[0]").press()

    session.findById("wnd[2]/tbar[0]/btn[12]").press()    # red cross
    session.findById("wnd[2]/tbar[0]/btn[0]").press()     # continue
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    print(f"Version entered. SAP status: '{get_status_text(session)}'")
    return warning

def save_material(session):
    # Same Save button as BOM / routing
    session.findById("wnd[0]/tbar[0]/btn[11]").press()
    if session.findById("wnd[1]", False) is not None:
        # Unknown popup after Save: stop this routing and show its text
        raise RuntimeError(f"Popup after Save: {read_popup_text(session, 'wnd[1]')}")
    return get_status_text(session)

def read_popup_text(session, window):
    # Joins all visible texts of a popup (title + labels / message lines)
    texts = []
    popup = session.findById(window)
    if popup.Text.strip():
        texts.append(popup.Text.strip())
    for element in session.findById(f"{window}/usr").Children:
        text = getattr(element, "Text", "").strip()
        if text:
            texts.append(text)
    return " | ".join(texts)

# ============================================================
# MAIN FLOW: one routing = one new production version
# ============================================================
def process_production_version(session, pv: ProductionVersion):
    print("-" * 60)
    print(f"Production version for {pv.key}: {pv.description}")
    if is_trial(pv.description):
        pv.status, pv.message = "SKIPPED", "Trial routing, no production version"
        return
    pv.text = version_text(pv.description, pv.alt_bom, pv.material)
    open_production_versions(session, pv.material)

    versions = read_versions(session)
    pv.version = next_version(versions)
    print(f"Existing versions: {versions} -> new {pv.version}")

    warning = enter_new_version(session, pv)
    status_text = save_material(session)
    pv.status = "SUCCESS"
    pv.message = f"Version {pv.version} saved. {status_text}"
    if warning:
        pv.message += f". Check warning: {warning}"
    print(f"Saved: {pv.message}")


def build_worklist(path):
    """Production versions from the "Routing Result" sheet (phase 2) - no Snowflake.

    One production version per routing x new alt BOM. Rows that were not assigned in
    phase 2 are not created yet: deleted / trial routings are SKIPPED with the same reason,
    routings still PENDING / FAILED in phase 2 are SKIPPED "Waiting for ..." and picked up
    by "fresh" after phase 2 is done.
    (Future standalone production-version task: get the routing list from Snowflake with
    routing_sf.load_routings() instead - see docs/routing.md.)
    """
    if not has_sheet(path, ROUTING_SHEET):
        raise ValueError(f"Sheet '{ROUTING_SHEET}' not found in {path}: run routing_main.py first")
    versions = []
    for _, row in read_sheet(path, ROUTING_SHEET).iterrows():
        pv = ProductionVersion(
            material=cell(row, "Material"), group=cell(row, "Group"), counter=cell(row, "GrC"),
            description=cell(row, "Description"), alt_bom=cell(row, "Alt BOM"))
        routing_status, routing_message = cell(row, "STATUS"), cell(row, "MESSAGE")
        if not pv.group:
            pv.status, pv.message = "SKIPPED", routing_message or "No routing"
        elif routing_message.startswith("Deleted routing"):
            pv.status, pv.message = "SKIPPED", "Deleted routing"
        elif is_trial(pv.description):
            pv.status, pv.message = "SKIPPED", "Trial routing, no production version"
        elif routing_status == "SUCCESS" or routing_message.startswith("No unassigned components on screen"):
            pv.status, pv.message = "PENDING", ""
        else:
            pv.status, pv.message = "SKIPPED", f"Waiting for routing assignment (STATUS {routing_status})"
        versions.append(pv)
    return versions


def versions_from_result(df):
    # Continue: rebuild from the "Prod Version Result" sheet (no Snowflake, no Routing Result)
    versions, updated_at = [], []
    for _, row in df.iterrows():
        pv = ProductionVersion(
            material=cell(row, "Material"), group=cell(row, "Group"), counter=cell(row, "GrC"),
            description=cell(row, "Description"), alt_bom=cell(row, "Alt BOM"))
        pv.text, pv.version = cell(row, "Version Text"), cell(row, "New Version")
        pv.status, pv.message = cell(row, "STATUS"), cell(row, "MESSAGE")
        versions.append(pv)
        updated_at.append(cell(row, "Updated At"))
    return versions, updated_at


def keep_previous_success(versions, path):
    # "fresh": versions that were SUCCESS stay SUCCESS - never create a production version twice
    if not has_sheet(path, PROD_VERSION_SHEET):
        return [None] * len(versions)
    old, old_stamps = versions_from_result(read_sheet(path, PROD_VERSION_SHEET))
    done = {(p.material, p.group, p.counter, p.alt_bom): (p, old_stamps[i])
            for i, p in enumerate(old) if p.status == "SUCCESS"}
    stamps = []
    for pv in versions:
        prev = done.get((pv.material, pv.group, pv.counter, pv.alt_bom))
        if prev:
            pv.status, pv.message, pv.text, pv.version = "SUCCESS", prev[0].message, prev[0].text, prev[0].version
            stamps.append(prev[1])
        else:
            stamps.append(None)
    return stamps


def result_row(pv, updated_at=None):
    return {
        "Material": pv.material, "Group": pv.group, "GrC": pv.counter,
        "Description": pv.description, "Alt BOM": pv.alt_bom, "Version Text": pv.text,
        "New Version": pv.version, "STATUS": pv.status, "MESSAGE": pv.message,
        "Updated At": updated_at or now(),
    }


def run(session, path, versions, updated_at):
    rows = [result_row(pv, updated_at[i]) for i, pv in enumerate(versions)]
    todo = [i for i, pv in enumerate(versions) if pv.status in CONTINUE_STATUSES and pv.group]
    print(f"{len(versions)} rows, {len(todo)} production versions to create")
    if not save_sheet(path, PROD_VERSION_SHEET, rows, RESULT_COLUMNS):
        return
    if not todo:
        return
    saved = True
    try:
        open_mm02(session)
        for i in todo:
            pv = versions[i]
            try:
                process_production_version(session, pv)
            except Exception as e:
                pv.status, pv.message = "FAILED", str(e)
                try:
                    open_mm02(session)      # start the next row from a clean MM02 screen
                except Exception:
                    pass
            print(f"{pv.key}: {pv.status} {pv.message}")
            rows[i] = result_row(pv)
            saved = save_sheet(path, PROD_VERSION_SHEET, rows, RESULT_COLUMNS)    # after every version
            if not saved:
                # never create more versions than are recorded: next run would create them twice
                print(f"STOPPED. The result of {pv.key} is NOT in Excel: it was {pv.status} "
                      f"(version {pv.version}) - set its row by hand before the next run.")
                break
    finally:
        if saved and save_sheet(path, PROD_VERSION_SHEET, rows, RESULT_COLUMNS):
            print(f"Result saved to sheet '{PROD_VERSION_SHEET}' of {path}")


def main():
    args = sys.argv[1:]
    path = CHAIN_FILE
    print(f"Phase 3 - production version, file: {path}")
    try:
        if "continue" in args or "retry" in args:
            versions, updated_at = versions_from_result(read_sheet(path, PROD_VERSION_SHEET))
            todo = sum(pv.status in CONTINUE_STATUSES and bool(pv.group) for pv in versions)
            print(f"Loaded {len(versions)} rows from sheet '{PROD_VERSION_SHEET}', {todo} PENDING / FAILED")
            if todo == 0:
                print("Nothing to do.")
                return
            run(get_session(slow=slow_from_args()), path, versions, updated_at)
            return

        if has_sheet(path, PROD_VERSION_SHEET) and "fresh" not in args:
            print(f"Sheet '{PROD_VERSION_SHEET}' already exists in {path}.\n"
                  f"  To carry on with it:  python production_version.py continue\n"
                  f"  To rebuild the list:  python production_version.py fresh  "
                  f"(from '{ROUTING_SHEET}', keeps SUCCESS rows)")
            return

        versions = build_worklist(path)
        updated_at = keep_previous_success(versions, path)
        counts = {s: sum(pv.status == s for pv in versions) for s in ("PENDING", "SUCCESS", "SKIPPED", "FAILED")}
        print(f"{len(versions)} rows from '{ROUTING_SHEET}': {counts}")
        if "check" in args:
            save_sheet(path, PROD_VERSION_SHEET,
                       [result_row(pv, updated_at[i]) for i, pv in enumerate(versions)], RESULT_COLUMNS)
            print(f"Written to sheet '{PROD_VERSION_SHEET}' of {path}")
            return
        if counts["PENDING"] == 0:
            save_sheet(path, PROD_VERSION_SHEET,
                       [result_row(pv, updated_at[i]) for i, pv in enumerate(versions)], RESULT_COLUMNS)
            print("Nothing to do.")
            return
        run(get_session(slow=slow_from_args()), path, versions, updated_at)

    except PermissionError:
        print(f"Cannot write {path}. Close it in Excel and run again.")
    except SAPConnectionError as e:
        print(f"SAP Connection Error: {e}")


def test():
    # python production_version.py test [slow] -> one hard-coded version, no Excel / Snowflake
    pv = ProductionVersion(material="5013301000", group="M501334", counter="5",
                           description="sample description", alt_bom="2", bom_usage="1")
    try:
        session = get_session(slow=slow_from_args())
        open_mm02(session)
        process_production_version(session, pv)
    except SAPConnectionError as e:
        print(f"SAP Connection Error: {e}")
        return
    except Exception as e:
        pv.status, pv.message = "FAILED", str(e)
    print(f"{pv.key}: {pv.status} {pv.message}")


if __name__ == "__main__":
    test() if "test" in sys.argv else main()
