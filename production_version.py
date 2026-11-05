from datetime import date
import re
import pandas as pd
from connect import get_session, slow_from_args, SAPConnectionError
from excel_utils import load_excel
from bom import clean_component, clean_text, set_text, press_enter
from models import ProductionVersion
from routing import get_cell_text, get_status_text
from routing_sf import load_routings

# Result file of create_bom.py: only materials with a newly created BOM
INPUT_FILE = r"C:\Users\nguyeb3\Downloads\create_bom_final.xlsx"
RESULT_FILE = INPUT_FILE.replace(".xlsx", "_prod_version.xlsx")

# Production version list popup (MM02 -> MRP 4 -> production versions)
VERSION_TABLE_ID = "wnd[1]/usr/tblSAPLMDIAT_MKAL_NEW"
VERSION_COL = 0             # ctxt[0,row] = version, e.g. "0001"
MAX_REAL_VERSION = 5000     # ignore special versions like 9998 / 9999

VALID_TO = "9999/12/31"
DATE_FORMAT = "%Y/%m/%d"    # same as the recording, e.g. 2026/11/03
MAX_TEXT_LENGTH = 40        # production version text field length

def is_trial(description):
    # Trial routings get no production version (rule may change: keep it in this one place)
    return "TRIAL" in description.upper()

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


def get_alt_boms(df):
    # {material: alt BOM} for materials with STATUS = SUCCESS.
    # More than one alt BOM per material -> None (ambiguous, reported as FAILED)
    alt_boms = {}
    for _, row in df.iterrows():
        if str(row.get("STATUS")).strip() != "SUCCESS":
            continue
        material = clean_component(row["Parent Material"])
        alt_bom = clean_component(row["Alternative BOM"])
        if not material or not alt_bom:
            continue
        alt_boms.setdefault(material, set()).add(alt_bom)
    return {m: (boms.pop() if len(boms) == 1 else None) for m, boms in alt_boms.items()}


def save_results(versions, missing):
    rows = [{
        "Material": pv.material,
        "Group": pv.group,
        "GrC": pv.counter,
        "Description": pv.description,
        "Alt BOM": pv.alt_bom,
        "Version Text": pv.text,
        "New Version": pv.version,
        "STATUS": pv.status,
        "MESSAGE": pv.message,
    } for pv in versions]
    rows += [{"Material": m, "STATUS": "FAILED", "MESSAGE": msg} for m, msg in missing]
    try:
        pd.DataFrame(rows).to_excel(RESULT_FILE, index=False)
    except PermissionError:
        print(f"Could not save {RESULT_FILE}. Close it in Excel.")


def main():
    try:
        session = get_session(slow=slow_from_args())    # python production_version.py slow

        alt_boms = get_alt_boms(load_excel(INPUT_FILE))
        missing = [(m, "More than one Alternative BOM in Excel")
                   for m, bom in alt_boms.items() if bom is None]
        materials = [m for m, bom in alt_boms.items() if bom is not None]

        routings = load_routings(materials)
        found = {r.material for r in routings}
        missing += [(m, "No routing found in MLX_DATALAKE") for m in materials if m not in found]

        versions = [ProductionVersion(
            material=r.material, group=r.group, counter=r.counter,
            description=r.description, alt_bom=alt_boms[r.material])
            for r in routings]
        
        print(f"{len(materials)} materials -> {len(versions)} production versions")
        open_mm02(session)
        for pv in versions:
            try:
                process_production_version(session, pv)
            except Exception as e:
                pv.status = "FAILED"
                pv.message = str(e)
            print(f"{pv.key}: {pv.status} {pv.message}")
            save_results(versions, missing)

    except SAPConnectionError as e:
        print(f"SAP Connection Error: {e}")
    except Exception as e:
        print(f"Unexpected Error: {e}")


if __name__ == "__main__":
    import sys
    if "test" in sys.argv:
        # python production_version.py test [slow] -> one version, no Excel / Snowflake
        session = get_session(slow=slow_from_args())
        pv = ProductionVersion(material="5013301000", group="M501334", counter="5",
                               description="sample description", alt_bom="2", bom_usage="1")
        try:
            open_mm02(session)
            process_production_version(session, pv)
        except Exception as e:
            pv.status, pv.message = "FAILED", str(e)
        print(f"{pv.key}: {pv.status} {pv.message}")
    else:
        main()
