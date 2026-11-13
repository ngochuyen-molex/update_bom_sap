from bom import clean_component, set_text, press_enter
from models import Routing

# Popup after Enter: "Entry Option: Group" - all groups and GrC of the material
GROUP_POPUP_TABLE_ID = "wnd[1]/usr/tblSAPLCPSLTCTRL_2110"
POPUP_GROUP_COL = 0         # Group (PLNNR) = ctxt[0,row], only filled on the first row of each group
POPUP_COUNTER_COL = 1       # GrC (PLNAL) = txt[1,row]
POPUP_DESC_COL = 2          # Task list description = txt[2,row]

# Change Routing: Overview -> header overview (GrC of the selected group only)
HEADER_TABLE_ID = "wnd[0]/usr/tblSAPLCPDITCTRL_1200"
HEADER_COUNTER_COL = 0      # GrC = txt[0,row]
HEADER_DESC_COL = 1         # Task list description = txt[1,row]

# Component allocation table
COMP_TABLE_ID = "wnd[0]/usr/tblSAPLCMDITCTRL_1000"
COMP_OPERATION_COL = 10     # txt[10,row] = assigned operation ("0030", ...)
COMP_DESC_COL = 13          # txt[13,row] = material description

# "Choose Alternative BOM" popup after btn[7]: list, element = txt[<row>, <column position>]
BOM_POPUP_ID = "wnd[1]/usr/sub/1"
BOM_POPUP_ALT_COL = 19      # txt[row,19] = AltBOM

# Operation per component (first match wins):
#   packaging                                         -> 0030
#   lubricant AND op 0030 work center in YMHI list    -> 0030
#   everything else                                   -> 0020
DEFAULT_OPERATION = "0020"
PACKAGING_OPERATION = "0030"
PACKAGING_WORDS = ("carton", "sheet", "bag", "silica", "bubble", "box")
LUBRICANT_WORDS = ("lubricant",)
LUBRICANT_0030_WORK_CENTERS = ("YMHI", "YMHI-1")     # exact match

# ============================================================
# RULES (one place each, so a rule change is a one-line edit)
# ============================================================
def is_trial(text):
    # Trial routing (description) / trial alt BOM (alt BOM text), covers MDTRIAL
    return "TRIAL" in (text or "").upper()

def is_packaging(description):
    # Case-insensitive "contains", so plurals match too (box -> boxes, sheet -> sheets)
    text = (description or "").lower()
    return any(word in text for word in PACKAGING_WORDS)

def is_lubricant(description):
    text = (description or "").lower()
    return any(word in text for word in LUBRICANT_WORDS)

def is_lubricant_0030_routing(op_0030_wc):
    return (op_0030_wc or "").strip().upper() in LUBRICANT_0030_WORK_CENTERS

def planned_operation(description, op_0030_wc):
    # Operation for an unassigned component; returns (operation, reason)
    if is_packaging(description):
        return PACKAGING_OPERATION, "packaging"
    if is_lubricant(description) and is_lubricant_0030_routing(op_0030_wc):
        return PACKAGING_OPERATION, f"lubricant, op 0030 = {op_0030_wc}"
    return DEFAULT_OPERATION, ""

def skip_reason(routing: Routing):
    # Why this routing x alt BOM is not opened in SAP ("" = process it)
    if routing.deleted:
        return "Deleted routing"
    if is_trial(routing.description):
        return "Trial routing"
    if is_trial(routing.alt_bom_text):
        return "Trial alt BOM"
    if routing.total_items == 0:
        return "No BOM items (Snowflake)"
    if routing.unassigned == 0:
        return f"Already assigned (Snowflake): {routing.total_items} of {routing.total_items} components assigned"
    return ""

# ============================================================
# BASIC HELPERS
# ============================================================
def get_cell_text(session, table_id, col, row):
    # Cells are txt or ctxt depending on the field; try both
    for kind in ("txt", "ctxt"):
        cell = session.findById(f"{table_id}/{kind}[{col},{row}]", False)
        if cell is not None:
            return cell.text.strip()
    return ""

def clean_counter(value):
    # "022" / "22" / 22.0 -> "22"
    return clean_component(value) or "0"

def get_status_text(session):
    return session.findById("wnd[0]/sbar").text.strip()

def find_row(session, table_id, col, target, normalize=str.strip):
    # Pages through the table and returns the VISIBLE row index whose column
    # `col` equals `target`. The table stays scrolled to that page.
    # Stops at the first fully blank row. None if not found.
    # SAP re-renders after scrolling, so the table is always re-fetched.
    target = normalize(target)
    table = session.findById(table_id)
    total_rows = int(table.RowCount)
    visible_rows = int(table.VisibleRowCount)

    position = 0
    while position < total_rows:
        session.findById(table_id).VerticalScrollbar.Position = position
        first_row = int(session.findById(table_id).VerticalScrollbar.Position)

        for row_idx in range(min(visible_rows, total_rows - first_row)):
            value = get_cell_text(session, table_id, col, row_idx)
            if value == "" and get_cell_text(session, table_id, col + 1, row_idx) == "":
                return None     # blank row -> end of list
            if value and normalize(value) == target:
                return row_idx

        if first_row + visible_rows >= total_rows:
            break
        position = first_row + visible_rows
    return None

def double_click(session, window, table_id, col, row_idx):
    session.findById(f"{table_id}/txt[{col},{row_idx}]").setFocus()
    session.findById(window).sendVKey(2)

# ============================================================
# NAVIGATION
# ============================================================
def open_ca02(session):
    session.findById("wnd[0]/tbar[0]/okcd").text = "/NCA02"
    press_enter(session)

def open_routing(session, material, plant="1901"):
    print(f"Material={clean_component(material)}, Plant={plant}")
    set_text(session, "wnd[0]/usr/ctxt[0]", clean_component(material))
    set_text(session, "wnd[0]/usr/ctxt[1]", plant)
    press_enter(session)

def has_group_popup(session):
    return session.findById(GROUP_POPUP_TABLE_ID, False) is not None

def has_header_overview(session):
    return session.findById(HEADER_TABLE_ID, False) is not None

def select_group(session, group):
    # Group is only on the first row of each group; that row opens the group
    row_idx = find_row(session, GROUP_POPUP_TABLE_ID, POPUP_GROUP_COL, group)
    if row_idx is None:
        raise ValueError(f"Group {group} not found in group popup")
    double_click(session, "wnd[1]", GROUP_POPUP_TABLE_ID, POPUP_DESC_COL, row_idx)

def open_header_row(session, counter):
    row_idx = find_row(session, HEADER_TABLE_ID, HEADER_COUNTER_COL, counter,
                       normalize=clean_counter)
    if row_idx is None:
        raise ValueError(f"GrC {counter} not found in header overview")
    description = get_cell_text(session, HEADER_TABLE_ID, HEADER_DESC_COL, row_idx)
    print(f"Opening GrC {counter} (row {row_idx}): {description}")
    double_click(session, "wnd[0]", HEADER_TABLE_ID, HEADER_DESC_COL, row_idx)

# ============================================================
# COMPONENT ALLOCATION
# ============================================================
def open_component_allocation(session, routing: Routing):
    # btn[7] -> "Choose Alternative BOM" popup -> double-click the row with the wanted alt BOM.
    # Rows are found by value (AltBOM column), not by position.
    session.findById("wnd[0]/tbar[1]/btn[7]").press()

    if session.findById("wnd[1]", False) is None:
        # No popup: SAP opened the only BOM directly
        if routing.bom_count > 1:
            raise RuntimeError(
                f"BOM popup not shown, but material has {routing.bom_count} active alt BOMs. "
                f"SAP status: '{get_status_text(session)}'")
        print("No BOM popup, single BOM opened directly")
        return

    wanted = clean_counter(routing.alt_bom)
    row = 0
    while True:
        cell = session.findById(f"{BOM_POPUP_ID}/txt[{row},{BOM_POPUP_ALT_COL}]", False)
        if cell is None:
            break
        if clean_counter(cell.text) == wanted:
            print(f"BOM popup: alt BOM {wanted} on row {row}")
            cell.setFocus()
            session.findById("wnd[1]").sendVKey(2)
            return
        row += 1
    raise ValueError(f"Alt BOM {wanted} not found in BOM popup ({row} rows)")

def is_blank_comp_row(session, row_idx):
    # Blank = no text in any cell before the operation column
    return all(get_cell_text(session, COMP_TABLE_ID, col, row_idx) == ""
               for col in range(COMP_OPERATION_COL))

def select_components(session, operation, op_0030_wc):
    # Selects visible rows with an empty operation (txt[10]) whose planned operation
    # is `operation` (pass 1: 0030). For the default operation (pass 2: 0020) every
    # remaining unassigned row is selected. Stops at the first blank row. No scrolling.
    table = session.findById(COMP_TABLE_ID)
    visible_rows = min(int(table.VisibleRowCount), int(table.RowCount))
    selected = []

    for row_idx in range(visible_rows):
        if is_blank_comp_row(session, row_idx):
            break
        description = get_cell_text(session, COMP_TABLE_ID, COMP_DESC_COL, row_idx)
        unassigned = get_cell_text(session, COMP_TABLE_ID, COMP_OPERATION_COL, row_idx) == ""
        planned, reason = planned_operation(description, op_0030_wc)
        select = unassigned and (operation == DEFAULT_OPERATION or planned == operation)

        # Tick rows to assign; untick only rows that are still ticked (e.g. from pass 1).
        # Rows that are already unticked are left alone (no extra SAP actions).
        table_row = session.findById(COMP_TABLE_ID).getAbsoluteRow(row_idx)
        if select or table_row.selected:
            table_row.selected = select
        if select:
            selected.append((row_idx, reason))
            print(f"  select row {row_idx}: {description}" + (f" ({reason})" if reason else ""))
    else:
        if int(table.RowCount) > visible_rows:
            print(f"WARNING: table has more rows than visible ({visible_rows}); "
                  "rows below were not checked")

    return selected

def list_windows(session):
    # Titles of all open windows of the session, e.g. ["wnd[0]: Change Routing...", "wnd[1]: ..."]
    return [f"{w.Id.split('/')[-1]}: {w.Text}" for w in session.Children]

def assign_to_operation(session, operation, last_row):
    # Selected rows -> btn[5] -> type operation in popup -> Enter
    session.findById(f"{COMP_TABLE_ID}/chk[0,{last_row}]").setFocus()
    session.findById("wnd[0]/tbar[1]/btn[5]").press()

    # Diagnostics: show what SAP did after btn[5]
    print(f"  after btn[5]: status='{get_status_text(session)}', windows={list_windows(session)}")

    field = session.findById("wnd[1]/usr/txt[0]", False)
    if field is None:
        raise RuntimeError(
            f"btn[5] did not open the operation popup. SAP status: '{get_status_text(session)}'")
    field.text = operation
    session.findById("wnd[1]").sendVKey(0)      # Enter in the popup
    print(f"Assigned to {operation}. SAP status: '{get_status_text(session)}'")

def assign_components(session, op_0030_wc=""):
    # Pass 1: planned 0030 (packaging, lubricant on YMHI routings).
    # Pass 2: everything still unassigned -> 0020.
    # Returns {operation: [(row, reason), ...]} for the rows that were assigned.
    assigned = {}
    for operation in (PACKAGING_OPERATION, DEFAULT_OPERATION):
        print(f"Selecting rows for {operation}")
        rows = select_components(session, operation, op_0030_wc)
        if rows:
            assign_to_operation(session, operation, rows[-1][0])
            assigned[operation] = rows
    return assigned

def summarize(assigned):
    # {"0030": [(r, "packaging"), (r, "lubricant, ...")], "0020": [...]} -> "0030: 2 rows (1 lubricant), 0020: 5 rows"
    parts = []
    for operation, rows in assigned.items():
        lubricant = sum(1 for _, reason in rows if reason.startswith("lubricant"))
        extra = f" ({lubricant} lubricant)" if lubricant else ""
        parts.append(f"{operation}: {len(rows)} rows{extra}")
    return ", ".join(parts)

def save_routing(session):
    session.findById("wnd[0]/tbar[0]/btn[11]").press()

# ============================================================
# MAIN FLOW: one routing x alt BOM = one CA02 entry, saved at the end
# ============================================================
def process_routing(session, routing: Routing):
    print("-" * 60)
    print(f"Routing {routing.key}: {routing.description} | op 0030 WC: {routing.op_0030_wc or '-'}")

    open_ca02(session)
    open_routing(session, routing.material, routing.plant)
    if has_group_popup(session):
        select_group(session, routing.group)
    # No popup = material has only one group, SAP opens the header overview directly

    note = ""
    if has_header_overview(session):
        open_header_row(session, routing.counter)
    else:
        # SAP opened the only routing of the group directly (operation overview).
        # The GrC can't be checked on screen, so only continue if one routing is expected.
        if routing.group_size > 1:
            raise RuntimeError(
                f"Header overview not shown, but group {routing.group} has "
                f"{routing.group_size} routings. SAP status: '{get_status_text(session)}'")
        note = "Single routing, no header overview. "
        print(note.strip())

    open_component_allocation(session, routing)

    assigned = assign_components(session, routing.op_0030_wc)
    save_routing(session)       # also when nothing was assigned
    if not assigned:
        routing.status = "SKIPPED"
        routing.message = f"{note}No unassigned components on screen, saved. {get_status_text(session)}"
        print(routing.message)
        return

    routing.status = "SUCCESS"
    routing.message = f"{note}{summarize(assigned)}. {get_status_text(session)}"
    print(f"Saved: {routing.message}")


if __name__ == "__main__":
    # Test one routing x alt BOM: python routing.py [slow]
    from connect import get_session, slow_from_args
    session = get_session(slow=slow_from_args())
    test = Routing(material="5013301300", group="M5013301", counter="45", alt_bom="2", bom_count=2)
    try:
        process_routing(session, test)
    except Exception as e:
        test.status, test.message = "FAILED", str(e)
    print(f"{test.key}: {test.status} {test.message}")
