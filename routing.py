from bom import clean_component, set_text, press_enter
from models import RoutingGroup, RoutingTask
import time

# TEMPORARY: wait after every SAP action so the run can be observed. 0 = normal speed.
STEP_DELAY = 2

class SlowSAP:
    # Wraps a SAP GUI object: sleeps STEP_DELAY after actions, reads are not delayed
    ACTIONS = {"press", "sendVKey", "setFocus", "select"}
    SET_ACTIONS = {"text", "selected", "Position", "caretPosition"}

    def __init__(self, obj):
        object.__setattr__(self, "_obj", obj)

    def __getattr__(self, name):
        value = getattr(self._obj, name)
        # COM objects (Info, VerticalScrollbar, ...) are callable too, so check them first
        if hasattr(value, "_oleobj_"):
            return SlowSAP(value)
        if callable(value):
            def call(*args):
                result = value(*args)
                if name in self.ACTIONS:
                    print(f"  [slow] {name}{args}")
                    time.sleep(STEP_DELAY)
                return wrap(result)
            return call
        return value

    def __setattr__(self, name, value):
        setattr(self._obj, name, value)
        if name in self.SET_ACTIONS:
            print(f"  [slow] {name} = {value!r}")
            time.sleep(STEP_DELAY)

def wrap(value):
    # Wrap SAP COM objects (findById results, rows, scrollbars); leave str/int/None as is
    if hasattr(value, "_oleobj_"):
        return SlowSAP(value)
    return value

def slow_session(session):
    return SlowSAP(session) if STEP_DELAY > 0 else session

# Popup after Enter: "Entry Option: Group" - all groups and GrC of the material
GROUP_POPUP_TABLE_ID = "wnd[1]/usr/tblSAPLCPSLTCTRL_2110"
POPUP_GROUP_COL = 0         # Group (PLNNR) = ctxt[0,row], only filled on the first row of each group
POPUP_COUNTER_COL = 1       # GrC (PLNAL) = txt[1,row]
POPUP_DESC_COL = 2          # Task list description (the recording double-clicks this)

# Change Routing: Overview -> header overview (GrC of the selected group only)
HEADER_TABLE_ID = "wnd[0]/usr/tblSAPLCPDITCTRL_1200"
HEADER_COUNTER_COL = 0      # GrC = txt[0,row]
HEADER_DESC_COL = 1         # Task list description

# Component allocation table
COMP_TABLE_ID = "wnd[0]/usr/tblSAPLCMDITCTRL_1000"
COMP_OPERATION_COL = 10     # txt[10,row] = assigned operation ("0030", ...)

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

def scroll_to(session, table_id, absolute_row):
    # Scroll so absolute_row is visible, return its visible row index.
    # SAP re-renders after scrolling, so always re-fetch the table.
    table = session.findById(table_id)
    if absolute_row < int(table.VisibleRowCount) and int(table.VerticalScrollbar.Position) == 0:
        return absolute_row
    session.findById(table_id).VerticalScrollbar.Position = absolute_row
    first_row = int(session.findById(table_id).VerticalScrollbar.Position)
    return absolute_row - first_row

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

# ============================================================
# GROUP POPUP ("Entry Option: Group")
# ============================================================
def read_group_popup(session):
    # Reads all rows of the popup into RoutingGroup objects.
    # Group is only shown on the first row of each group, so blank = same as above.
    table = session.findById(GROUP_POPUP_TABLE_ID)
    total_rows = int(table.RowCount)
    visible_rows = int(table.VisibleRowCount)

    groups = []
    absolute_row = 0
    while absolute_row < total_rows:
        session.findById(GROUP_POPUP_TABLE_ID).VerticalScrollbar.Position = absolute_row
        first_row = int(session.findById(GROUP_POPUP_TABLE_ID).VerticalScrollbar.Position)

        for row_idx in range(min(visible_rows, total_rows - first_row)):
            group = get_cell_text(session, GROUP_POPUP_TABLE_ID, POPUP_GROUP_COL, row_idx)
            counter = get_cell_text(session, GROUP_POPUP_TABLE_ID, POPUP_COUNTER_COL, row_idx)
            if group == "" and counter == "":
                total_rows = 0      # blank row -> end of list
                break
            if group:
                groups.append(RoutingGroup(group=group, popup_row=first_row + row_idx))
            if not groups:
                raise ValueError("First popup row has no group number")
            groups[-1].counters.append(clean_counter(counter))

        absolute_row = first_row + visible_rows

    session.findById(GROUP_POPUP_TABLE_ID).VerticalScrollbar.Position = 0
    return groups

def select_popup_row(session, absolute_row):
    row_idx = scroll_to(session, GROUP_POPUP_TABLE_ID, absolute_row)
    session.findById(f"{GROUP_POPUP_TABLE_ID}/txt[{POPUP_DESC_COL},{row_idx}]").setFocus()
    session.findById("wnd[1]").sendVKey(2)

# ============================================================
# HEADER OVERVIEW (routings of the selected group)
# ============================================================
def get_header_counters(session):
    # {GrC: visible row index}, stops at the first blank row
    table = session.findById(HEADER_TABLE_ID)
    counters = {}
    for row_idx in range(int(table.VisibleRowCount)):
        counter = get_cell_text(session, HEADER_TABLE_ID, HEADER_COUNTER_COL, row_idx)
        if counter == "":
            break
        counters[clean_counter(counter)] = row_idx
    return counters

def open_header_row(session, row_idx):
    description = get_cell_text(session, HEADER_TABLE_ID, HEADER_DESC_COL, row_idx)
    print(f"Opening routing row {row_idx}: {description}")
    session.findById(f"{HEADER_TABLE_ID}/txt[{HEADER_DESC_COL},{row_idx}]").setFocus()
    session.findById("wnd[0]").sendVKey(2)

def back_to_header(session):
    # Back button x3: component allocation -> ... -> routing header overview
    for _ in range(3):
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
    if session.findById(HEADER_TABLE_ID, False) is None:
        raise RuntimeError("Not back on routing header overview after 3x Back")

# ============================================================
# COMPONENT ALLOCATION
# ============================================================
def open_component_allocation(session, bom_row=8):
    # bom_row = row of the BOM in the popup (from recording: row 8)
    session.findById("wnd[0]/tbar[1]/btn[7]").press()
    session.findById(f"wnd[1]/usr/sub/1/txt[1,{bom_row}]").setFocus()
    session.findById("wnd[1]").sendVKey(2)

def is_blank_comp_row(session, row_idx):
    # Blank = no text in any cell before the operation column
    return all(get_cell_text(session, COMP_TABLE_ID, col, row_idx) == ""
               for col in range(COMP_OPERATION_COL))

def select_unassigned_components(session):
    # Selects every non-blank row whose operation (txt[10]) is empty.
    table = session.findById(COMP_TABLE_ID)
    total_rows = int(table.RowCount)
    visible_rows = int(table.VisibleRowCount)
    selected = []

    position = 0
    while position < total_rows:
        session.findById(COMP_TABLE_ID).VerticalScrollbar.Position = position
        first_row = int(session.findById(COMP_TABLE_ID).VerticalScrollbar.Position)

        reached_end = False
        for row_idx in range(min(visible_rows, total_rows - first_row)):
            if is_blank_comp_row(session, row_idx):
                reached_end = True
                break
            if get_cell_text(session, COMP_TABLE_ID, COMP_OPERATION_COL, row_idx) == "":
                absolute_row = first_row + row_idx
                session.findById(COMP_TABLE_ID).getAbsoluteRow(absolute_row).selected = True
                selected.append(absolute_row)

        if reached_end or first_row + visible_rows >= total_rows:
            break
        position = first_row + visible_rows

    print(f"Selected component rows: {selected}")
    return selected

def get_status_text(session):
    return session.findById("wnd[0]/sbar").text.strip()

def assign_to_operation(session, operation="0020"):
    # Back to the top and focus the table (recording did chk[0,row].setFocus)
    # so SAP picks up the row selection when btn[5] is pressed
    session.findById(COMP_TABLE_ID).VerticalScrollbar.Position = 0
    checkbox = session.findById(f"{COMP_TABLE_ID}/chk[0,0]", False)
    if checkbox is not None:
        checkbox.setFocus()

    session.findById("wnd[0]/tbar[1]/btn[5]").press()

    field = session.findById("wnd[1]/usr/txt[0]", False)
    if field is None:
        raise RuntimeError(
            f"btn[5] did not open the operation popup. SAP status: '{get_status_text(session)}'")
    field.text = operation
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    if session.findById("wnd[1]", False) is not None:
        raise RuntimeError(
            f"Operation popup still open after OK. SAP status: '{get_status_text(session)}'")
    press_enter(session)
    print(f"Assigned to {operation}. SAP status: '{get_status_text(session)}'")

def save_routing(session):
    session.findById("wnd[0]/tbar[0]/btn[11]").press()

# ============================================================
# MAIN FLOW
# ============================================================
def read_routing_groups(session, task: RoutingTask):
    # Open CA02 once just to read the group popup
    open_ca02(session)
    open_routing(session, task.material, task.plant)
    if has_group_popup(session):
        task.groups = read_group_popup(session)
    else:
        # Only one group: SAP goes straight to the header overview
        counters = list(get_header_counters(session))
        task.groups = [RoutingGroup(group="", popup_row=0, counters=counters)]
    for g in task.groups:
        print(f"Group {g.group or '(single)'}: GrC {g.counters}")

def process_group(session, task: RoutingTask, group: RoutingGroup, bom_row=8):
    # Each group is a separate CA02 entry, so it is saved before the next /NCA02
    open_ca02(session)
    open_routing(session, task.material, task.plant)
    if has_group_popup(session):
        select_popup_row(session, group.popup_row)

    header_counters = get_header_counters(session)
    for counter in group.counters:
        if counter not in header_counters:
            raise ValueError(f"GrC {counter} of group {group.group} not in header overview")

        print("-" * 60)
        open_header_row(session, header_counters[counter])
        open_component_allocation(session, bom_row)

        if select_unassigned_components(session):
            assign_to_operation(session, task.operation)
        else:
            print("No unassigned components, skipping")

        back_to_header(session)
        task.done.append(f"{group.group}/{counter}")

    save_routing(session)
    print(f"Saved group {group.group}: {session.findById('wnd[0]/sbar').text}")

def assign_bom_to_routings(session, task: RoutingTask, bom_row=8):
    read_routing_groups(session, task)
    for group in task.groups:
        process_group(session, task, group, bom_row)


if __name__ == "__main__":
    from connect import get_session
    session = slow_session(get_session())
    assign_bom_to_routings(session, RoutingTask(material="5016441210"))
