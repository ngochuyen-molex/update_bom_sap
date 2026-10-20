import math
import pandas as pd
# SAP Bom item table
BOM_TABLE_ID = (
    "wnd[0]/usr/tabsTS_ITOV/tabpTCMA/"
    "ssub/2/3/tblSAPLCSDITCMAT")

# ============================================================
# BASIC HELPERS / UTILITY
# ============================================================

def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()

def clean_component(value):
    value = clean_text(value)
    try:
        # handles (.0, .000)
        return str(int(float(value))).lstrip("0")
    except ValueError:
        return value.lstrip("0")

def clean_item(value):
    value = clean_text(value)

    if value.endswith(".0"):
        value = value[:-2]

    if value.isdigit():
        if len(value) == 1 and value != "0":
            value = str(int(value) * 10)
        if len(value) < 4:
            value = value.zfill(4)

    return value

def clean_number(value):
    if pd.isna(value):
        return None

    value = str(value).strip()

    if value == "":
        return None

    # SAP negative format
    # 160- -> -160
    if value.endswith("-"):
        value = "-" + value[:-1]

    # remove thousand separators
    value = value.replace(",", "")

    return float(value)

def numbers_equal(a, b, decimal_places=2):
    a = clean_number(a)
    b = clean_number(b)
    if a is None and b is None:
        return True

    if a is None or b is None:
        return False

    return round(a, decimal_places) == round(b, decimal_places)

def set_text(session, element_id, value):
    session.findById(element_id).text = str(value)

def press_enter(session):
    session.findById("wnd[0]").sendVKey(0)
# ============================================================
# NAVIGATION
# ============================================================
def open_cs02(session):
    session.findById("wnd[0]/tbar[0]/okcd").text = "/NCS02"
    press_enter(session)

def open_bom(
    session,
    material,
    alt_bom,
    bom_usage="1",
    plant="1901",
    component=None,
    item=None,
):
    print(f"Material={clean_component(material)}, "
        f"Alt BOM={clean_component(alt_bom)}", 
        f"Component={component}",
        f"Item={clean_item(item)}")
    set_text(session, "wnd[0]/usr/ctxt[0]", clean_component(material))
    set_text(session, "wnd[0]/usr/ctxt[1]", plant)
    set_text(session, "wnd[0]/usr/ctxt[2]", bom_usage)
    set_text(session, "wnd[0]/usr/txt[3]", clean_component(alt_bom))

    press_enter(session)

def get_bom_item_table(session):
    return session.findById(BOM_TABLE_ID)

def open_bom_header(session):
    session.findById("wnd[0]/tbar[1]/btn[6]").press()
# ============================================================
# READ CURRENT SAP VALUES
# ============================================================

def get_item(session, row_idx):
    return session.findById(
        f"{BOM_TABLE_ID}/txt[0,{row_idx}]").text.strip()

def get_component(session, row_idx):
    return session.findById(
        f"{BOM_TABLE_ID}/ctxt[2,{row_idx}]").text.strip()

def get_quantity(session, row_idx):
    value = session.findById(
        f"{BOM_TABLE_ID}/txt[4,{row_idx}]").text
    return clean_number(value)

# to get status bar of material number (component one)
# underneath the SAP window, use this function
def get_component_status_bar_text(session):
    return session.findById("wnd[0]/sbar/pane[0]").text.strip()

# ============================================================
# FIND ROW
# ============================================================

def find_item_row(session, target_item):
    table = get_bom_item_table(session)

    target_item = clean_item(target_item)
    if not target_item:
        raise ValueError("Item number is blank.")

    total_rows = int(table.RowCount)
    visible_rows = int(table.VisibleRowCount)
    if visible_rows <= 0:
        raise ValueError("SAP BOM table has no visible rows to search.")

    scrollbar = table.VerticalScrollbar
    last_start = max(0, total_rows - visible_rows)
    page_starts = list(range(0, last_start + 1, visible_rows))
    if not page_starts or page_starts[-1] != last_start:
        page_starts.append(last_start)

    scanned_rows = set()
    for requested_start in page_starts:
        scrollbar.Position = requested_start
        first_row = int(scrollbar.Position)
        page_row_count = min(visible_rows, total_rows - first_row)

        for visible_idx in range(page_row_count):
            absolute_idx = first_row + visible_idx
            if absolute_idx in scanned_rows:
                continue
            scanned_rows.add(absolute_idx)

            current_item = clean_item(get_item(session, visible_idx))
            if current_item == target_item:
                return visible_idx

    raise ValueError(f"Item {target_item} not found in BOM table.")

def find_component_row(session, target_component):
    table = get_bom_item_table(session)
    target_component = clean_component(target_component)
    if not target_component:
        raise ValueError("Component before is blank.")

    total_rows = int(table.RowCount)
    visible_rows = int(table.VisibleRowCount)
    if visible_rows <= 0:
        raise ValueError("SAP BOM table has no visible rows to search.")

    scrollbar = table.VerticalScrollbar
    last_start = max(0, total_rows - visible_rows)
    page_starts = list(range(0, last_start + 1, visible_rows))
    if not page_starts or page_starts[-1] != last_start:
        page_starts.append(last_start)

    scanned_rows = set()
    matching_rows = []
    for requested_start in page_starts:
        scrollbar.Position = requested_start
        first_row = int(scrollbar.Position)
        page_row_count = min(visible_rows, total_rows - first_row)

        for visible_idx in range(page_row_count):
            absolute_idx = first_row + visible_idx
            if absolute_idx in scanned_rows:
                continue
            scanned_rows.add(absolute_idx)

            current_component = clean_component(
                get_component(session, visible_idx)
            )
            if current_component == target_component:
                matching_rows.append((absolute_idx, first_row))

    if not matching_rows:
        raise ValueError(
            f"Component {target_component} not found in BOM table."
        )
    if len(matching_rows) > 1:
        raise ValueError(
            f"Component {target_component} matches multiple BOM rows: "
            f"{matching_rows}."
        )

    absolute_row, page_start = matching_rows[0]
    scrollbar.Position = page_start
    visible_row = absolute_row - int(scrollbar.Position)
    if not 0 <= visible_row < visible_rows:
        raise ValueError(
            f"Found component {target_component} at SAP row {absolute_row}, "
            "but could not restore that row into the visible table."
        )

    actual_component = get_component(session, visible_row)
    if clean_component(actual_component) != target_component:
        raise ValueError(
            f"Found component {target_component} at SAP row {absolute_row}, "
            f"but after restoring the table SAP row {visible_row} contains "
            f"{actual_component!r}. No update was made."
        )

    return visible_row

def is_blank_row(row):
    return (
        clean_text(row["Parent Material"]) == ""
        and clean_text(row["Alternative BOM"]) == ""
        and clean_text(row.get("Item")) == ""
    )
# ============================================================
# VALIDATION BEFORE UPDATE
# ============================================================

def has_changes(excel_row):
    component_before = clean_component(excel_row["Component before"])
    component_after = clean_component(excel_row["Component after"])

    component_changed = component_before != component_after

    qty_before = excel_row["BOM rate before"]
    qty_after = excel_row["BOM rate after"]

    qty_changed = not numbers_equal(qty_before, qty_after)

    return component_changed or qty_changed

def validate_excel_row_has_required_data(excel_row):
    required_columns = [
        "Parent Material",
        "Alternative BOM",
        "Item",
        "Component before",
        "Component after",
        "BOM rate before",
        "BOM rate after",
    ]

    for col in required_columns:
        if col not in excel_row.index:
            raise ValueError(f"Missing Excel column: {col}")

    if clean_text(excel_row["Parent Material"]) == "":
        raise ValueError("Parent Material is blank.")

    if clean_text(excel_row["Alternative BOM"]) == "":
        raise ValueError("Alternative BOM is blank.")

    if clean_text(excel_row["Item"]) == "":
        raise ValueError("Item is blank.")

def check_current_bom_status(session, row_idx, excel_row):
    current_component = get_component(session, row_idx)
    current_qty = get_quantity(session, row_idx)

    expected_component = clean_component(
        excel_row["Component before"])

    expected_qty = clean_number(
        excel_row["BOM rate before"])

    if clean_component(current_component) != expected_component:
        raise ValueError(
            f"Component mismatch. "
            f"SAP={current_component}, "
            f"Excel before={expected_component}")

    # if not numbers_equal(current_qty, expected_qty):
    #     raise ValueError(
    #         f"Quantity mismatch. "
    #         f"SAP={current_qty}, "
    #         f"Excel before={expected_qty}")

# ============================================================
# UPDATE
# ============================================================

def update_component(session, row_idx, component_after):
    session.findById(
        f"{BOM_TABLE_ID}/ctxt[2,{row_idx}]").text = clean_component(component_after)

def update_quantity(session, row_idx, qty_after):
    session.findById(
        f"{BOM_TABLE_ID}/txt[4,{row_idx}]").text = clean_number(qty_after)

def save_bom(session):
    # Standard SAP Save button is usually btn[11].
    session.findById("wnd[0]/tbar[0]/btn[11]").press()

def update_base_quantity(session,base_qty):
    session.findById("wnd[0]/usr/tabsTS_HEAD/tabpKHPT/ssub/2/3/txt[2]").text = clean_number(base_qty)
    press_enter
# ============================================================
# DEBUG HELPERS
# ============================================================

def print_table_info(session):
    table = get_bom_item_table(session)

    print("Table type:", type(table))
    print("Row count:", table.RowCount)

    try:
        print("Column count:", table.Columns.Count)

        for i in range(table.Columns.Count):
            col = table.Columns(i)
            print(i, col.Name)

    except Exception as e:
        print("Could not print columns:", e)

def print_focused_field(session):
    field = session.ActiveWindow.GuiFocus

    print("ID:", field.Id)
    print("Name:", field.Name)
    print("Type:", field.Type)