from bom import BOM_TABLE_ID, clean_component, clean_number, clean_text, set_text, press_enter
from models import BomItem

# Item table columns (row index = visible row, starts at 0)
ICT_COL = "ctxt[1,{row}]"        # item category, e.g. "L"
COMPONENT_COL = "ctxt[2,{row}]"
QUANTITY_COL = "txt[4,{row}]"
UNIT_COL = "ctxt[5,{row}]"


def press_button(session, element_id, step):
    # findById(id, False) returns None instead of raising when the element is missing
    button = session.findById(element_id, False)
    if button is None:
        status = session.findById("wnd[0]/sbar").text
        raise RuntimeError(f"Step '{step}': {element_id} not found. SAP status: '{status}'")
    button.press()

# ============================================================
# NAVIGATION (CS01 - CREATE BOM)
# ============================================================
def open_cs01(session):
    session.findById("wnd[0]/tbar[0]/okcd").text = "/NCS01"
    press_enter(session)

def open_bom_cs01(
    session,
    material,
    bom_usage="1",
    plant="1901",
):
    # CS01 assigns the alternative BOM itself, so alt_bom is not typed here
    print(f"Material={clean_component(material)}, "
        f"Plant={plant}, Usage={bom_usage}")
    set_text(session, "wnd[0]/usr/ctxt[0]", clean_component(material))
    set_text(session, "wnd[0]/usr/ctxt[1]", plant)
    set_text(session, "wnd[0]/usr/ctxt[2]", bom_usage)

    # Steps from the CS01 recording, up to the item table
    ## press_button(session, "wnd[0]/tbar[1]/btn[7]", "btn[7] first press")
    ## press_button(session, "wnd[0]/tbar[1]/btn[7]", "btn[7] second press")
    press_enter(session)
    press_enter(session)
    ## press_button(session, "wnd[1]/tbar[0]/btn[0]", "first popup OK")
    ## press_button(session, "wnd[0]/tbar[1]/btn[5]", "btn[5]")
    ## press_button(session, "wnd[1]/usr/btn[0]", "second popup")
    # Item table is now open

def get_created_alt_bom(session):
    # Alternative BOM number SAP assigned, shown in the header area of the item table view
    return session.findById("wnd[0]/usr/ssub/1/txt[2]").text.strip()

# ============================================================
# ITEM TABLE
# ============================================================
def format_quantity(value):
    # 8542.0 -> "8542", 2.8 -> "2.8", -1391.0 -> "1391-" (SAP style minus)
    number = clean_number(value)
    if number is None:
        raise ValueError(f"Invalid quantity: {value}")
    text = f"{abs(number):.3f}".rstrip("0").rstrip(".")
    return f"{text}-" if number < 0 else text

def enter_item(session, row_idx, component, quantity, unit, item_category="L"):
    print(f"Row {row_idx}: ICt={item_category}, Component={clean_component(component)}, "
        f"Qty={format_quantity(quantity)}, Unit={clean_text(unit)}")
    set_text(session, f"{BOM_TABLE_ID}/{ICT_COL.format(row=row_idx)}", item_category)
    set_text(session, f"{BOM_TABLE_ID}/{COMPONENT_COL.format(row=row_idx)}", clean_component(component))
    set_text(session, f"{BOM_TABLE_ID}/{QUANTITY_COL.format(row=row_idx)}", format_quantity(quantity))
    set_text(session, f"{BOM_TABLE_ID}/{UNIT_COL.format(row=row_idx)}", clean_text(unit))

def enter_items(session, items: list[BomItem]):
    table = session.findById(BOM_TABLE_ID)
    visible_rows = table.VisibleRowCount
    if len(items) > visible_rows:
        # TODO: scroll the table for BOMs with more items than visible rows
        raise RuntimeError(f"{len(items)} items but only {visible_rows} visible rows in item table")

    for row_idx, item in enumerate(items):
        enter_item(session, row_idx, item.component, item.quantity, item.unit, item.item_category)
