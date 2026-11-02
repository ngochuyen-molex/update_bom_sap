import pandas as pd
from bom import clean_component, clean_number, clean_text
from models import BomHeader, BomItem

# Excel column names (after unmerge)
PARENT_COL = "Parent Material"
ALT_BOM_COL = "Alternative BOM"
COMPONENT_COL = "Component after"
QUANTITY_COL = "BOM rate after"
UNIT_COL = "Unit after"
SCRAP_COL = "Component Scrap after"
BASE_QTY_COL = "Base Qty after"
ACTION_COL = "Action"           # optional: CREATE / CHANGE

VALID_ACTIONS = ("CREATE", "CHANGE")


def decide_action(alt_bom, action_value):
    # Action column wins. If empty: alt 1 = existing BOM -> CHANGE, else CREATE.
    action = clean_text(action_value).upper()
    if action:
        if action not in VALID_ACTIONS:
            raise ValueError(f"Invalid Action '{action_value}', use CREATE or CHANGE")
        return action
    return "CHANGE" if alt_bom == "1" else "CREATE"


def read_boms(file_path):
    df = pd.read_excel(file_path, dtype=str)

    required = [PARENT_COL, ALT_BOM_COL, COMPONENT_COL, QUANTITY_COL, UNIT_COL, SCRAP_COL, BASE_QTY_COL]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing Excel columns: {missing}")
    has_action = ACTION_COL in df.columns

    boms = {}
    for idx, row in df.iterrows():
        excel_row = idx + 2     # +1 header, +1 because Excel starts at 1
        component = clean_component(row[COMPONENT_COL])
        if not component:
            continue

        parent = clean_component(row[PARENT_COL])
        alt_bom = clean_text(row[ALT_BOM_COL])
        if not parent or not alt_bom:
            raise ValueError(f"Row {excel_row}: Parent Material or Alternative BOM is blank (run unmerge first)")

        key = (parent, alt_bom)
        if key not in boms:
            boms[key] = BomHeader(
                parent=parent,
                alt_bom=alt_bom,
                base_qty=clean_number(row[BASE_QTY_COL]),
                action=decide_action(alt_bom, row[ACTION_COL] if has_action else None),
            )

        boms[key].items.append(BomItem(
            excel_row=excel_row,
            component=component,
            quantity=clean_number(row[QUANTITY_COL]),
            unit=clean_text(row[UNIT_COL]),
            scrap=clean_number(row[SCRAP_COL]) or 0.0,
        ))

    return list(boms.values())


def validate_boms(boms):
    # Returns a list of problems. Empty list = OK to run.
    problems = []
    for bom in boms:
        name = f"{bom.parent} alt {bom.alt_bom} (rows {bom.excel_rows[0]}-{bom.excel_rows[-1]})"
        if bom.base_qty is None:
            problems.append(f"{name}: base quantity is blank")
        if bom.action == "CHANGE" and len(bom.items) != 1:
            problems.append(f"{name}: CHANGE supports 1 component only, found {len(bom.items)}")

        seen = set()
        for item in bom.items:
            if item.quantity is None:
                problems.append(f"{name}, row {item.excel_row}: quantity is blank")
            if not item.unit:
                problems.append(f"{name}, row {item.excel_row}: unit is blank")
            if item.component in seen:
                problems.append(f"{name}, row {item.excel_row}: component {item.component} appears twice")
            seen.add(item.component)
    return problems


def print_boms(boms):
    for i, bom in enumerate(boms, 1):
        print(f"[{i}] {bom.action:<6} parent={bom.parent} alt={bom.alt_bom} "
            f"base={bom.base_qty:g} rows={bom.excel_rows[0]}-{bom.excel_rows[-1]}")
        for item in bom.items:
            print(f"      row {item.excel_row:<4} {item.item_category}  {item.component:<10} "
                f"qty {item.quantity:>9g}  {item.unit:<3} scrap {item.scrap:g}")
