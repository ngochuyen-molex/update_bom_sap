"""CA02 component assignment: every routing x active alt BOM of a list of materials.

Usage:
    python routing_main.py [check] [slow]               task "newbom": materials of the BOM creation result
    python routing_main.py altbom [check] [slow]        task "altbom": materials of md_assign_component.xlsx
    python routing_main.py [altbom] continue [slow]     carry on with the Result sheet (no Snowflake)

    (no option)  get the work list from Snowflake, write Result + Preview sheets, run SAP
    check        get the work list from Snowflake, write Result + Preview sheets, stop (no SAP)
    continue     read the Result sheet (no Snowflake), process PENDING / FAILED rows, update in place
                 ("retry" works too); Preview is not rebuilt
    fresh        rebuild Result + Preview from Snowflake although Result already exists
    slow         wait after every SAP action

Sheets (order: source, Result, Preview):
    Result   one row per routing x alt BOM, STATUS PENDING (needs assignment) / SKIPPED (reason
             in MESSAGE) / SUCCESS / FAILED
    Preview  one row per component of the PENDING entries with the planned operation
Both tasks use the same logic; they only differ in the input file and how the materials are
read. The first sheet of the input file is never changed.
"""
import os
import shutil
import sys
from datetime import datetime
import pandas as pd
from connect import get_session, slow_from_args, SAPConnectionError
from models import Routing
from routing import process_routing, skip_reason, planned_operation
from routing_sf import load_assignment_worklist, clean_material

TASKS = {
    # task: (input file, material column, only rows with STATUS = SUCCESS)
    "newbom": (r"C:\Users\nguyeb3\Downloads\create_bom_result.xlsx", "Parent Material", True),
    "altbom": (r"C:\Users\nguyeb3\Downloads\md_assign_component.xlsx", "Material", False),
}
RESULT_SHEET = "Result"
PREVIEW_SHEET = "Preview"                   # planned operation per component (PENDING entries)
CONTINUE_STATUSES = ("PENDING", "FAILED")   # PENDING = needs assignment, not processed yet
SAVE_EVERY = 5                              # save Result after every N processed entries

RESULT_COLUMNS = ["Material", "Group", "GrC", "Description", "Alt BOM", "Alt BOM Text",
                  "Op 0030 WC", "Unassigned (Snowflake)", "STATUS", "MESSAGE", "Updated At"]
PREVIEW_COLUMNS = ["Material", "Group", "GrC", "Alt BOM", "Op 0030 WC", "Item", "Component",
                   "Component Description", "Current Operation", "Planned"]


# ============================================================
# EXCEL (first sheet untouched, result sheets replaced)
# ============================================================
def write_sheets(path, sheets):
    # Replaces only the given sheets; temp copy + replace, so a crash can't corrupt the file.
    # Sheet order afterwards: first sheet (source), Result, Preview, any others.
    temp = path.replace(".xlsx", "_temp.xlsx")
    shutil.copy2(path, temp)
    with pd.ExcelWriter(temp, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)
        book = writer.book
        first, rest = book.worksheets[0], book.worksheets[1:]
        order = [RESULT_SHEET, PREVIEW_SHEET]
        rest.sort(key=lambda ws: order.index(ws.title) if ws.title in order else len(order))
        book._sheets = [first] + rest
        book.active = 0
    os.replace(temp, path)


def save_result(path, rows):
    try:
        write_sheets(path, {RESULT_SHEET: pd.DataFrame(rows, columns=RESULT_COLUMNS)})
    except PermissionError:
        print(f"Could not save {path}. Close it in Excel.")


def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def cell(row, column):
    value = row.get(column)
    return "" if pd.isna(value) else str(value).strip()


# ============================================================
# WORK LIST
# ============================================================
def read_materials(path, column, success_only):
    df = pd.read_excel(path, sheet_name=0, dtype=str)       # first sheet = source
    if column not in df.columns:
        raise ValueError(f"Column '{column}' not found in {path}. Columns: {list(df.columns)}")
    if success_only:
        if "STATUS" not in df.columns:
            raise ValueError(f"Column 'STATUS' not found in {path} (needed for this task)")
        df = df[df["STATUS"].astype(str).str.strip() == "SUCCESS"]
    materials = []
    for value in df[column].dropna():
        material = clean_material(value)
        if material and material not in materials:
            materials.append(material)
    return materials


def result_row(r: Routing, updated_at=None):
    return {
        "Material": r.material, "Group": r.group, "GrC": r.counter, "Description": r.description,
        "Alt BOM": r.alt_bom, "Alt BOM Text": r.alt_bom_text, "Op 0030 WC": r.op_0030_wc,
        "Unassigned (Snowflake)": "" if r.unassigned < 0 else r.unassigned,
        "STATUS": r.status, "MESSAGE": r.message, "Updated At": updated_at or now(),
    }


def preview_rows(entries, items):
    # Planned operation per component, only for entries that will be opened in SAP (PENDING).
    # Built once from the work list data; "continue" does not rebuild it.
    todo = {(e.material, e.group, e.counter, e.alt_bom): e for e in entries if e.status == "PENDING"}
    rows = []
    for it in items.itertuples(index=False):
        e = todo.get((it.MATERIAL, it.ROUTING_GROUP, it.GROUP_COUNTER or "0", it.ALT_BOM))
        if e is None:
            continue
        if it.OPERATION:
            planned = f"already {it.OPERATION}"
        else:
            operation, reason = planned_operation(it.COMPONENT_DESCRIPTION, e.op_0030_wc)
            planned = f"{operation} ({reason})" if reason else operation
        rows.append({
            "Material": e.material, "Group": e.group, "GrC": e.counter, "Alt BOM": e.alt_bom,
            "Op 0030 WC": e.op_0030_wc, "Item": it.ITEM, "Component": it.COMPONENT,
            "Component Description": it.COMPONENT_DESCRIPTION,
            "Current Operation": it.OPERATION, "Planned": planned,
        })
    rows.sort(key=lambda r: (r["Material"], r["Group"], int(r["GrC"] or 0), int(r["Alt BOM"] or 0), r["Item"]))
    return pd.DataFrame(rows, columns=PREVIEW_COLUMNS)


def build_worklist(materials):
    # Returns (entries, preview): preview = one row per component of the PENDING entries
    entries, items = load_assignment_worklist(materials)

    for e in entries:
        reason = skip_reason(e)
        e.status, e.message = ("SKIPPED", reason) if reason else ("PENDING", f"{e.unassigned} unassigned")

    # Materials without a routing or without an active alt BOM
    found = {e.material for e in entries}
    missing = [Routing(material=m, group="", counter="", status="SKIPPED",
                       message="No routing or no active alt BOM (usage 1) in MLX_DATALAKE")
               for m in materials if m not in found]
    return entries + missing, preview_rows(entries, items)


def entries_from_result(df):
    # Continue: rebuild entries from the Result sheet (no Snowflake).
    # Returns (entries, updated_at) so rows that are not processed keep their timestamp.
    with_group = df[df["Group"] != ""]
    group_sizes = with_group.groupby(["Material", "Group"])["GrC"].nunique().to_dict()
    bom_counts = with_group.groupby("Material")["Alt BOM"].nunique().to_dict()
    entries, updated_at = [], []
    for _, row in df.iterrows():
        unassigned = cell(row, "Unassigned (Snowflake)")
        entries.append(Routing(
            material=cell(row, "Material"), group=cell(row, "Group"), counter=cell(row, "GrC"),
            description=cell(row, "Description"), alt_bom=cell(row, "Alt BOM"),
            alt_bom_text=cell(row, "Alt BOM Text"), op_0030_wc=cell(row, "Op 0030 WC"),
            group_size=int(group_sizes.get((cell(row, "Material"), cell(row, "Group")), 1)),
            bom_count=int(bom_counts.get(cell(row, "Material"), 1)),
            unassigned=int(float(unassigned)) if unassigned else -1,
            status=cell(row, "STATUS"), message=cell(row, "MESSAGE")))
        updated_at.append(cell(row, "Updated At"))
    return entries, updated_at


# ============================================================
# RUN
# ============================================================
def run(session, path, entries, updated_at=None):
    updated_at = updated_at or [None] * len(entries)
    rows = [result_row(e, updated_at[i]) for i, e in enumerate(entries)]
    todo = [i for i, e in enumerate(entries) if e.status in CONTINUE_STATUSES and e.group]
    print(f"{len(entries)} entries, {len(todo)} to process in SAP")
    save_result(path, rows)

    done = 0
    try:
        for i in todo:
            e = entries[i]
            try:
                process_routing(session, e)
            except Exception as exc:
                e.status, e.message = "FAILED", str(exc)
            print(f"{e.key}: {e.status} {e.message}")
            rows[i] = result_row(e)
            done += 1
            if done % SAVE_EVERY == 0:
                save_result(path, rows)
    finally:
        save_result(path, rows)     # also on Ctrl+C / unexpected error
        print(f"Result saved to sheet '{RESULT_SHEET}' of {path}")


def has_result_sheet(path):
    with pd.ExcelFile(path) as xl:      # close the file again, or Windows blocks the next save
        return RESULT_SHEET in xl.sheet_names


def main():
    args = sys.argv[1:]
    task = "altbom" if "altbom" in args else "newbom"
    path, column, success_only = TASKS[task]
    print(f"Task: {task}, file: {path}")

    try:
        if "continue" in args or "retry" in args:
            df = pd.read_excel(path, sheet_name=RESULT_SHEET, dtype=str).fillna("")
            entries, updated_at = entries_from_result(df)
            todo = sum(e.status in CONTINUE_STATUSES and bool(e.group) for e in entries)
            print(f"Loaded {len(entries)} entries from sheet '{RESULT_SHEET}', {todo} PENDING / FAILED")
            if todo == 0:
                print("Nothing to do.")
                return
            session = get_session(slow=slow_from_args())
            run(session, path, entries, updated_at)
            return

        if has_result_sheet(path) and "fresh" not in args:
            # Don't overwrite the progress of an earlier run by accident
            print(f"Sheet '{RESULT_SHEET}' already exists in {path}.\n"
                  f"  To carry on with it:      python routing_main.py {task} continue\n"
                  f"  To rebuild from Snowflake: python routing_main.py {task} fresh  (replaces '{RESULT_SHEET}')")
            return

        materials = read_materials(path, column, success_only)
        print(f"{len(materials)} materials")
        entries, preview = build_worklist(materials)

        skipped = [e for e in entries if e.status == "SKIPPED"]
        reasons = pd.Series([e.message.split(":")[0] for e in skipped]).value_counts()
        print(f"Work list: {len(entries)} routing x alt BOM entries, "
              f"{sum(e.status == 'PENDING' for e in entries)} to process, {len(skipped)} skipped")
        for reason, count in reasons.items():
            print(f"  SKIPPED {count:>5}  {reason}")

        write_sheets(path, {
            RESULT_SHEET: pd.DataFrame([result_row(e) for e in entries], columns=RESULT_COLUMNS),
            PREVIEW_SHEET: preview,
        })
        print(f"Sheets '{RESULT_SHEET}' and '{PREVIEW_SHEET}' ({len(preview)} components) written to {path}")
        if "check" in args:
            return

        session = get_session(slow=slow_from_args())
        print(f"Connected to SAP system {session.Info.SystemName}")
        run(session, path, entries)

    except PermissionError:
        print(f"Cannot write {path}. Close it in Excel and run again.")
    except SAPConnectionError as e:
        print(f"SAP Connection Error: {e}")


if __name__ == "__main__":
    main()
