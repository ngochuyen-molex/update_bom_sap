"""Phase 2 - CA02 component assignment.

Usage:
    python routing_main.py [check] [slow]               task "newbom" (chain, CREATE_BOM.xlsx)
    python routing_main.py altbom [check] [slow]        task "altbom" (md_assign_component.xlsx)
    python routing_main.py [altbom] continue [slow]     carry on with the result sheet (no Snowflake)

    (no option)  build the work list, write the result sheet, run SAP
    check        build the work list, write the result sheet, stop (no SAP)
    continue     read the result sheet (no Snowflake), process PENDING / FAILED rows, update in place
                 ("retry" works too)
    fresh        rebuild the work list although the result sheet exists (SUCCESS rows are kept)
    slow         wait after every SAP action

Tasks:
    newbom  CREATE_BOM.xlsx: materials + new alt BOM from the phase 1 result columns of the
            first sheet (STATUS SUCCESS, Created Alt BOM), routings from Snowflake -> sheet
            "Routing Result". Only the newly created alt BOM
            (it is not in Snowflake yet, ~20 min lag, so no Snowflake allocation check).
    altbom  md_assign_component.xlsx: materials from the first sheet, every active alt BOM and
            the current allocation from Snowflake -> sheets "Result" + "Preview".
Both tasks: deleted / trial routings and trial alt BOMs are SKIPPED, same 0030 / 0020 rules.
Phase 2 never writes into the first sheet.
"""
import sys
import pandas as pd
import config
from connect import get_session, slow_from_args, SAPConnectionError
from models import Routing
from routing import process_routing, skip_reason, planned_operation
from routing_sf import load_assignment_worklist, load_routing_list, clean_material
from workbook import (CHAIN_FILE, ROUTING_SHEET, CONTINUE_STATUSES,
                      PARENT_COL, ALT_BOM_COL, CREATED_ALT_BOM_COL, STATUS_COL,
                      has_sheet, read_sheet, save_sheets, save_sheet, now, cell)

TASKS = {
    # task: (input file, result sheet, preview sheet or None)
    "newbom": (CHAIN_FILE, ROUTING_SHEET, None),
    "altbom": (config.ALTBOM_FILE, "Result", "Preview"),
}
ALTBOM_MATERIAL_COLUMN = "Material"         # first sheet of the altbom file
SAVE_EVERY = config.SAVE_EVERY              # save the result sheet after every N processed entries

RESULT_COLUMNS = ["Material", "Group", "GrC", "Description", "Alt BOM", "Alt BOM Text",
                  "Op 0030 WC", "Unassigned (Snowflake)", "STATUS", "MESSAGE", "Updated At"]
PREVIEW_COLUMNS = ["Material", "Group", "GrC", "Alt BOM", "Op 0030 WC", "Item", "Component",
                   "Component Description", "Current Operation", "Planned"]


def result_row(r: Routing, updated_at=None):
    return {
        "Material": r.material, "Group": r.group, "GrC": r.counter, "Description": r.description,
        "Alt BOM": r.alt_bom, "Alt BOM Text": r.alt_bom_text, "Op 0030 WC": r.op_0030_wc,
        "Unassigned (Snowflake)": "" if r.unassigned < 0 else r.unassigned,
        "STATUS": r.status, "MESSAGE": r.message, "Updated At": updated_at or now(),
    }


def entry_key(e: Routing):
    return (e.material, e.group, e.counter, e.alt_bom)


def set_skip_or_pending(entries, pending_message):
    for e in entries:
        reason = skip_reason(e)
        e.status, e.message = ("SKIPPED", reason) if reason else ("PENDING", pending_message(e))


def missing_rows(materials, entries, message):
    found = {e.material for e in entries}
    return [Routing(material=m, group="", counter="", status="SKIPPED", message=message)
            for m in materials if m not in found]


# ============================================================
# WORK LIST - task "newbom" (chain: new alt BOM from the phase 1 columns of Sheet1)
# ============================================================
def read_new_boms(path):
    # [(material, created alt BOM)] of the BOMs created in phase 1 (result columns in Sheet1)
    df = read_sheet(path, 0)
    if STATUS_COL not in df.columns or CREATED_ALT_BOM_COL not in df.columns:
        raise ValueError(f"No phase 1 result columns in the first sheet of {path}: run bom_create_main.py first")
    new_boms = []
    for _, row in df.iterrows():
        if cell(row, STATUS_COL) != "SUCCESS":
            continue
        material = clean_material(cell(row, PARENT_COL))
        alt_bom = clean_material(cell(row, CREATED_ALT_BOM_COL) or cell(row, ALT_BOM_COL))
        if material and alt_bom and (material, alt_bom) not in new_boms:
            new_boms.append((material, alt_bom))
    return new_boms


def build_newbom_worklist(path):
    new_boms = read_new_boms(path)
    materials = sorted({m for m, _ in new_boms})
    print(f"{len(new_boms)} new BOMs (STATUS SUCCESS in the first sheet), {len(materials)} materials")
    routings = load_routing_list(materials)

    entries = []
    for material, alt_bom in new_boms:
        for r in routings.get(material, []):
            entries.append(Routing(
                material=material, group=r.group, counter=r.counter, description=r.description,
                plant=r.plant, group_size=r.group_size, deleted=r.deleted,
                alt_bom=alt_bom, bom_count=max(1, int(alt_bom)),   # alt BOM n -> at least n BOMs
                op_0030_wc=r.op_0030_wc))                          # unassigned unknown (-1)
    set_skip_or_pending(entries, lambda e: "New alt BOM")
    return entries + missing_rows(materials, entries, "No routing in MLX_DATALAKE"), None


# ============================================================
# WORK LIST - task "altbom" (all active alt BOMs + allocation from Snowflake)
# ============================================================
def read_altbom_materials(path):
    df = read_sheet(path, 0)                                # first sheet = source
    if ALTBOM_MATERIAL_COLUMN not in df.columns:
        raise ValueError(f"Column '{ALTBOM_MATERIAL_COLUMN}' not found in {path}. Columns: {list(df.columns)}")
    materials = []
    for value in df[ALTBOM_MATERIAL_COLUMN]:
        material = clean_material(value)
        if material and material not in materials:
            materials.append(material)
    return materials


def preview_rows(entries, items):
    # Planned operation per component, only for the PENDING entries. Built once.
    todo = {entry_key(e): e for e in entries if e.status == "PENDING"}
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


def build_altbom_worklist(path):
    materials = read_altbom_materials(path)
    print(f"{len(materials)} materials")
    entries, items = load_assignment_worklist(materials)
    set_skip_or_pending(entries, lambda e: f"{e.unassigned} unassigned")
    missing = missing_rows(materials, entries, "No routing or no active alt BOM (usage 1) in MLX_DATALAKE")
    return entries + missing, preview_rows(entries, items)


BUILDERS = {"newbom": build_newbom_worklist, "altbom": build_altbom_worklist}


# ============================================================
# CONTINUE (result sheet -> entries, no Snowflake)
# ============================================================
def entries_from_result(df):
    # Returns (entries, updated_at) so rows that are not processed keep their timestamp
    with_group = df[df["Group"] != ""]
    group_sizes = with_group.groupby(["Material", "Group"])["GrC"].nunique().to_dict()
    bom_counts = with_group.groupby("Material")["Alt BOM"].nunique().to_dict()
    entries, updated_at = [], []
    for _, row in df.iterrows():
        unassigned = cell(row, "Unassigned (Snowflake)")
        alt_bom = cell(row, "Alt BOM")
        entries.append(Routing(
            material=cell(row, "Material"), group=cell(row, "Group"), counter=cell(row, "GrC"),
            description=cell(row, "Description"), alt_bom=alt_bom,
            alt_bom_text=cell(row, "Alt BOM Text"), op_0030_wc=cell(row, "Op 0030 WC"),
            group_size=int(group_sizes.get((cell(row, "Material"), cell(row, "Group")), 1)),
            bom_count=max(int(bom_counts.get(cell(row, "Material"), 1)), int(alt_bom or 1)),
            unassigned=int(float(unassigned)) if unassigned else -1,
            status=cell(row, "STATUS"), message=cell(row, "MESSAGE")))
        updated_at.append(cell(row, "Updated At"))
    return entries, updated_at


def keep_previous_success(entries, path, sheet):
    # "fresh": rows that were SUCCESS in the old result sheet stay SUCCESS (never redone),
    # same rule as production_version.keep_previous_success
    if not has_sheet(path, sheet):
        return [None] * len(entries)
    old, old_stamps = entries_from_result(read_sheet(path, sheet))
    done = {entry_key(e): (e, old_stamps[i]) for i, e in enumerate(old) if e.status == "SUCCESS"}
    stamps = []
    for e in entries:
        prev = done.get(entry_key(e))
        if prev:
            e.status, e.message = "SUCCESS", prev[0].message
            stamps.append(prev[1])
        else:
            stamps.append(None)
    kept = sum(s is not None for s in stamps)
    if kept:
        print(f"Kept {kept} SUCCESS rows from the previous '{sheet}' sheet")
    return stamps


# ============================================================
# RUN
# ============================================================
def run(session, path, sheet, entries, updated_at=None):
    updated_at = updated_at or [None] * len(entries)
    rows = [result_row(e, updated_at[i]) for i, e in enumerate(entries)]
    todo = [i for i, e in enumerate(entries) if e.status in CONTINUE_STATUSES and e.group]
    print(f"{len(entries)} entries, {len(todo)} to process in SAP")
    if not save_sheet(path, sheet, rows, RESULT_COLUMNS):
        return

    done = 0
    saved = True
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
                saved = save_sheet(path, sheet, rows, RESULT_COLUMNS)
                if not saved:
                    print(f"STOPPED. The last {SAVE_EVERY} results are NOT in Excel "
                          f"(a rerun with continue just finds them assigned).")
                    break
    finally:
        if saved and save_sheet(path, sheet, rows, RESULT_COLUMNS):    # also on Ctrl+C / unexpected error
            print(f"Result saved to sheet '{sheet}' of {path}")


def print_summary(entries):
    skipped = [e for e in entries if e.status == "SKIPPED"]
    reasons = pd.Series([e.message.split(":")[0] for e in skipped], dtype=str).value_counts()
    counts = {s: sum(e.status == s for e in entries) for s in ("PENDING", "SUCCESS", "SKIPPED", "FAILED")}
    print(f"Work list: {len(entries)} entries {counts}")
    for reason, count in reasons.items():
        print(f"  SKIPPED {count:>5}  {reason}")


def main():
    args = sys.argv[1:]
    task = "altbom" if "altbom" in args else "newbom"
    path, sheet, preview_sheet = TASKS[task]
    print(f"Phase 2 - assign components, task: {task}, file: {path}")

    try:
        if "continue" in args or "retry" in args:
            entries, updated_at = entries_from_result(read_sheet(path, sheet))
            todo = sum(e.status in CONTINUE_STATUSES and bool(e.group) for e in entries)
            print(f"Loaded {len(entries)} entries from sheet '{sheet}', {todo} PENDING / FAILED")
            if todo == 0:
                print("Nothing to do.")
                return
            run(get_session(slow=slow_from_args()), path, sheet, entries, updated_at)
            return

        if has_sheet(path, sheet) and "fresh" not in args:
            # Don't overwrite the progress of an earlier run by accident
            print(f"Sheet '{sheet}' already exists in {path}.\n"
                  f"  To carry on with it:  python routing_main.py {task} continue\n"
                  f"  To rebuild the list:  python routing_main.py {task} fresh  "
                  f"(replaces '{sheet}', keeps its SUCCESS rows)")
            return

        entries, preview = BUILDERS[task](path)
        updated_at = keep_previous_success(entries, path, sheet)
        print_summary(entries)

        sheets = {sheet: pd.DataFrame([result_row(e, updated_at[i]) for i, e in enumerate(entries)],
                                      columns=RESULT_COLUMNS)}
        if preview_sheet:
            sheets[preview_sheet] = preview
        if not save_sheets(path, sheets):
            return
        print(f"Written to {path}: {', '.join(sheets)}")
        if "check" in args:
            return

        session = get_session(slow=slow_from_args())
        print(f"Connected to SAP system {session.Info.SystemName}")
        run(session, path, sheet, entries, updated_at)

    except PermissionError:
        print(f"Cannot write {path}. Close it in Excel and run again.")
    except SAPConnectionError as e:
        print(f"SAP Connection Error: {e}")


if __name__ == "__main__":
    main()
