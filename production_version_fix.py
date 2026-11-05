"""One-off: fix the production versions of the half-finished run.

Reads the result file of production_version.py (no Snowflake) and updates it in place.
Only rows whose FIX STATUS is not SUCCESS (blank, FAILED or SKIPPED) are processed, so it can be rerun:
  STATUS = SUCCESS        -> version exists with the old text: rename it to "<description> B<alt BOM>"
  STATUS = FAILED / blank -> not created yet: create it (same as production_version.py)
  TRIAL routing           -> SKIPPED

    python production_version_fix.py [slow]
"""
import pandas as pd
from connect import get_session, slow_from_args, SAPConnectionError
from excel_utils import save_progress
from models import ProductionVersion
from routing import find_row, get_cell_text
from production_version import (
    RESULT_FILE, VERSION_TABLE_ID, VERSION_COL,
    open_mm02, open_production_versions, process_production_version,
    save_material, version_text, is_trial)

INPUT_FILE = r"C:\Users\nguyeb3\Downloads\create_bom_final_prod_version_fix.xlsx"       # ..._prod_version.xlsx, updated in place
TEXT_COL = VERSION_COL + 1      # description is the column after the version
DONE = ("SUCCESS",)         # FIX STATUS values that are not processed again (blank / FAILED / SKIPPED are)


def cell(row, column):
    value = row.get(column)
    return "" if pd.isna(value) else str(value).strip()


def set_cell_text(session, table_id, col, row_idx, value):
    # Cells are txt or ctxt depending on the field; try both
    for kind in ("txt", "ctxt"):
        element = session.findById(f"{table_id}/{kind}[{col},{row_idx}]", False)
        if element is not None:
            element.text = value
            return
    raise RuntimeError(f"No text cell at column {col}, row {row_idx}")


def leave_without_saving(session):
    # Close the version list and go back to the MM02 initial screen, nothing saved.
    # Without this the next row starts on the MRP 4 screen and can't find the material field.
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    open_mm02(session)


def rename_version(session, pv: ProductionVersion):
    # Existing version pv.version: replace its text (routing description) with the new text
    print("-" * 60)
    print(f"Rename {pv.key} version {pv.version}: '{pv.description}'")
    if is_trial(pv.description):
        pv.status, pv.message = "SKIPPED", "Trial routing, not renamed"
        return
    pv.text = version_text(pv.description, pv.alt_bom, pv.material)
    open_production_versions(session, pv.material)

    row_idx = find_row(session, VERSION_TABLE_ID, VERSION_COL, pv.version)
    if row_idx is None:
        raise ValueError(f"Version {pv.version} not found in production version list")

    current = get_cell_text(session, VERSION_TABLE_ID, TEXT_COL, row_idx)
    if current == pv.text:
        leave_without_saving(session)
        pv.status, pv.message = "SKIPPED", f"Version {pv.version} already '{pv.text}'"
        return
    if current != pv.description:
        # Safety check: this version does not have the text we created -> don't touch it
        leave_without_saving(session)
        raise ValueError(f"Version {pv.version} has text '{current}', expected '{pv.description}'. Not changed")

    print(f"  '{current}' -> '{pv.text}'")
    set_cell_text(session, VERSION_TABLE_ID, TEXT_COL, row_idx, pv.text)
    session.findById("wnd[1]/tbar[0]/btn[0]").press()     # continue
    status_text = save_material(session)
    pv.status = "SUCCESS"
    pv.message = f"Version {pv.version} renamed. {status_text}"
    print(f"Saved: {pv.message}")


def main():
    try:
        session = get_session(slow=slow_from_args())

        df = pd.read_excel(INPUT_FILE, dtype=str)     # dtype=str keeps "0004", "2"
        for column in ("Version Text", "FIX ACTION", "FIX STATUS", "FIX MESSAGE"):
            if column not in df.columns:
                df[column] = ""
        df = df.fillna("")
        todo = [idx for idx, row in df.iterrows()
                if cell(row, "Group") and cell(row, "FIX STATUS") not in DONE]
        print(f"Loaded {len(df)} rows from {INPUT_FILE}, {len(todo)} to process")

        open_mm02(session)
        for idx in todo:
            row = df.loc[idx]
            pv = ProductionVersion(
                material=cell(row, "Material"), group=cell(row, "Group"),
                counter=cell(row, "GrC"), description=cell(row, "Description"),
                alt_bom=cell(row, "Alt BOM"))

            # FIX ACTION from an earlier run wins; otherwise: created before -> rename, else create
            action = cell(row, "FIX ACTION")
            if action not in ("RENAME", "CREATE"):
                action = "RENAME" if cell(row, "STATUS") == "SUCCESS" else "CREATE"
            try:
                if action == "RENAME":
                    pv.version = cell(row, "New Version")
                    rename_version(session, pv)
                else:
                    process_production_version(session, pv)
            except Exception as e:
                pv.status, pv.message = "FAILED", str(e)
                try:
                    open_mm02(session)      # start the next row from a clean MM02 screen
                except Exception:
                    pass
            print(f"{action} {pv.key}: {pv.status} {pv.message}")

            df.at[idx, "FIX ACTION"] = action
            df.at[idx, "FIX STATUS"] = pv.status
            df.at[idx, "FIX MESSAGE"] = pv.message
            df.at[idx, "Version Text"] = pv.text
            if pv.version:
                df.at[idx, "New Version"] = pv.version
            try:
                save_progress(INPUT_FILE, df)   # temp file + replace, so a crash can't corrupt it
            except PermissionError:
                print(f"Could not save {INPUT_FILE}. Close it in Excel.")

    except SAPConnectionError as e:
        print(f"SAP Connection Error: {e}")
    except Exception as e:
        print(f"Unexpected Error: {e}")


if __name__ == "__main__":
    main()
