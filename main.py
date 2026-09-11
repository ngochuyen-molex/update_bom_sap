import os
import time
from connect import get_session, SAPConnectionError
from excel_utils import load_excel
from bom import *

INPUT_FILE = r"C:\Users\nguyeb3\Downloads\bom_status_error.xlsx"

def main():
    try:
        # --------------------------------------------------
        # Connect to SAP
        # --------------------------------------------------
        session = get_session()

        print(f"Connected to SAP system {session.Info.SystemName}")

        open_cs02(session)

        # --------------------------------------------------
        # Load Excel
        # --------------------------------------------------
        df = load_excel(INPUT_FILE)

        print(f"Loaded {len(df)} rows")

        # --------------------------------------------------
        # Loop through Excel rows
        # --------------------------------------------------
        for idx, row in df.iterrows():
            try:
                print("=" * 80)

                if is_blank_row(row):
                    print(f"Skipping blank row {idx + 1}")
                    continue

                # ------------------------------------------
                # Check if there is anything to update
                # ------------------------------------------
                print(f"Processing row {idx + 1}")

                if not has_changes(row):
                    print("Skip - Component and Qty unchanged")
                    continue

                open_bom(
                    session=session,
                    material=row["Parent Material"],
                    alt_bom=row["Alternative BOM"],
                    bom_usage="1")

                # ------------------------------------------
                # Find matching Item
                # ------------------------------------------
                row_idx = find_item_row(
                    session=session,
                    target_item=row["Item"])

                print(f"Matched item at SAP row {row_idx}")

                # ------------------------------------------
                # Validate SAP current values
                # ------------------------------------------
                check_current_bom_status(
                    session=session,
                    row_idx=row_idx,
                    excel_row=row)

                print("Component before and Qty before matched")

                # ------------------------------------------
                # Update Component
                # ------------------------------------------
                update_component(
                    session=session,
                    row_idx=row_idx,
                    component_after=row["Component after"])

                print("Component updated")

                # ------------------------------------------
                # Update Qty
                # ------------------------------------------
                update_quantity(
                    session=session,
                    row_idx=row_idx,
                    qty_after=row["BOM rate after"])

                # press enter
                press_enter(session)
                press_enter(session)

                print("Quantity updated")

                time.sleep(2)

                save_bom(session)

                print("BOM saved successfully")

            except Exception as e:
                print(f"FAILED row {idx + 1}")
                print(str(e))
                continue

    except SAPConnectionError as e:
        print(f"SAP Connection Error: {e}")

    except Exception as e:
        print(f"Unexpected Error: {e}")


if __name__ == "__main__":
    main()