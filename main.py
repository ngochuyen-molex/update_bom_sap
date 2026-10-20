import os
import time
from connect import get_session, SAPConnectionError
from excel_utils import load_excel
from bom import *

INPUT_FILE = r"C:\Users\nguyeb3\Downloads\change bom source.xlsx"


def add_log(df, idx, message):
    print(message)
    df.at[idx, "LOG"] += f"{message}\n"

def save_progress(df):
    temp_file = INPUT_FILE.replace(".xlsx", "_result.xlsx")
    df.to_excel(temp_file, index=False, engine="openpyxl")
    os.replace(temp_file, INPUT_FILE)

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

        # Reset result columns for this run
        df["STATUS"] = ""
        df["MESSAGE"] = ""
        df["LOG"] = ""

        print(f"Loaded {len(df)} rows")

        # --------------------------------------------------
        # Loop through Excel rows
        # --------------------------------------------------
        for idx, row in df.iterrows():
            try:
                print("=" * 80)

                df.at[idx, "STATUS"] = "PROCESSING"
                df.at[idx, "MESSAGE"] = ""
                add_log(df, idx, f"Processing row {idx + 1}")

                # ------------------------------------------
                # Check blank row
                # ------------------------------------------
                if is_blank_row(row):
                    df.at[idx, "STATUS"] = "SKIPPED"
                    df.at[idx, "MESSAGE"] = "Blank row"
                    add_log(df, idx, f"Skipping blank row {idx + 1}")
                    continue

                # ------------------------------------------
                # Check if there is anything to update
                # ------------------------------------------
                add_log(df, idx, "About to check for changes")

                if not has_changes(row):
                    df.at[idx, "STATUS"] = "SKIPPED"
                    df.at[idx, "MESSAGE"] = "Component and Qty unchanged"
                    add_log(df, idx, "Skip - Component and Qty unchanged")
                    continue

                add_log(df, idx, "Changes detected")

                # ------------------------------------------
                # Open BOM
                # ------------------------------------------
                add_log(df, idx, "About to open BOM")
                add_log(df, idx, f"Excel Item value: {clean_text(row.get('Item'))!r}")

                open_bom(
                    session=session,
                    material=row["Parent Material"],
                    alt_bom=row["Alternative BOM"],
                    component=row["Component before"],
                    item=row.get("Item"))
               
                add_log(df, idx, "BOM opened successfully")


                # ------------------------------------------
                # Find matching item
                
                # ------------------------------------------
                add_log(df, idx, "About to find matching item")

                target_item = row.get("Item")
                if clean_item(target_item):
                    match_method = "item number"
                    search_value = clean_item(target_item)
                    add_log(
                        df,
                        idx,
                        f"Searching by {match_method}: {search_value}",
                    )
                    row_idx = find_item_row(
                        session=session,
                        target_item=target_item,
                    )
                else:
                    match_method = "Component before"
                    search_value = clean_component(row.get("Component before"))
                    add_log(
                        df,
                        idx,
                        f"Searching by {match_method}: {search_value}",
                    )
                    row_idx = find_component_row(
                        session=session,
                        target_component=row.get("Component before"),
                    )

                add_log(
                    df,
                    idx,
                    f"Matched SAP visible row {row_idx} using {match_method}",
                )

                # ------------------------------------------
                # Validate SAP current values
                # ------------------------------------------
                add_log(df, idx, "About to validate current BOM values")

                check_current_bom_status(
                    session=session,
                    row_idx=row_idx,
                    excel_row=row)

                add_log(df, idx, "Component before and Qty before matched")

                update_component(
                    session=session,
                    row_idx=row_idx,
                    component_after=row["Component after"])

                add_log(df, idx, "Component updated")

                update_quantity(
                    session=session,
                    row_idx=row_idx,
                    qty_after=row["BOM rate after"])

                add_log(df, idx, "Quantity entered")


                press_enter(session)
                press_enter(session)

                add_log(df, idx, "SAP changes confirmed")

                time.sleep(2)
                save_bom(session)

                df.at[idx, "STATUS"] = "SUCCESS"
                df.at[idx, "MESSAGE"] = "BOM saved successfully"
                add_log(df, idx, "BOM saved successfully")

            except Exception as e:
                error_message = str(e)

                df.at[idx, "STATUS"] = "FAILED"
                df.at[idx, "MESSAGE"] = error_message
                add_log(df, idx, f"FAILED row {idx + 1}: {error_message}")

            finally:
                try:
                    save_progress(df)
                    print(f"Excel progress saved after row {idx + 1}")

                except Exception as save_error:
                    print(f"Could not save Excel after row {idx + 1}: {save_error}")
                    print("Make sure the Excel file is closed.")

    except SAPConnectionError as e:
        print(f"SAP Connection Error: {e}")

    except Exception as e:
        print(f"Unexpected Error: {e}")


if __name__ == "__main__":
    main()