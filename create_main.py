import time
from connect import get_session, SAPConnectionError
from excel_utils import load_excel, save_excel
from bom import clean_text, clean_number, save_bom
from bom_create import open_cs01, open_bom_cs01, enter_items, get_created_alt_bom
from models import BomHeader, BomItem

# Output of unmerge_excel.py (run that first)
INPUT_FILE = r"C:\Users\nguyeb3\Downloads\create_bom_demo_fix_unit.xlsx"
RESULT_FILE = INPUT_FILE.replace(".xlsx", "_result.xlsx")

UNIT_COLUMN = "Unit after"   # TODO: change to the real unit column name in your Excel


def build_headers(df):
    # One BomHeader per Parent Material + Alternative BOM, items in Excel order
    headers = {}
    for idx, row in df.iterrows():
        component = clean_text(row.get("Component after"))
        if component == "":
            continue
        key = (clean_text(row["Parent Material"]), clean_text(row["Alternative BOM"]))
        if key not in headers:
            headers[key] = BomHeader(
                parent=key[0],
                alt_bom=key[1],
                base_qty=clean_number(row.get("Base Qty after")))
        headers[key].items.append(BomItem(
            excel_row=idx,
            component=component,
            quantity=row["BOM rate after"],
            unit=clean_text(row.get(UNIT_COLUMN))))
    return list(headers.values())


def write_result(df, header):
    for idx in header.excel_rows:
        df.at[idx, "STATUS"] = header.status
        df.at[idx, "MESSAGE"] = header.message
        df.at[idx, "NEW ALT BOM"] = header.created_alt_bom


def main():
    try:
        session = get_session()
        print(f"Connected to SAP system {session.Info.SystemName}")

        df = load_excel(INPUT_FILE)
        df["STATUS"] = ""
        df["MESSAGE"] = ""
        df["NEW ALT BOM"] = ""

        headers = build_headers(df)
        print(f"Loaded {len(df)} rows -> {len(headers)} BOMs to create")

        for header in headers:
            print("=" * 80)
            try:
                open_cs01(session)
                open_bom_cs01(session, material=header.parent,
                              bom_usage=header.usage, plant=header.plant)

                header.created_alt_bom = get_created_alt_bom(session)
                enter_items(session, header.items)
                press_enter_twice(session)

                time.sleep(2)
                save_bom(session)

                header.status = "SUCCESS"
                header.message = session.findById("wnd[0]/sbar").text
            except Exception as e:
                header.status = "FAILED"
                header.message = str(e)
            print(f"{header.parent}: {header.status} {header.message}")

            write_result(df, header)
            try: 
                save_excel(df, RESULT_FILE)
            except PermissionError:
                print(f"Could not save {RESULT_FILE}. Close it in Excel.")

    except SAPConnectionError as e:
        print(f"SAP Connection Error: {e}")
    except Exception as e:
        print(f"Unexpected Error: {e}")


def press_enter_twice(session):
    # Confirm item entries / warnings, same as main.py
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]").sendVKey(0)


if __name__ == "__main__":
    main()
