import pandas as pd
from connect import get_session, slow_from_args, SAPConnectionError
from excel_utils import load_excel, save_progress
from bom import clean_component
from models import Routing
from routing import process_routing
from routing_sf import load_routings

# Usage:
#   python routing_main.py [slow]         all routings from Snowflake -> RESULT_FILE
#   python routing_main.py retry [slow]   only STATUS = FAILED rows of RESULT_FILE, updated in place

# Result file of create_main.py: only materials with a newly created BOM
#gott solve this because retry and normal mode use different files
INPUT_FILE = r"C:\Users\nguyeb3\Downloads\create_bom_final_routing.xlsx"
RESULT_FILE = INPUT_FILE.replace(".xlsx", ".xlsx")

BOM_ROW = 8             # row of the new BOM in the btn[7] popup (from recording)
# Operations (0020 default, 0030 packaging) are set in routing.py


def get_materials(df):
    # Unique Parent Materials whose BOM was created (STATUS = SUCCESS)
    materials = []
    for _, row in df.iterrows():
        if str(row.get("STATUS")).strip() != "SUCCESS":
            continue
        material = clean_component(row["Parent Material"])
        if material and material not in materials:
            materials.append(material)
    return materials


def save_results(routings, missing):
    # One row per routing, plus materials with no routing in MLX_DATALAKE
    rows = [{
        "Material": r.material,
        "Group": r.group,
        "GrC": r.counter,
        "Description": r.description,
        "STATUS": r.status,
        "MESSAGE": r.message,
    } for r in routings]
    rows += [{"Material": m, "STATUS": "FAILED",
              "MESSAGE": "No routing found in MLX_DATALAKE"} for m in missing]
    try:
        pd.DataFrame(rows).to_excel(RESULT_FILE, index=False)
    except PermissionError:
        print(f"Could not save {RESULT_FILE}. Close it in Excel.")


def main():
    try:
        session = get_session(slow=slow_from_args())    # python routing_main.py slow
        print(f"Connected to SAP system {session.Info.SystemName}")

        materials = get_materials(load_excel(INPUT_FILE))
        routings = load_routings(materials)
        missing = sorted(set(materials) - {r.material for r in routings})
        print(f"{len(materials)} materials -> {len(routings)} routings")
        for m in missing:
            print(f"WARNING: no routing found for {m}")

        for routing in routings:
            try:
                process_routing(session, routing, BOM_ROW)
            except Exception as e:
                routing.status = "FAILED"
                routing.message = str(e)
            print(f"{routing.key}: {routing.status} {routing.message}")
            save_results(routings, missing)

    except SAPConnectionError as e:
        print(f"SAP Connection Error: {e}")
    except Exception as e:
        print(f"Unexpected Error: {e}")


def cell(row, column):
    value = row.get(column)
    return "" if pd.isna(value) else str(value).strip()


def retry_failed():
    # python routing_main.py retry [slow]
    # Reads RESULT_FILE (no Snowflake), reprocesses only routings with STATUS = FAILED,
    # and updates STATUS / MESSAGE in the same file.
    try:
        session = get_session(slow=slow_from_args())

        df = pd.read_excel(RESULT_FILE, dtype=str).fillna("")   # dtype=str keeps "0020", GrC as text
        # GrC per material + group, for the "no header overview" check
        sizes = df[df["Group"] != ""].groupby(["Material", "Group"]).size()

        todo = [idx for idx, row in df.iterrows() if cell(row, "STATUS") == "FAILED"]
        print(f"Loaded {len(df)} rows from {RESULT_FILE}, {len(todo)} FAILED to retry")

        for idx in todo:
            row = df.loc[idx]
            if cell(row, "Group") == "":
                # Material-only row ("No routing found in MLX_DATALAKE"): nothing to open in CA02
                print(f"{cell(row, 'Material')}: no routing, skipped ({cell(row, 'MESSAGE')})")
                continue

            routing = Routing(
                material=cell(row, "Material"), group=cell(row, "Group"),
                counter=cell(row, "GrC"), description=cell(row, "Description"),
                group_size=int(sizes.get((cell(row, "Material"), cell(row, "Group")), 1)))
            try:
                process_routing(session, routing, BOM_ROW)
            except Exception as e:
                routing.status = "FAILED"
                routing.message = str(e)
            print(f"{routing.key}: {routing.status} {routing.message}")
            save_results
            df.at[idx, "STATUS"] = routing.status
            df.at[idx, "MESSAGE"] = routing.message
            try:
                save_progress(RESULT_FILE, df)   # temp file + replace, so a crash can't corrupt it
            except PermissionError:
                print(f"Could not save {RESULT_FILE}. Close it in Excel.")

    except SAPConnectionError as e:
        print(f"SAP Connection Error: {e}")
    except Exception as e:
        print(f"Unexpected Error: {e}")


if __name__ == "__main__":
    import sys
    if "retry" in sys.argv:
        retry_failed()
    else:
        main()
