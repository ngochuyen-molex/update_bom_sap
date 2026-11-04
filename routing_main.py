import pandas as pd
from connect import get_session, slow_from_args, SAPConnectionError
from excel_utils import load_excel
from bom import clean_component
from routing import process_routing
from routing_sf import load_routings

# Result file of create_main.py: only materials with a newly created BOM
INPUT_FILE = r"C:\Users\nguyeb3\Downloads\create_bom_final.xlsx"
RESULT_FILE = INPUT_FILE.replace(".xlsx", "_routing.xlsx")

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


if __name__ == "__main__":
    main()
