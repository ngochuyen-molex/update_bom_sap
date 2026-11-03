from connect import get_session, SAPConnectionError
from excel_utils import load_excel, save_excel
from bom import clean_component
from models import RoutingTask
from routing import read_routing_groups, process_group, slow_session

# Result file of create_main.py: only materials with a newly created BOM
INPUT_FILE = r"C:\Users\nguyeb3\Downloads\create_bom_final.xlsx"
RESULT_FILE = INPUT_FILE.replace(".xlsx", "_routing.xlsx")

OPERATION = "0020"
BOM_ROW = 8             # row of the new BOM in the btn[7] popup (from recording)


def build_tasks(df):
    # One RoutingTask per unique Parent Material with BOM STATUS = SUCCESS
    tasks = {}
    for idx, row in df.iterrows():
        if str(row.get("STATUS")).strip() != "SUCCESS":
            continue
        material = clean_component(row["Parent Material"])
        if material == "":
            continue
        task = tasks.setdefault(material, RoutingTask(material=material, operation=OPERATION))
        task.excel_rows.append(idx)
    return list(tasks.values())


def write_result(df, task):
    for idx in task.excel_rows:
        df.at[idx, "ROUTING STATUS"] = task.status
        df.at[idx, "ROUTING MESSAGE"] = task.message
        df.at[idx, "ROUTINGS DONE"] = ", ".join(task.done)


def main():
    try:
        session = slow_session(get_session())
        print(f"Connected to SAP system {session.Info.SystemName}")

        df = load_excel(INPUT_FILE)
        df["ROUTING STATUS"] = ""
        df["ROUTING MESSAGE"] = ""
        df["ROUTINGS DONE"] = ""

        tasks = build_tasks(df)
        print(f"{len(tasks)} materials to process")

        for task in tasks:
            print("=" * 80)
            try:
                read_routing_groups(session, task)

                for group in task.groups:
                    process_group(session, task, group, BOM_ROW)

                task.status = "SUCCESS"
                task.message = ""
            except Exception as e:
                task.status = "FAILED"
                task.message = str(e)
            print(f"{task.material}: {task.status} {task.message}")

            write_result(df, task)
            try:
                save_excel(df, RESULT_FILE)
            except PermissionError:
                print(f"Could not save {RESULT_FILE}. Close it in Excel.")

    except SAPConnectionError as e:
        print(f"SAP Connection Error: {e}")
    except Exception as e:
        print(f"Unexpected Error: {e}")


if __name__ == "__main__":
    main()
