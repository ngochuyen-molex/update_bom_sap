import os
import pandas as pd
from connect import SAPConnectionError, get_session
from excel_utils import load_excel

INPUT_FILE = r"C:\Users\nguyeb3\Downloads\Set MU APS.xlsx"
REQUIRED_COLUMNS = ("Work_Center", "ME", "MU")


def excel_value(value):
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def save_progress(df):
    temp_file = INPUT_FILE.replace(".xlsx", "_temp.xlsx")
    df.to_excel(temp_file, index=False, engine="openpyxl")
    os.replace(temp_file, INPUT_FILE)


def update_workcenter(session, workcenter, efficiency, utilization):
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nCR02"
    session.findById("wnd[0]").sendVKey(0)

    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/usr/ctxt[1]").text = workcenter
    session.findById("wnd[0]/usr/ctxt[1]").setFocus()
    session.findById("wnd[0]/usr/ctxt[1]").caretPosition = len(workcenter)
    session.findById("wnd[0]").sendVKey(0)

    session.findById(
        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpKAUE"
    ).select()
    session.findById(
        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/"
        "tabpKAUE/ssub/2/3/sub/2/3/2/btn[5]"
    ).press()

    if efficiency:
        session.findById("wnd[1]/usr/txt[3]").text = efficiency
    if utilization:
        utilization_field = session.findById("wnd[1]/usr/txt[5]")
        utilization_field.text = utilization
        utilization_field.setFocus()
        utilization_field.caretPosition = len(utilization)
    session.findById("wnd[1]/tbar[0]/btn[0]").press()


def main():
    try:
        df = load_excel(INPUT_FILE)
        missing_columns = [column for column in REQUIRED_COLUMNS if column not in df.columns]
        if missing_columns:
            raise ValueError(
                f"Missing required Excel column(s): {', '.join(missing_columns)}"
            )

        df["STATUS"] = ""
        df["MESSAGE"] = ""

        session = get_session()
        print(f"Connected to SAP system {session.Info.SystemName}")

        print(f"Loaded {len(df)} rows")
        for idx, row in df.iterrows():
            try:
                workcenter = excel_value(row["Work_Center"])
                efficiency = excel_value(row["ME"])
                utilization = excel_value(row["MU"])

                if not workcenter and not efficiency and not utilization:
                    df.at[idx, "STATUS"] = "SKIPPED"
                    df.at[idx, "MESSAGE"] = "Blank row"
                    continue

                if not efficiency and not utilization:
                    df.at[idx, "STATUS"] = "SKIPPED"
                    df.at[idx, "MESSAGE"] = "Both ME and MU are blank; no update performed"
                    continue

                if not workcenter:
                    raise ValueError("Workcenter is blank.")

                print(f"Updating row {idx + 1}: {workcenter}")
                update_workcenter(session, workcenter, efficiency, utilization)
                session.findById("wnd[0]/tbar[0]/btn[11]").press()
                df.at[idx, "STATUS"] = "SUCCESS"
                updated_fields = []
                unchanged_fields = []
                for name, value in (("ME", efficiency), ("MU", utilization)):
                    (updated_fields if value else unchanged_fields).append(name)
                message = f"Updated {', '.join(updated_fields)}"
                if unchanged_fields:
                    message += f"; left blank {', '.join(unchanged_fields)} unchanged"
                df.at[idx, "MESSAGE"] = message
            except Exception as error:
                df.at[idx, "STATUS"] = "FAILED"
                df.at[idx, "MESSAGE"] = str(error)
                print(f"Failed row {idx + 1}: {error}")
            finally:
                save_progress(df)
                print(f"Excel progress saved after row {idx + 1}")

    except SAPConnectionError as error:
        print(f"SAP Connection Error: {error}")
    except Exception as error:
        print(f"Unexpected Error: {error}")


if __name__ == "__main__":
    main()
