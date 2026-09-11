import os
from connect import get_session, SAPConnectionError
from excel_utils import load_excel
from bom import *
INPUT_FILE = r"C:\Users\nguyeb3\Downloads\change bom source.xlsx"

def main():
    try:
        session = get_session()
        open_cs02(session)
        df = load_excel(INPUT_FILE)
        print(f"Loaded {len(df)} rows")
        unique_df = (df.groupby(["Parent Material","Alternative BOM","New Base"],
                    dropna=False,as_index=False).first())

        print(f"found {len(unique_df)} unique material and alt bom combinations")

        for idx, row in unique_df.iterrows():
            try:
                print("=" * 80)
                if is_blank_row(row):
                    print(f"Skipping blank row {idx + 1}")
                    continue
                print(f"Processing row {idx + 1}")

                open_bom(
                    session=session,
                    material=row["Parent Material"],
                    alt_bom=row["Alternative BOM"],
                    bom_usage="1")
                
                open_bom_header(session)

                update_base_quantity(session,base_qty=row["New Base"])

                save_bom(session)

            except Exception as e:
                print(f"Error processing row {idx + 1}: {e}")

    except SAPConnectionError as e:
            print(f"SAP Connection Error: {e}")
    
    except Exception as e:
        print(f"Unexpected Error: {e}")

if __name__ == "__main__":
    main()

