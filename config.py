"""Settings shared by all scripts. Change paths and fixed SAP values here only."""

# ---------- Input / result files ----------
# One workbook for the chain create BOM -> assign components -> production version
CHAIN_FILE = r"C:\Users\nguyeb3\Downloads\CREATE_BOM.xlsx"
ALTBOM_FILE = r"C:\Users\nguyeb3\Downloads\md_assign_component.xlsx"       # routing_main.py altbom

CHANGE_BOM_FILE = r"C:\Users\nguyeb3\Downloads\CHANGE_BOM_list.xlsx"        # bom_update_main.py
BASE_QTY_FILE = r"C:\Users\nguyeb3\Downloads\create_bom_final.xlsx"         # update_baseqty.py
UNMERGE_FILE = r"C:\Users\nguyeb3\Downloads\CHANGE_BOM_demo.xlsx"           # unmerge_excel.py
WORKCENTER_FILE = r"C:\Users\nguyeb3\Downloads\Set MU APS.xlsx"             # update_workcenter_efficiency.py

# ---------- SAP ----------
SAP_CONNECTION_NAME = "M04 [SSO]"     # entry name in SAP Logon
PLANT = "1901"
BOM_USAGE = "1"

# ---------- Snowflake ----------
SNOWFLAKE_CONNECTION = "default"      # name in connections.toml

# ---------- Run ----------
SAVE_EVERY = 5                        # save the result sheet after every N processed entries
