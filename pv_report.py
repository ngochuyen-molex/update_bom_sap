"""Read-only report: routings + production versions per material (Snowflake only, no SAP).

    python pv_report.py "<input.xlsx>"         -> <input>_pv_report.xlsx next to it

Input: first sheet, column "Material". The input file is not changed.

Output sheets:
    Plan          one row per active routing x active alt BOM:
                  EXISTS <version> (same group + GrC + alt BOM) / TO CREATE / SKIP <reason>
    Routings      every routing of the materials (also deleted / trial / blocked material)
    Prod Versions existing production versions (MKAL, not deleted)
    Materials     per material: routings, alt BOMs, versions, status

Snowflake lags SAP by about 20 minutes: changes of the last minutes are not in it yet.
"""
import os
import sys
import pandas as pd
import snowflake.connector
from config import PLANT, SNOWFLAKE_CONNECTION
from routing_sf import SQL_ROUTINGS, SQL_ALT_BOMS, run_query, clean_material
from routing import is_trial
from production_version import version_text

MATERIAL_COLUMN = "Material"

# Existing production versions at the plant. DW_SOFT_DELETE_FLAG = 1 = deleted in SAP.
SQL_PROD_VERSIONS = """
    SELECT
        LTRIM(MATNR, '0') AS MATERIAL,
        VERID AS VERSION,
        COALESCE(TRIM(TEXT1), '') AS VERSION_TEXT,
        TRIM(PLNTY) AS ROUTING_TYPE,
        TRIM(PLNNR) AS ROUTING_GROUP,
        LTRIM(ALNAL, '0') AS GROUP_COUNTER,
        LTRIM(STLAL, '0') AS ALT_BOM,
        TRIM(STLAN) AS BOM_USAGE,
        ADATU AS VALID_FROM,
        BDATU AS VALID_TO,
        COALESCE(TRIM(MKSP), '') AS LOCKED
    FROM MLX_DATALAKE.MLX_SAP_ECC.MKAL
    WHERE WERKS = '1901'
        AND DW_SOFT_DELETE_FLAG = 0
        AND LTRIM(MATNR, '0') IN ({placeholders})
    ORDER BY MATERIAL, VERSION
"""

# Material status at the plant (E3 = not active in this plant, LVORM = flagged for deletion)
SQL_MATERIAL_STATUS = """
    SELECT
        LTRIM(MATNR, '0') AS MATERIAL,
        COALESCE(TRIM(MMSTA), '') AS MATERIAL_STATUS,
        COALESCE(TRIM(LVORM), '') AS MATERIAL_DELETED
    FROM MLX_DATALAKE.MLX_SAP_ECC.MARC
    WHERE WERKS = '1901'
        AND DW_SOFT_DELETE_FLAG = 0
        AND LTRIM(MATNR, '0') IN ({placeholders})
"""


def read_materials(path):
    df = pd.read_excel(path, sheet_name=0, dtype=str)
    if MATERIAL_COLUMN not in df.columns:
        raise ValueError(f"Column '{MATERIAL_COLUMN}' not found in {path}. Columns: {list(df.columns)}")
    materials = []
    for value in df[MATERIAL_COLUMN].fillna(""):
        material = clean_material(value)
        if material and material not in materials:
            materials.append(material)
    return materials


def load(materials):
    conn = snowflake.connector.connect(connection_name=SNOWFLAKE_CONNECTION)
    try:
        print("Snowflake: routings");            routings = run_query(conn, SQL_ROUTINGS, materials)
        print("Snowflake: alt BOMs");            alt_boms = run_query(conn, SQL_ALT_BOMS, materials)
        print("Snowflake: production versions"); versions = run_query(conn, SQL_PROD_VERSIONS, materials)
        print("Snowflake: material status");     status = run_query(conn, SQL_MATERIAL_STATUS, materials)
    finally:
        conn.close()
    for df in (routings, alt_boms, versions):
        for col in ("GROUP_COUNTER", "ALT_BOM"):
            if col in df.columns:
                df[col] = df[col].fillna("").replace("", "0")   # GrC / alt BOM 0 -> "0", same key everywhere
    return routings, alt_boms, versions, status


def routing_skip_reason(r, material_status):
    status, deleted = material_status.get(r.MATERIAL, ("", ""))
    if status == "E3":
        return "Material status E3 (not active in this plant)"
    if deleted == "X":
        return "Material flagged for deletion"
    if r.ROUTING_DELETED == "X":
        return "Deleted routing (DELKZ)"
    if is_trial(r.ROUTING_DESCRIPTION or ""):
        return "Trial routing"
    return ""


def build(materials, routings, alt_boms, versions, status):
    material_status = {r.MATERIAL: (r.MATERIAL_STATUS, r.MATERIAL_DELETED) for r in status.itertuples(index=False)}
    boms = {}
    for b in alt_boms.itertuples(index=False):
        boms.setdefault(b.MATERIAL, []).append((b.ALT_BOM, b.ALT_BOM_TEXT))
    existing = {}
    for v in versions.itertuples(index=False):
        existing.setdefault((v.MATERIAL, v.ROUTING_GROUP, v.GROUP_COUNTER, v.ALT_BOM), []).append(v.VERSION)

    routing_rows, plan_rows = [], []
    for r in routings.itertuples(index=False):
        reason = routing_skip_reason(r, material_status)
        routing_rows.append({
            "Material": r.MATERIAL, "Group": r.ROUTING_GROUP, "GrC": r.GROUP_COUNTER,
            "Description": r.ROUTING_DESCRIPTION, "Deleted (DELKZ)": r.ROUTING_DELETED,
            "Material Status": material_status.get(r.MATERIAL, ("", ""))[0],
            "Active": "N" if reason else "Y", "Reason": reason})
        material_boms = sorted(boms.get(r.MATERIAL, []), key=lambda b: int(b[0]) if b[0].isdigit() else 0)
        if not material_boms:
            plan_rows.append({"Material": r.MATERIAL, "Group": r.ROUTING_GROUP, "GrC": r.GROUP_COUNTER,
                              "Description": r.ROUTING_DESCRIPTION, "Alt BOM": "", "Alt BOM Text": "",
                              "Planned Text": "", "Existing Version": "",
                              "Action": f"SKIP {reason or 'No active alt BOM (usage 1, status 01)'}"})
            continue
        for alt_bom, alt_text in material_boms:
            found = existing.get((r.MATERIAL, r.ROUTING_GROUP, r.GROUP_COUNTER, alt_bom), [])
            try:
                text = version_text(r.ROUTING_DESCRIPTION or "", alt_bom, r.MATERIAL)
            except ValueError as e:
                text = f"(too long) {e}"
            if reason:
                action = f"SKIP {reason}"
            elif is_trial(alt_text):
                action = "SKIP Trial alt BOM"
            elif found:
                action = f"EXISTS {', '.join(found)}"
            elif text.startswith("(too long)"):
                action = "SKIP Text longer than 40"
            else:
                action = "TO CREATE"
            plan_rows.append({"Material": r.MATERIAL, "Group": r.ROUTING_GROUP, "GrC": r.GROUP_COUNTER,
                              "Description": r.ROUTING_DESCRIPTION, "Alt BOM": alt_bom,
                              "Alt BOM Text": alt_text, "Planned Text": text,
                              "Existing Version": ", ".join(found), "Action": action})

    plan = pd.DataFrame(plan_rows, columns=["Material", "Group", "GrC", "Description", "Alt BOM",
                                            "Alt BOM Text", "Planned Text", "Existing Version", "Action"])
    material_rows = []
    for m in materials:
        st, dl = material_status.get(m, ("", ""))
        mp = plan[plan["Material"] == m]
        material_rows.append({
            "Material": m, "Material Status": st, "Material Deleted": dl,
            "Routings": int((routings["MATERIAL"] == m).sum()) if not routings.empty else 0,
            "Active Routings": sum(1 for x in routing_rows if x["Material"] == m and x["Active"] == "Y"),
            "Active Alt BOMs": len(boms.get(m, [])),
            "Existing Versions": int((versions["MATERIAL"] == m).sum()) if not versions.empty else 0,
            "To Create": int((mp["Action"] == "TO CREATE").sum()),
            "Exists": int(mp["Action"].str.startswith("EXISTS").sum()),
            "Note": "" if m in set(routings["MATERIAL"]) else "No routing in MLX_DATALAKE"})

    version_sheet = versions.rename(columns={
        "MATERIAL": "Material", "VERSION": "Version", "VERSION_TEXT": "Text", "ROUTING_TYPE": "Routing Type",
        "ROUTING_GROUP": "Group", "GROUP_COUNTER": "GrC", "ALT_BOM": "Alt BOM", "BOM_USAGE": "BOM Usage",
        "VALID_FROM": "Valid From", "VALID_TO": "Valid To", "LOCKED": "Locked"})
    return plan, pd.DataFrame(routing_rows), version_sheet, pd.DataFrame(material_rows)


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return
    path = sys.argv[1]
    out = os.path.splitext(path)[0] + "_pv_report.xlsx"
    materials = read_materials(path)
    print(f"{len(materials)} materials from {path}")
    plan, routing_sheet, version_sheet, material_sheet = build(materials, *load(materials))
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        plan.to_excel(writer, sheet_name="Plan", index=False)
        routing_sheet.to_excel(writer, sheet_name="Routings", index=False)
        version_sheet.to_excel(writer, sheet_name="Prod Versions", index=False)
        material_sheet.to_excel(writer, sheet_name="Materials", index=False)
    print(plan["Action"].str.split(" ").str[0].value_counts().to_string())
    print(f"Written: {out}")


if __name__ == "__main__":
    main()
