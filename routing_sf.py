import pandas as pd
import snowflake.connector

# One row per routing (MATNR + group + group counter).
# No PLPO/CRHD join: that would repeat each routing once per operation / work center.
SQL_ROUTINGS = """
    SELECT DISTINCT
        LTRIM(mp.MATNR, '0') AS MATERIAL,
        pk.PLNNR AS ROUTING_GROUP,
        LTRIM(pk.PLNAL, '0') AS GROUP_COUNTER,
        pk.KTEXT AS ROUTING_DESCRIPTION
    FROM MLX_DATALAKE.MLX_SAP_ECC.MAPL mp
    JOIN MLX_DATALAKE.MLX_SAP_ECC.PLKO pk
        ON mp.MANDT = pk.MANDT
        AND mp.PLNTY = pk.PLNTY
        AND mp.PLNNR = pk.PLNNR
        AND mp.PLNAL = pk.PLNAL
    WHERE pk.WERKS = '1901'
        AND TRIM(pk.LOEKZ) = ''
        AND TRIM(mp.LOEKZ) = ''
        AND pk.DW_SOFT_DELETE_FLAG = 0
        AND mp.DW_SOFT_DELETE_FLAG = 0
        AND LTRIM(mp.MATNR, '0') IN ({placeholders})
    ORDER BY MATERIAL, ROUTING_GROUP, GROUP_COUNTER
"""


def load_routings(materials):
    # Returns {material: {group: [counters]}} from MLX_DATALAKE
    materials = sorted({str(m).lstrip("0") for m in materials})
    if not materials:
        return {}

    sql = SQL_ROUTINGS.format(placeholders=", ".join(["%s"] * len(materials)))
    conn = snowflake.connector.connect(connection_name="default")
    try:
        cur = conn.cursor()
        cur.execute(sql, materials)
        df = pd.DataFrame(cur.fetchall(), columns=[c[0] for c in cur.description])
    finally:
        conn.close()

    result = {}
    for row in df.itertuples(index=False):
        groups = result.setdefault(row.MATERIAL, {})
        groups.setdefault(row.ROUTING_GROUP.strip(), []).append(row.GROUP_COUNTER or "0")
    return result


def compare_with_sap(task, sf_groups):
    # Messages for differences between SAP popup and Snowflake (empty list = match)
    sap = {f"{g.group}/{c}" for g in task.groups for c in g.counters}
    sf = {f"{group}/{c}" for group, counters in sf_groups.items() for c in counters}
    messages = []
    if sap - sf:
        messages.append(f"Only in SAP: {sorted(sap - sf)}")
    if sf - sap:
        messages.append(f"Only in Snowflake: {sorted(sf - sap)}")
    return messages
