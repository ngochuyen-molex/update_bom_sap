import pandas as pd
import snowflake.connector
from models import Routing

# One row per routing (material + group + group counter) at plant 1901.
# No PLAS/PLPO/CRHD join: those repeat each routing once per operation / work center.
SQL_ROUTINGS = """
    SELECT
        LTRIM(mp.MATNR, '0') AS MATERIAL,
        TRIM(pk.PLNNR) AS ROUTING_GROUP,
        LTRIM(pk.PLNAL, '0') AS GROUP_COUNTER,
        pk.KTEXT AS ROUTING_DESCRIPTION
    FROM MLX_DATALAKE.MLX_SAP_ECC.MAPL mp
    JOIN MLX_DATALAKE.MLX_SAP_ECC.PLKO pk
        ON mp.MANDT = pk.MANDT
        AND mp.PLNTY = pk.PLNTY
        AND mp.PLNNR = pk.PLNNR
        AND mp.PLNAL = pk.PLNAL
    WHERE pk.WERKS = '1901'
        AND mp.PLNTY = 'N'          -- standard routing (CA02); Q = rate routing (CA22)
        AND TRIM(pk.LOEKZ) = ''
        AND TRIM(mp.LOEKZ) = ''
        AND pk.DW_SOFT_DELETE_FLAG = 0
        AND mp.DW_SOFT_DELETE_FLAG = 0
        AND LTRIM(mp.MATNR, '0') IN ({placeholders})
    -- PLKO can hold several change states of a routing: keep the latest one
    QUALIFY ROW_NUMBER() OVER (
        PARTITION BY mp.MATNR, pk.PLNNR, pk.PLNAL
        ORDER BY pk.ZAEHL DESC) = 1
    ORDER BY MATERIAL, ROUTING_GROUP, GROUP_COUNTER
"""


def load_routings(materials, plant="1901"):
    # Returns a flat list of Routing, one per material + group + GrC
    materials = sorted({str(m).lstrip("0") for m in materials})
    if not materials:
        return []

    sql = SQL_ROUTINGS.format(placeholders=", ".join(["%s"] * len(materials)))
    conn = snowflake.connector.connect(connection_name="default")
    try:
        cur = conn.cursor()
        cur.execute(sql, materials)
        df = pd.DataFrame(cur.fetchall(), columns=[c[0] for c in cur.description])
    finally:
        conn.close()

    # GrC per material + group: SAP skips the header overview when a group has only one
    sizes = df.groupby(["MATERIAL", "ROUTING_GROUP"])["GROUP_COUNTER"].transform("count")

    return [
        Routing(
            material=row.MATERIAL,
            group=row.ROUTING_GROUP,
            counter=row.GROUP_COUNTER or "0",
            description=(row.ROUTING_DESCRIPTION or "").strip(),
            plant=plant,
            group_size=int(size))
        for row, size in zip(df.itertuples(index=False), sizes)
    ]
