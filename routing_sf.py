import pandas as pd
import snowflake.connector
from models import Routing

BATCH_SIZE = 500            # materials per query (long IN lists are split)

# One row per routing (material + group + group counter) at plant 1901.
# No PLAS/PLPO/CRHD join: those repeat each routing once per operation / work center.
SQL_ROUTINGS = """
    SELECT
        LTRIM(mp.MATNR, '0') AS MATERIAL,
        TRIM(pk.PLNNR) AS ROUTING_GROUP,
        LTRIM(pk.PLNAL, '0') AS GROUP_COUNTER,
        pk.KTEXT AS ROUTING_DESCRIPTION,
        COALESCE(TRIM(pk.DELKZ), '') AS ROUTING_DELETED
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

# Active alt BOMs: usage 1, plant 1901, status 01 on the latest change state (STKOZ)
SQL_ALT_BOMS = """
    SELECT
        LTRIM(m.MATNR, '0') AS MATERIAL,
        LTRIM(m.STLAL, '0') AS ALT_BOM,
        COALESCE(TRIM(s.STKTX), '') AS ALT_BOM_TEXT
    FROM MLX_DATALAKE.MLX_SAP_ECC.MAST m
    JOIN MLX_DATALAKE.MLX_SAP_ECC.STKO s
        ON m.MANDT = s.MANDT AND s.STLTY = 'M' AND m.STLNR = s.STLNR AND m.STLAL = s.STLAL
    WHERE m.STLAN = '1' AND m.WERKS = '1901'
        AND m.DW_SOFT_DELETE_FLAG = 0 AND s.DW_SOFT_DELETE_FLAG = 0
        AND LTRIM(m.MATNR, '0') IN ({placeholders})
    QUALIFY ROW_NUMBER() OVER (PARTITION BY m.MATNR, m.STLAL ORDER BY s.STKOZ DESC) = 1
        AND s.STLST = '01'
"""

# Work center of operation 0030 per routing (latest change state, not deleted)
SQL_OP_0030_WC = """
    SELECT
        LTRIM(mp.MATNR, '0') AS MATERIAL,
        TRIM(pk.PLNNR) AS ROUTING_GROUP,
        LTRIM(pk.PLNAL, '0') AS GROUP_COUNTER,
        TRIM(cr.ARBPL) AS WORK_CENTER
    FROM MLX_DATALAKE.MLX_SAP_ECC.MAPL mp
    JOIN MLX_DATALAKE.MLX_SAP_ECC.PLKO pk
        ON mp.MANDT = pk.MANDT AND mp.PLNTY = pk.PLNTY AND mp.PLNNR = pk.PLNNR AND mp.PLNAL = pk.PLNAL
    JOIN MLX_DATALAKE.MLX_SAP_ECC.PLAS ps
        ON pk.MANDT = ps.MANDT AND pk.PLNTY = ps.PLNTY AND pk.PLNNR = ps.PLNNR AND pk.PLNAL = ps.PLNAL
        AND ps.LOEKZ = ' ' AND ps.DW_SOFT_DELETE_FLAG = 0
    JOIN MLX_DATALAKE.MLX_SAP_ECC.PLPO po
        ON ps.MANDT = po.MANDT AND ps.PLNTY = po.PLNTY AND ps.PLNNR = po.PLNNR AND ps.PLNKN = po.PLNKN
        AND po.DW_SOFT_DELETE_FLAG = 0
    JOIN MLX_DATALAKE.MLX_SAP_ECC.CRHD cr
        ON po.ARBID = cr.OBJID AND po.MANDT = cr.MANDT
    WHERE pk.WERKS = '1901' AND mp.PLNTY = 'N'
        AND TRIM(pk.LOEKZ) = '' AND TRIM(mp.LOEKZ) = ''
        AND pk.DW_SOFT_DELETE_FLAG = 0 AND mp.DW_SOFT_DELETE_FLAG = 0
        AND po.VORNR = '0030'
        AND LTRIM(mp.MATNR, '0') IN ({placeholders})
    QUALIFY ROW_NUMBER() OVER (PARTITION BY mp.MATNR, pk.PLNNR, pk.PLNAL ORDER BY po.ZAEHL DESC) = 1
"""

# Component allocation: one row per routing x alt BOM x BOM item, OPERATION empty = not assigned.
# Material list is used twice (routings + alt BOMs).
SQL_ALLOCATION = """
    WITH routings AS (
        SELECT mp.MATNR, pk.PLNNR, pk.PLNAL
        FROM MLX_DATALAKE.MLX_SAP_ECC.MAPL mp
        JOIN MLX_DATALAKE.MLX_SAP_ECC.PLKO pk
            ON mp.MANDT = pk.MANDT AND mp.PLNTY = pk.PLNTY AND mp.PLNNR = pk.PLNNR AND mp.PLNAL = pk.PLNAL
        WHERE pk.WERKS = '1901' AND mp.PLNTY = 'N'
            AND TRIM(pk.LOEKZ) = '' AND TRIM(mp.LOEKZ) = ''
            AND pk.DW_SOFT_DELETE_FLAG = 0 AND mp.DW_SOFT_DELETE_FLAG = 0
            AND LTRIM(mp.MATNR, '0') IN ({placeholders})
        QUALIFY ROW_NUMBER() OVER (PARTITION BY mp.MATNR, pk.PLNNR, pk.PLNAL ORDER BY pk.ZAEHL DESC) = 1
    ),
    alt_boms AS (
        SELECT m.MATNR, m.STLNR, m.STLAL
        FROM MLX_DATALAKE.MLX_SAP_ECC.MAST m
        JOIN MLX_DATALAKE.MLX_SAP_ECC.STKO s
            ON m.MANDT = s.MANDT AND s.STLTY = 'M' AND m.STLNR = s.STLNR AND m.STLAL = s.STLAL
        WHERE m.STLAN = '1' AND m.WERKS = '1901'
            AND m.DW_SOFT_DELETE_FLAG = 0 AND s.DW_SOFT_DELETE_FLAG = 0
            AND LTRIM(m.MATNR, '0') IN ({placeholders})
        QUALIFY ROW_NUMBER() OVER (PARTITION BY m.MATNR, m.STLAL ORDER BY s.STKOZ DESC) = 1
            AND s.STLST = '01'
    ),
    items AS (
        SELECT a.MATNR, a.STLNR, a.STLAL, st.STLKN, p.POSNR, p.IDNRK
        FROM alt_boms a
        JOIN MLX_DATALAKE.MLX_SAP_ECC.STAS st
            ON st.STLTY = 'M' AND st.STLNR = a.STLNR AND st.STLAL = a.STLAL AND st.DW_SOFT_DELETE_FLAG = 0
        JOIN MLX_DATALAKE.MLX_SAP_ECC.STPO p
            ON p.STLTY = 'M' AND p.STLNR = st.STLNR AND p.STLKN = st.STLKN AND p.DW_SOFT_DELETE_FLAG = 0
        QUALIFY ROW_NUMBER() OVER (PARTITION BY a.MATNR, a.STLAL, st.STLKN ORDER BY st.STASZ DESC, p.STPOZ DESC) = 1
            AND TRIM(st.LKENZ) = '' AND TRIM(p.LKENZ) = ''
    ),
    alloc AS (
        SELECT z.PLNNR, z.PLNAL, z.STLNR, z.STLAL, z.STLKN, po.VORNR
        FROM MLX_DATALAKE.MLX_SAP_ECC.PLMZ z
        JOIN MLX_DATALAKE.MLX_SAP_ECC.PLPO po
            ON po.MANDT = z.MANDT AND po.PLNTY = z.PLNTY AND po.PLNNR = z.PLNNR AND po.PLNKN = z.PLNKN
            AND po.DW_SOFT_DELETE_FLAG = 0
        WHERE z.PLNTY = 'N' AND z.STLTY = 'M' AND z.DW_SOFT_DELETE_FLAG = 0
            AND z.PLNNR IN (SELECT PLNNR FROM routings)
        QUALIFY ROW_NUMBER() OVER (PARTITION BY z.PLNNR, z.ZUONR ORDER BY z.ZAEHL DESC, po.ZAEHL DESC) = 1
            AND TRIM(z.LOEKZ) = ''
    )
    SELECT
        LTRIM(r.MATNR, '0') AS MATERIAL,
        TRIM(r.PLNNR) AS ROUTING_GROUP,
        LTRIM(r.PLNAL, '0') AS GROUP_COUNTER,
        LTRIM(i.STLAL, '0') AS ALT_BOM,
        i.POSNR AS ITEM,
        LTRIM(i.IDNRK, '0') AS COMPONENT,
        COALESCE(mk.MAKTX, '') AS COMPONENT_DESCRIPTION,
        COALESCE(a.VORNR, '') AS OPERATION
    FROM routings r
    JOIN items i ON i.MATNR = r.MATNR
    LEFT JOIN alloc a
        ON a.PLNNR = r.PLNNR AND a.PLNAL = r.PLNAL
        AND a.STLNR = i.STLNR AND a.STLAL = i.STLAL AND a.STLKN = i.STLKN
    LEFT JOIN MLX_DATALAKE.MLX_SAP_ECC.MAKT mk
        ON mk.MATNR = i.IDNRK AND mk.SPRAS = 'E' AND mk.DW_SOFT_DELETE_FLAG = 0
"""


def clean_material(value):
    # 5052760489 / "005052760489" / 5052760489.0 -> "5052760489"
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text.lstrip("0")


def run_query(conn, sql, materials, repeat=1):
    # Runs sql for all materials in batches; {placeholders} appears `repeat` times
    frames = []
    for start in range(0, len(materials), BATCH_SIZE):
        batch = materials[start:start + BATCH_SIZE]
        cur = conn.cursor()
        cur.execute(sql.format(placeholders=", ".join(["%s"] * len(batch))), batch * repeat)
        frames.append(pd.DataFrame(cur.fetchall(), columns=[c[0] for c in cur.description]))
        print(f"  query batch {start // BATCH_SIZE + 1}: {len(frames[-1])} rows")
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def load_routings(materials, plant="1901"):
    # Returns a flat list of Routing, one per material + group + GrC
    materials = sorted({clean_material(m) for m in materials if clean_material(m)})
    if not materials:
        return []

    conn = snowflake.connector.connect(connection_name="default")
    try:
        df = run_query(conn, SQL_ROUTINGS, materials)
    finally:
        conn.close()
    if df.empty:
        return []

    # GrC per material + group: SAP skips the header overview when a group has only one
    sizes = df.groupby(["MATERIAL", "ROUTING_GROUP"])["GROUP_COUNTER"].transform("count")

    return [
        Routing(
            material=row.MATERIAL,
            group=row.ROUTING_GROUP,
            counter=row.GROUP_COUNTER or "0",
            description=(row.ROUTING_DESCRIPTION or "").strip(),
            plant=plant,
            group_size=int(size),
            deleted=row.ROUTING_DELETED == "X")
        for row, size in zip(df.itertuples(index=False), sizes)
    ]


def load_assignment_worklist(materials, plant="1901"):
    """Component assignment work list: one Routing per routing x active alt BOM.

    Returns (entries, items):
      entries  list of Routing with alt_bom, alt_bom_text, bom_count, op_0030_wc,
               unassigned, total_items filled in
      items    DataFrame, one row per routing x alt BOM x BOM item:
               MATERIAL, ROUTING_GROUP, GROUP_COUNTER, ALT_BOM, ITEM, COMPONENT,
               COMPONENT_DESCRIPTION, OPERATION ("" = not assigned)
    """
    materials = sorted({clean_material(m) for m in materials if clean_material(m)})
    if not materials:
        return [], pd.DataFrame()

    conn = snowflake.connector.connect(connection_name="default")
    try:
        print("Snowflake: routings");        routings = run_query(conn, SQL_ROUTINGS, materials)
        print("Snowflake: alt BOMs");        alt_boms = run_query(conn, SQL_ALT_BOMS, materials)
        print("Snowflake: op 0030 WC");      wcs = run_query(conn, SQL_OP_0030_WC, materials)
        print("Snowflake: allocation");      items = run_query(conn, SQL_ALLOCATION, materials, repeat=2)
    finally:
        conn.close()
    if routings.empty:
        return [], items

    group_sizes = routings.groupby(["MATERIAL", "ROUTING_GROUP"]).size().to_dict()
    boms_by_material = {}
    for row in alt_boms.itertuples(index=False):
        boms_by_material.setdefault(row.MATERIAL, []).append((row.ALT_BOM, row.ALT_BOM_TEXT))
    wc_by_routing = {(r.MATERIAL, r.ROUTING_GROUP, r.GROUP_COUNTER): r.WORK_CENTER
                     for r in wcs.itertuples(index=False)}

    counts = {}
    if not items.empty:
        grouped = items.groupby(["MATERIAL", "ROUTING_GROUP", "GROUP_COUNTER", "ALT_BOM"])
        for key, group in grouped:
            counts[key] = (int((group["OPERATION"] == "").sum()), len(group))

    entries = []
    for r in routings.itertuples(index=False):
        counter = r.GROUP_COUNTER or "0"
        boms = sorted(boms_by_material.get(r.MATERIAL, []), key=lambda b: int(b[0] or 0))
        for alt_bom, alt_text in boms:
            unassigned, total = counts.get((r.MATERIAL, r.ROUTING_GROUP, counter, alt_bom), (0, 0))
            entries.append(Routing(
                material=r.MATERIAL, group=r.ROUTING_GROUP, counter=counter,
                description=(r.ROUTING_DESCRIPTION or "").strip(), plant=plant,
                group_size=int(group_sizes.get((r.MATERIAL, r.ROUTING_GROUP), 1)),
                deleted=r.ROUTING_DELETED == "X",
                alt_bom=alt_bom, alt_bom_text=alt_text, bom_count=len(boms),
                op_0030_wc=wc_by_routing.get((r.MATERIAL, r.ROUTING_GROUP, counter), ""),
                unassigned=unassigned, total_items=total))
    return entries, items
