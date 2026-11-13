# Plan - CA02 component assignment for all active alternative BOMs

Status: **implemented 2026-11-12** - current usage and logic: see [`routing.md`](routing.md).
This file is kept as the design record (decisions, data checks).

Differences between this plan and the implementation:

| Plan | Implemented |
|---|---|
| `Preview` sheet with planned operation per component | Implemented as third sheet (source, `Result`, `Preview`), PENDING entries only, built once; not rebuilt by `continue` |
| BOM popup: `txt[<alt BOM - 1>,8]` (by position) | Row found **by value** in the AltBOM column `txt[row,19]`; the popup lists all usages, so position was wrong |
| `retry` = FAILED rows | `continue` (alias `retry`) = `PENDING` + `FAILED` rows, no Snowflake |
| Separate result file `..._routing_altbom.xlsx` | `Result` sheet in the input file; first sheet unchanged |
| Existing task keeps alt BOM 2 | Both tasks use the same alt BOM logic (O2) |
| - | `check` (work list only, no SAP) and `fresh` (rebuild `Result`) options |

---

## 1. Task

For every routing of a material, assign the components that are **not assigned
yet**, for **every active alternative BOM** of the material.

Today the script handles one BOM per routing (always the BOM at `txt[1,8]` in the
BOM popup = alt BOM 2) and only materials whose BOM was just created
(`STATUS = SUCCESS`).

---

## 2. Business rules

### 2.0 Which routings

All active standard routings of the material at plant 1901 (as today), with two
skip rules. Skipped routings are written to the result file with the reason,
nothing is done in SAP:

| Rule | Detail | Result |
|---|---|---|
| Deleted routing | Deletion indicator set on the routing header: `PLKO.DELKZ = 'X'` | `SKIPPED - Deleted routing` |
| Trial routing | Routing description (`PLKO.KTEXT`) contains `TRIAL` (case-insensitive; covers `MDTRIAL`) | `SKIPPED - Trial routing` |

Example 5052760489 (checked 2026-11-12):

| Group / GrC | Description | DELKZ | Result |
|---|---|---|---|
| M0527604 / 16 | 5052760489 - MF15-17 - L505276A | | processed |
| M0527604 / 9 | MN15-10-5052760489-L505276C-YMHI-1 | | processed |
| M0527604 / 17 | MDTRIAL - MN15-10 - L505276A - YMHI-1 | | SKIPPED - Trial routing |
| M0527604 / 6 | YMHI-1 - 5052760489 | X | SKIPPED - Deleted routing |

Note: the current `routing_sf.py` (used by the existing task and by
`production_version.py`) does **not** filter `DELKZ` yet. Both skip rules apply
to the existing task too (decision Q12), so they are added there when this plan
is implemented.

### 2.1 Which alternative BOMs

| Rule | Detail |
|---|---|
| Usage / plant | BOM usage `1`, plant `1901` (`MAST.STLAN = '1'`, `MAST.WERKS = '1901'`) |
| Active | Status `01` (`STKO.STLST = '01'`) on the **latest change state** of the alt BOM |
| Exclude TRIAL | Skip an alt BOM whose alt BOM text (`STKO.STKTX`, latest change state) contains `TRIAL` (case-insensitive) |
| All of them | Each remaining alt BOM is processed for **every** routing of the material |

**Latest change state**: STKO can hold several rows per alt BOM (one per change
state `STKOZ`), and status / text can differ between them. Example 5016441210:

| Alt BOM | STKOZ | STLST | STKTX | Valid from |
|---|---|---|---|---|
| 01 | 1 | 01 | (empty) | 2010-06-01 |
| 02 | 4 | 02 | REGRIND | 2022-03-25 |
| 02 | 11 | 01 | BOM 2 - REGRIND | 2026-11-02 |

Only the row with the highest `STKOZ` per alt BOM decides status and text
(here: alt 01 active, alt 02 active with text "BOM 2 - REGRIND").

Skipped alt BOMs (TRIAL) are still written to the result file
(`SKIPPED - Trial alt BOM`) so it is visible why they were not processed.

### 2.2 Which components (unchanged)

- Only rows with an **empty operation** (`txt[10]`) are assigned.
- Rows that already have an operation (0030, 0040, ...) are never touched.
- The scan stops at the first blank row.
- Only the visible rows are checked (no scrolling, for now).

### 2.3 Which operation

Decided per component, first matching rule wins:

| # | Component description (`txt[13]`) contains | Routing condition | Operation |
|---|---|---|---|
| 1 | packaging word: carton, sheet, bag, silica, bubble, box | always (every routing has operation 0030) | **0030** |
| 2 | `LUBRICANT` | work center of operation 0030 is **exactly** `YMHI` or `YMHI-1` | **0030** |
| 3 | anything else (incl. LUBRICANT when rule 2 does not apply) | - | **0020** |

- Description matches are case-insensitive "contains" (plurals match: boxes, sheets).
- Work center match is **exact**: `YMHI-2` ... `YMHI-27` do **not** count.

Implementation: the existing two passes stay, pass 1 gets a wider filter:

```
Pass 1 (0030): empty operation AND (packaging OR (lubricant AND routing_is_ymhi))
Pass 2 (0020): empty operation (everything still left)
```

### 2.4 Rules that may change later

Kept in one place each (function / constant), so a rule change is a one-line edit:

| Rule | Where (planned) |
|---|---|
| Packaging words | `PACKAGING_WORDS` |
| Lubricant words | `LUBRICANT_WORDS = ("lubricant",)` |
| YMHI work centers | `LUBRICANT_0030_WORK_CENTERS = ("YMHI", "YMHI-1")` |
| Active alt BOM | `ACTIVE_BOM_STATUS = "01"` |
| Trial alt BOM | `is_trial_bom(text)` |
| Deleted routing | `is_deleted_routing(delkz)` (`DELKZ = 'X'`) |
| Trial routing | `is_trial(description)`, shared with `production_version.py` |

---

## 3. Choosing the alt BOM in the `btn[7]` popup

After `wnd[0]/tbar[1]/btn[7]` the BOM popup lists the alt BOMs. The alt BOM is
chosen by double-clicking:

```
wnd[1]/usr/sub/1/txt[<alt BOM - 1>, 8]
```

| Alt BOM | Element |
|---|---|
| 1 | `txt[0,8]` |
| 2 | `txt[1,8]` (the old fixed recording) |
| 3 | `txt[2,8]` |

Assumption: the popup lists alt BOMs in order 1, 2, 3, ... without gaps. If a
column in that popup shows the alt BOM number, the script reads it first and
fails if it does not match (instead of assigning to the wrong BOM).

---

## 4. Data needed per iteration

One iteration = **one routing x one alt BOM**.

| Field | Example | Source | Used for |
|---|---|---|---|
| material | 5016441210 | input file | CA02 material field |
| plant | 1901 | fixed | CA02 plant field |
| group | M0164412 | Snowflake MAPL/PLKO (as today) | group popup |
| GrC | 22 | Snowflake MAPL/PLKO (as today) | header overview |
| group_size | 3 | counted from routing list | "no header overview" safety check |
| alt BOM | 01, 02 | Snowflake MAST/STKO (new) | `txt[alt-1,8]` in the BOM popup |
| alt BOM text | BOM 2 - REGRIND | Snowflake STKO.STKTX (new) | TRIAL filter, result file |
| op 0030 work center | OMINS / YMHI-1 | Snowflake PLPO/CRHD (new) | lubricant rule |
| routing deleted | X / blank | Snowflake PLKO.DELKZ (new) | skip rule |
| routing description | MF50-07 - ... | Snowflake | trial skip rule, log |
| unassigned items | 0 / 5 | Snowflake allocation query (new, 4.2) | skip entries with nothing to assign, preview |

### 4.1 Queries (based on `ROUTING/src/routing/data`)

**Active alt BOMs** - like `SQL_ALT_BOM_COUNT` (`data/bom.py`), but one row per
alt BOM with the latest change state:

```sql
SELECT LTRIM(m.MATNR, '0') AS MATERIAL, m.STLAL AS ALT_BOM, s.STLST, s.STKTX AS ALT_BOM_TEXT
FROM MLX_DATALAKE.MLX_SAP_ECC.MAST m
JOIN MLX_DATALAKE.MLX_SAP_ECC.STKO s ON m.STLNR = s.STLNR AND m.STLAL = s.STLAL
WHERE m.STLAN = '1' AND m.WERKS = '1901' AND LTRIM(m.MATNR, '0') IN (...)
QUALIFY ROW_NUMBER() OVER (PARTITION BY m.MATNR, m.STLAL ORDER BY s.STKOZ DESC) = 1
-- then keep STLST = '01'
```

**Operation 0030 work center** - like `SQL_WORK_CENTERS_0020` (`data/routing.py`)
with `po.VORNR = '0030'`: one work center per material + group + GrC.

Both checked on 2026-11-12 with material 5016441210 (read-only): data is there;
all its routings have work center `OMINS` at 0030.

### 4.2 Component allocation per routing x alt BOM (pre-run)

One row per **routing x alt BOM x BOM item**, with the operation the item is
allocated to. **Empty `OPERATION` = not assigned yet.**

| Step | Tables | What it does |
|---|---|---|
| routings | MAPL + PLKO | active standard routings at 1901, latest change state; `DELKZ` and description returned for the skip rules |
| alt_boms | MAST + STKO | usage 1, plant 1901, latest change state, status `01` |
| items | STAS + STPO | items of each alt BOM, latest change state, not deleted |
| alloc | PLMZ + PLPO | operation (`VORNR`) per item and routing, latest change state, not deleted |
| final | + MAKT | component description (English); `LEFT JOIN` alloc, so unassigned items stay with empty operation |

```sql
WITH routings AS (
    SELECT mp.MATNR, pk.PLNTY, pk.PLNNR, pk.PLNAL, pk.KTEXT, pk.DELKZ
    FROM MLX_DATALAKE.MLX_SAP_ECC.MAPL mp
    JOIN MLX_DATALAKE.MLX_SAP_ECC.PLKO pk
        ON mp.MANDT = pk.MANDT AND mp.PLNTY = pk.PLNTY AND mp.PLNNR = pk.PLNNR AND mp.PLNAL = pk.PLNAL
    WHERE pk.WERKS = '1901' AND mp.PLNTY = 'N'
        AND TRIM(pk.LOEKZ) = '' AND TRIM(mp.LOEKZ) = ''
        AND pk.DW_SOFT_DELETE_FLAG = 0 AND mp.DW_SOFT_DELETE_FLAG = 0
        AND LTRIM(mp.MATNR, '0') IN (...)
    QUALIFY ROW_NUMBER() OVER (PARTITION BY mp.MATNR, pk.PLNNR, pk.PLNAL ORDER BY pk.ZAEHL DESC) = 1
),
alt_boms AS (
    SELECT m.MATNR, m.STLNR, m.STLAL, s.STKTX
    FROM MLX_DATALAKE.MLX_SAP_ECC.MAST m
    JOIN MLX_DATALAKE.MLX_SAP_ECC.STKO s
        ON m.MANDT = s.MANDT AND s.STLTY = 'M' AND m.STLNR = s.STLNR AND m.STLAL = s.STLAL
    WHERE m.STLAN = '1' AND m.WERKS = '1901' AND m.DW_SOFT_DELETE_FLAG = 0 AND s.DW_SOFT_DELETE_FLAG = 0
        AND LTRIM(m.MATNR, '0') IN (...)
    QUALIFY ROW_NUMBER() OVER (PARTITION BY m.MATNR, m.STLAL ORDER BY s.STKOZ DESC) = 1
        AND s.STLST = '01'
),
items AS (
    SELECT a.MATNR, a.STLNR, a.STLAL, a.STKTX, st.STLKN, p.POSNR, p.IDNRK
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
    QUALIFY ROW_NUMBER() OVER (PARTITION BY z.PLNNR, z.ZUONR ORDER BY z.ZAEHL DESC, po.ZAEHL DESC) = 1
        AND TRIM(z.LOEKZ) = ''
)
SELECT LTRIM(r.MATNR, '0') AS MATERIAL, TRIM(r.PLNNR) AS ROUTING_GROUP, LTRIM(r.PLNAL, '0') AS GROUP_COUNTER,
       r.KTEXT AS ROUTING_DESCRIPTION, COALESCE(TRIM(r.DELKZ), '') AS ROUTING_DELETED,
       LTRIM(i.STLAL, '0') AS ALT_BOM, i.STKTX AS ALT_BOM_TEXT,
       i.POSNR AS ITEM, LTRIM(i.IDNRK, '0') AS COMPONENT, mk.MAKTX AS COMPONENT_DESCRIPTION,
       a.VORNR AS OPERATION                      -- empty = not assigned yet
FROM routings r
JOIN items i ON i.MATNR = r.MATNR
LEFT JOIN alloc a
    ON a.PLNNR = r.PLNNR AND a.PLNAL = r.PLNAL AND a.STLNR = i.STLNR AND a.STLAL = i.STLAL AND a.STLKN = i.STLKN
LEFT JOIN MLX_DATALAKE.MLX_SAP_ECC.MAKT mk
    ON mk.MATNR = i.IDNRK AND mk.SPRAS = 'E' AND mk.DW_SOFT_DELETE_FLAG = 0
ORDER BY ROUTING_GROUP, GROUP_COUNTER, ALT_BOM, ITEM
```

Deleted and trial routings are **returned** (with `ROUTING_DELETED` /
`ROUTING_DESCRIPTION`) and skipped in Python, so they appear in the result file
as SKIPPED instead of disappearing.

**Checked 2026-11-12 (read-only):**

| Material | Result |
|---|---|
| 5016441210 | 4 routings x 2 alt BOMs, all items assigned (alt 2 items 899922452, 967502452, 682002452 on 0020 = the BOM update request) |
| 5052760489 | 4 routings x 2 alt BOMs; GrC 16 and 9 fully assigned; GrC 17 (MDTRIAL) alt 2: 5 unassigned; GrC 6 (`DELKZ = X`): 6 unassigned |

Verified against the raw tables for 5052760489:

- Deleted allocations are ignored: e.g. GrC 16 alt 1 item 0020 (LUBRICANT) has an
  old allocation to 0020 with `PLMZ.LOEKZ = X` and an active one to 0030 -> query returns 0030.
- Every PLMZ row appears twice in the raw join (two PLPO change states); the
  `QUALIFY` keeps one.
- Deleted BOM items (missing STAS node) are not returned.
- The lubricant rule matches existing SAP data: lubricants are on 0030 for the
  YMHI-1 routings (16, 9, 17) and on 0020 for GrC 6 (op 0030 = OMINS).

Caveat: the data lake can lag behind SAP. The query is used for planning and
skipping; during the run the script still checks `txt[10]` on screen before
ticking a row.

---

## 5. Input file

### 5.1 This task (new)

| Column | Required | Notes |
|---|---|---|
| `Material` | yes | one row per material, duplicates ignored, leading zeros not needed |

No `STATUS` column: every listed material is processed.

### 5.2 Existing task (unchanged)

`create_bom_final.xlsx`: `Parent Material` + `STATUS`; only `SUCCESS` rows,
only the new BOM (`txt[1,8]`).

### 5.3 Choosing the task

Explicit option on the command line (not guessed from the columns), so a file
for one task can't accidentally be run with the other task's rules:

```
python routing_main.py                     # existing task (new BOM only)
python routing_main.py altbom              # this task (all active alt BOMs)
```

---

## 6. Flow

```
READ input file -> materialList
LOAD from Snowflake:
    routings   (material, group, GrC, description, DELKZ, group_size)       # as today + DELKZ
    alt BOMs   (material, alt BOM, text)  latest change state, status 01   # new
    op 0030 WC (material, group, GrC, work center)                         # new
    allocation (routing x alt BOM x item -> operation, empty = unassigned)  # new, 4.2

BUILD work list = routing x alt BOM
    IF routing DELKZ = X                  THEN item = SKIPPED "Deleted routing"
    ELSE IF routing description has TRIAL THEN item = SKIPPED "Trial routing"
    ELSE IF alt BOM text has TRIAL        THEN item = SKIPPED "Trial alt BOM"
    ELSE IF no unassigned item (Snowflake) THEN item = SKIPPED "Already assigned (Snowflake)"
    routing_is_ymhi = op 0030 WC IN ("YMHI", "YMHI-1")   (exact)

WRITE preview file (work list + planned operation per unassigned component)

FOR EACH item IN work list (not SKIPPED)
    TRY
        GO TO CA02, ENTER material + 1901, Enter
        IF group popup THEN select group END IF
        IF header overview THEN open GrC
        ELSE IF group_size > 1 THEN FAILED ELSE log "single routing" END IF

        PRESS btn[7]
        DOUBLE-CLICK wnd[1]/usr/sub/1/txt[<alt BOM - 1>,8]

        Pass 1 -> 0030: empty op AND (packaging OR (lubricant AND routing_is_ymhi))
        Pass 2 -> 0020: empty op
        SAVE (also when nothing was assigned)
        SUCCESS / SKIPPED (nothing to assign on screen)
    CATCH error
        FAILED + message
    END TRY
    WRITE result row
END FOR
```

One CA02 entry and one save per routing x alt BOM, so a failure only affects that item.

Order of the skip checks matters only for the message; every skipped item is
written with the first reason that applies. Every skipped item is also printed
in the console and written to the result file (STATUS `SKIPPED`, MESSAGE = reason),
e.g. `SKIPPED - Already assigned (Snowflake): 3 of 3 components assigned`.

### 6.1 Pre-run preview

Before SAP is touched, the script writes the work list with the planned result
per component (from the allocation query + rules), e.g.:

| Material | Group / GrC | Alt BOM | Component | Description | Planned |
|---|---|---|---|---|---|
| 5052760489 | M0527604 / 16 | 2 | 2027490008 | LUBRICANT HANARL MLX-79A | already 0030 |
| ... | ... | ... | ... | ... | 0030 (packaging) / 0030 (lubricant, YMHI-1) / 0020 |

So the user can check the plan first; optional `preview` option that stops after
writing it (no SAP).

---

## 7. Output file

One row per **material x group x GrC x alt BOM** (`..._routing_altbom.xlsx`):

| Column | Example |
|---|---|
| Material | 5016441210 |
| Group | M0164412 |
| GrC | 22 |
| Description | MF50-07 - 5016441210 |
| Alt BOM | 02 |
| Alt BOM Text | BOM 2 - REGRIND |
| Op 0030 WC | OMINS |
| STATUS | SUCCESS / SKIPPED / FAILED |
| MESSAGE | `0030: 2 rows (1 lubricant), 0020: 5 rows. <SAP status>` |

Retry works as today (`python routing_main.py altbom retry`): only
`STATUS = FAILED` rows, file updated in place; `Alt BOM` is read from the file.

---

## 8. Making the script multi-purpose

One runner, task chosen on the command line:

| Setting | Existing task | `altbom` task |
|---|---|---|
| Material source | `Parent Material` where `STATUS = SUCCESS` | `Material`, all rows |
| BOM selection | alt BOM 2 (`txt[1,8]`) | every active, non-TRIAL alt BOM |
| Lubricant rule | off | on |
| Result file | `..._routing.xlsx` | `..._routing_altbom.xlsx` |

```
python routing_main.py [retry] [slow]          # existing task
python routing_main.py altbom [retry] [slow]   # this task
```

---

## 9. Decisions

| # | Question | Decision |
|---|---|---|
| Q1 | What is an active alt BOM? | Usage 1, plant 1901, `STLST = '01'` on the latest change state (`STKOZ`) |
| Q2 | How is the alt BOM chosen in the popup? | `wnd[1]/usr/sub/1/txt[<alt BOM - 1>,8]` |
| Q3 | YMHI match | Exact: `YMHI` or `YMHI-1` |
| Q4 | Source of op 0030 work center | Snowflake PLPO/CRHD (`VORNR = '0030'`) |
| Q5 | Packaging without op 0030? | Not needed: every routing has op 0030 |
| Q6 | How to choose the task | Explicit `altbom` option |
| Q7 | One CA02 entry per alt BOM? | Yes: one entry, assign, save, result row per routing x alt BOM |
| Q8 | Scrolling in component table | Not now (warning stays) |
| Q9 | Deleted routings (`PLKO.DELKZ = 'X'`) | Skip, `SKIPPED - Deleted routing` (also to be added to `routing_sf.py`) |
| Q10 | Trial routings (description contains `TRIAL` / `MDTRIAL`) | Skip, `SKIPPED - Trial routing` |
| Q11 | Use Snowflake allocation before the run | Yes: skip routing x alt BOM with nothing unassigned, write a preview; SAP screen still decides during the run |
| Q12 | Skip rules (deleted / trial routing) for the existing task too? | Yes: apply to both tasks |
| Q13 | Entries already fully assigned (Snowflake) | Skip, no CA02 entry, logged as `SKIPPED - Already assigned (Snowflake)` |

## 10. Still open

| # | Question |
|---|---|
| O1 | Does the BOM popup show the alt BOM number in a column (to check before double-click)? Element ID of that cell? |
| O2 | Existing task always picks alt BOM 2 (`txt[1,8]` in `open_component_allocation()`). Keep that, or take the alt BOM from the input file (`Alternative BOM` / `CREATED ALT BOM`) so a new BOM that is alt 3 is not missed? |
