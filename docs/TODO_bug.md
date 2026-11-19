# TODO - known bugs, open decisions, plans

Created 2026-11-16 after the code review. Nothing in **A** is fixed yet; each item says
what happens, where, and the suggested fix. Line numbers are as of 2026-11-16.

Status: `OPEN` = to do · `DECIDE` = waiting for your decision · `DONE`.

---

## A. Bugs

### A1 - Phase 1 keeps creating BOMs when the Excel save fails  `DONE` 2026-11-16
- **Done:** all 3 phases now show a prompt when the file is open: `[s]` save your edits + close it in Excel and continue, `[r]` I closed it, retry, `[q]` stop. On `[q]` the SAP loop **stops at once** and prints which row is in SAP but not in Excel (set it by hand before the next run). Code: `workbook.close_in_excel`, `ask_when_locked`, `save_sheets`; used by `save_sheet` / `update_source_columns`. Tested with real Excel on a temp copy: `[s]` kept the unsaved edit, closed the file and saved the results; `[q]` stopped.
- **Was:** `bom_create_main.py` ignored the result of `save(...)` in the loop.
- **What happens:** if `CREATE_BOM.xlsx` is open in Excel, the save fails with "Could not save ... Close it in Excel." but the loop goes on creating BOMs in SAP. Their SUCCESS is never written, so the **next run creates the same BOMs again (duplicates in SAP)**.
- **Limits of the shipped prompt** (still true, worth knowing):
  - In cell edit mode Excel rejects the call: it retries a few times, then stops.
  - Your edits **in result sheets** are overwritten by the script. Sheet1 edits are kept, except the 4 result columns.
  - It can't close a lock held by another user / OneDrive / a read-only copy: falls back to `[r]` / `[q]`.
  - Alternatives that were considered and rejected: writing results to a second
    `..._unsaved_<time>.xlsx` file (two files to merge by hand, and the next run doesn't know
    about them -> duplicates), and `taskkill /IM excel.exe` (closes every workbook, loses all
    unsaved work).

### A2 - No production version for single-routing groups  `DECIDE`
- **Where:** `routing.py:294` adds the prefix `"Single routing, no header overview. "` to the phase 2 message. `production_version.py:208` only accepts a message that **starts with** `"No unassigned components on screen"`.
- **What happens:** a routing whose group has only 1 routing, which phase 2 SKIPPED because nothing was left to assign, gets the message `Single routing, no header overview. No unassigned components on screen, saved. ...`. Phase 3 doesn't recognise it and shows `SKIPPED - Waiting for routing assignment (STATUS SKIPPED)`, so **no production version is created for it**.
- **Fix options:** (a) in phase 3, check `"No unassigned components on screen" in message` instead of `startswith`; (b) drop the note prefix in `routing.py`. (a) is the smaller change.

### A3 - `production_version_fix.py` does nothing  `DECIDE` (archived)
- **Where:** `archive/production_version_fix.py`.
- **What happens:** it reads `create_bom_result.xlsx`, which has no Group / Version columns, so it silently finds nothing. It also imports `INPUT_FILE` / `RESULT_FILE`, which were removed from `production_version.py`, so it won't even start.
- **Status:** moved to `archive/`. Only needed if the old half-finished PV run must still be repaired. Tell me if so, and which file it should read.

### A4 - `update_baseqty.py` never presses Enter  `DECIDE`
- **Where:** `bom.py:329`, the last line of `update_base_quantity()` is `press_enter` without `(session)`.
- **What happens:** the function is referenced but not called, so the base quantity change is not confirmed with Enter. Python raises no error. Only `update_baseqty.py` uses it, not the chain.
- **Fix:** `press_enter(session)`. Test it on one BOM first: pressing Enter may now bring up a pop-up the script doesn't expect.

### A5 - The same value is written two ways, so lookups miss  `OPEN`
The scripts match rows by building a key from values (material, alt BOM, GrC). If one place writes `"01"` and another writes `"1"`, or `""` and `"0"`, the keys differ and the lookup finds nothing, **without any error**. There are two places this happens:

**A5a - Alt BOM `"01"` vs `"1"` (phase 1)**
- **Where:** `reader.py:47` and `bom_create_main.py:35` clean the alt BOM with `clean_text`, which only trims spaces, so `"01"` stays `"01"`. SAP and phases 2/3 use `"1"` (via `clean_component`).
- **What happens if Sheet1 has `01`:**
  1. `decide_action` checks `alt == "1"` to choose CHANGE, so `"01"` becomes **CREATE instead of CHANGE**. A BOM that already exists is created as a new alternative.
  2. After the create, SAP's alt BOM (`2`) is compared with Excel's (`02`), which gives a **false "WARNING: SAP alt BOM 2, Excel alt BOM 02"**.
- **Only happens** if Sheet1 holds the alt BOM as text with a leading zero. Typed numbers (`1`, `2`) are fine.
- **Fix:** clean the alt BOM with `clean_component` in `reader.py` and `bom_create_main.previous_results`. Note that SUCCESS rows already written with `"01"` would then need the same cleaning when they're read back.

**A5b - Empty GrC `""` vs `"0"` (phase 2, Snowflake)**
- **Where:** `routing_sf.py:213` and `:256` build the work center dictionary with the raw `GROUP_COUNTER`, and `:261` groups the allocation items by the raw `GROUP_COUNTER`. The lookups at `:223`, `:270` and `:277` use `counter = GROUP_COUNTER or "0"`.
- **What happens for a routing whose GrC comes back empty from Snowflake** (`LTRIM('0','0')` = `""`, i.e. group counter 0):
  1. The dictionary key has `""`, but the lookup uses `"0"`, so the **op 0030 work center is not found**. A lubricant isn't recognised as YMHI / YMHI-1 and goes to 0020 instead of 0030.
  2. (altbom task) The item counts aren't found, so `total_items = 0` and the routing is **SKIPPED "No BOM items (Snowflake)"** although it has items.
- **Only happens** for GrC 0 / empty. Most routings have GrC 1, 2, ..., so this is rare.
- **Fix:** apply the same `or "0"` (or `routing.clean_counter`) when building the dictionary keys and the item groups.

### A6 - Unclear crashes  `OPEN` (low)
- `continue` before the result sheet exists (e.g. `python routing_main.py continue` before any normal run) ends with a Python traceback (`Worksheet named 'Routing Result' not found`) instead of a clear message. The same applies to `production_version.py continue`.
- `routing_main.py:105` uses `int(alt_bom)`, so a non-numeric `Created Alt BOM` crashes the whole work-list build.
- **Fix:** check `has_sheet` first and print "run without continue first"; mark a non-numeric alt BOM as FAILED for that row.

### A7 - Phase 1 trusts the save blindly  `OPEN` (medium)
- **Where:** `bom_create_main.py:88-90`. After `save_bom`, the status is always SUCCESS, and the status bar text is only copied into MESSAGE.
- **What happens:** if SAP shows an error after Save (status bar type `E` / `A`), the row is still SUCCESS, so it's **never retried and phase 2 looks for a BOM that doesn't exist**.
- **Fix:** read `wnd[0]/sbar` `MessageType` and only set SUCCESS for `S` (or empty with a "created" text); otherwise FAILED.

### A8 - Phase 3 hides the real skip reason  `OPEN` (low)
- **Where:** `production_version.py:211`.
- **What happens:** every phase 2 status other than SUCCESS or "No unassigned..." becomes `SKIPPED - Waiting for routing assignment (STATUS SKIPPED)`, also for final reasons like `Already assigned (Snowflake)`, `Trial alt BOM` or `No routing in MLX_DATALAKE`. That's misleading: nothing is waiting.
- **Fix:** only say "Waiting ..." for PENDING / FAILED; for SKIPPED, copy the phase 2 message (`Routing skipped: <reason>`). Decide together with A2: "Already assigned" probably **should** get a production version.

---

## B. One script for all phases (`run_all.py`)  `DONE`

Built and in use. Usage and behaviour: the docstring of `run_all.py` and
[`workflow.md`](workflow.md) section 1.

It turned out differently from the plan this section used to hold: the phases were **not**
changed to return their counts from `main()`. `run_all.py` runs each phase as a subprocess
and reads the STATUS counts back out of the result sheets (`run_all.py`, `status_counts`),
so no phase had to change at all.

- **Still open before trusting an unattended full run:** A7 (`OPEN`) and A2 (`DECIDE`).
  A1 is already `DONE`, so a locked Excel file no longer loses results silently. A7 is the
  important one left - a phase 1 row that SAP rejected is still marked SUCCESS, so the chain
  carries on with a BOM that does not exist.

---

## C. Doc mismatches (log only, not fixed yet)

Line numbers of `workflow.md` below are as of **2026-11-17**, after the doc pass that fixed
the renamed script names (`create_main.py` -> `bom_create_main.py`, `main.py` ->
`bom_update_main.py`) and added `run_all.py` to the commands section.

| Doc | Line | Says | Actually |
|---|---|---|---|
| `routing.md` | 27 | `routing_main.py` "writes the `Result` sheet" | task newbom writes `Routing Result`; only altbom writes `Result` |
| `routing.md` | 112-113 | "The first sheet (the source) is never changed" | true for phase 2; but in the chain file, phase 1 writes result columns into Sheet1 |
| `routing.md` | 200 (pseudocode) | single routing only logs "Single routing, no header overview" | that text is also put in front of MESSAGE (see A2) |
| `workflow.md` | 129 | "run `unmerge_excel.py` first" | `unmerge_excel.py` points to `CHANGE_BOM_demo.xlsx`; change `UNMERGE_FILE` in `config.py` to `CREATE_BOM.xlsx` (or unmerge by hand) |
| `workflow.md` | commands table | - | ~~`python production_version.py test [slow]` not listed~~ **FIXED 2026-11-17** (`test` is now in the options table). `python routing.py` (one routing) is still not listed |
| `workflow.md` | 191 | "Waiting for routing assignment" for PENDING / FAILED | also shown for final SKIPPED reasons (A8) |
| ~~`archive/docs/routing_altbom_plan.md`~~ | whole file | old design | **file is not in the repo** (checked 2026-11-17); see `routing.md` |

---

## D. Other notes from the review
- `routing_sf.load_routings()` is unused today. It's kept on purpose for a future standalone production version task (routing list from Snowflake instead of `Routing Result`).
- `bom_update_main.py` (CS02) saves its input with one sheet only. Never point it at the chain file.
- `docs/TODO.md` (TODO-1: search the first screen before scrolling in `bom.py`, base quantity for new BOMs) is still open and separate from this list.

---

## E. Plan - CA01 (create routing) later  `PLANNED`

CA01 and CA02 share most screens; only the way in differs.

| Part | CA01 | CA02 | Share? |
|---|---|---|---|
| Start + navigation | material, plant, key date, then create pop-ups (group / header data) | group pop-up / header overview | no, separate steps |
| Entry option pop-up | similar | similar | probably, if the IDs match |
| Operation overview, BOM pop-up, component allocation, Save | same screens | same screens | **yes** (incl. AltBOM search, 0030/0020 rules) |

**Watch out:** in create mode, display fields can become input fields: `txt[..]` can turn into `ctxt[..]`, or column numbers can shift. "Same screen" must be confirmed by a recording.

**Steps when the task starts:**
1. Record CA01 once (`.vbs`), for a material with 2 alt BOMs.
2. Compare its IDs with `routing.py`.
3. Move only the matching parts to `routing_screens.py` (shared CA01/CA02); check CA02 with the fake test plus one `slow` run.
4. Write the CA01 steps in `routing_create.py`.
5. The rules (`is_packaging`, `is_lubricant`, `planned_operation`, `skip_reason`) are reused as they are.

Same idea as the shared UI catalog (`sap_ui/`) discussed on 2026-11-16: do it per screen when there is a real second use.

---

## F. Plan - standalone production versions from a material list  `PLANNED` (2026-11-18)

Goal: any file with a `Material` column (e.g. `add_routing_test.xlsx`) -> a production version for
every active routing x every active alt BOM. Not possible today: `production_version.py` only reads
`Routing Result` of `CHAIN_FILE` (phase 3 of the chain).

**Decided (2026-11-18):**
- Alt BOM: **every active alt BOM** (`SQL_ALT_BOMS`), not the file's `Alt Bom` column.
- Rows: **all materials** in the file (the `Action` column is ignored).
- Duplicate = existing version with **same group + GrC + alt BOM** -> SKIPPED. A routing can get one
  version per alt BOM.

### F1 - Check for existing production versions  `TODO` (must be done before a real run)
- **Today there is no duplicate check at all.** `read_versions()` (`production_version.py:75`)
  reads only the version **number** (column 0) to find the next free number. Running twice, or on a
  routing that already has a version, **creates a duplicate**. The chain doesn't notice because it
  only works on BOMs it just created.
- **Plan:** in MM02, after opening the version list, read number + group + GrC + alt BOM of every
  existing version (once per material) and skip `(group, GrC, alt BOM)` that already exist:
  `SKIPPED "Production version exists (0003)"`.
- **Needs:** the column IDs of group / GrC / alt BOM in `wnd[1]/usr/tblSAPLMDIAT_MKAL_NEW` (only column
  0 is known). A screenshot of the list or one recording is enough.
- Why SAP and not Snowflake (MKAL): Snowflake lags about 20 min, so a version created in the last
  run would not be seen. Snowflake MKAL may be used as a pre-check in `check` mode only.

### F2 - Routing filter compared with ROUTING (`checks/routing_condition`)
`routing_sf.SQL_ROUTINGS` (this repo) vs `ROUTING/src/routing/data/routing.py`
`SQL_ROUTINGS_WITH_RESOURCE` (used by `checks/routing_condition`), checked 2026-11-18:

| Filter | change_bom `SQL_ROUTINGS` | ROUTING `routing_condition` | Effect |
|---|---|---|---|
| MAPL -> PLKO join, plant 1901 | yes | yes | same |
| `PLKO.LOEKZ` / `MAPL.LOEKZ` blank | yes (`TRIM(..) = ''`) | yes (`= ' '`) | same |
| `DW_SOFT_DELETE_FLAG = 0` (PLKO, MAPL) | yes | yes | same (permanently deleted gone) |
| Deletion flag `PLKO.DELKZ` | returned; `X` -> **SKIPPED "Deleted routing"** | returned, **not** filtered (flagged = still exists, for reactivation) | same data, different use: for production versions we skip X |
| `PLNTY = 'N'` (standard routing) | **yes** | no (only in the join) -> rate routings (Q) can appear | ours stricter |
| Latest change state | **QUALIFY latest `ZAEHL`** | `SELECT DISTINCT` (all change states) | ROUTING can return the same routing twice, e.g. once with DELKZ X and once blank; ours takes the current state |
| Needs operations (PLAS/PLPO/CRHD join) | no | yes (inner join) | ROUTING hides routings without active operations; ours includes them |
| Material status (`MARC.MMSTA` E3, `LVORM`) | **no** | not here (but `group_counters` and `delete_status` check it; the file shows "material status E3 (not active in this plant)") | **gap:** a material with status E3 would still get production versions tried - add a MARC filter |
| Trial routing (`TRIAL` in KTEXT) | skipped in code (`is_trial`) | ROUTING's matcher has its own trial rule | same intent |

To do for F: add the MARC filter (`MMSTA != 'E3'`, `LVORM` blank, `DW_SOFT_DELETE_FLAG = 0`).
Decide whether a routing without active operations should get a production version.

### F3 - Build (after F1)
`python production_version.py materials <file>`:
1. Read `Material`.
2. Snowflake: `load_routings` + `SQL_ALT_BOMS`.
3. Work list routing x active alt BOM (skip DELKZ / trial / E3).
4. MM02 with the F1 check.
5. `Prod Version Result` sheet in that file, with `check` / `continue` / `fresh` like phase 3.

**Lag:** routings the ROUTING tool changed in the last ~20 min (e.g. set inactive) still look active
in Snowflake. The SAP-side check (F1) and the MM02 "Check" button (group / GrC valid) are the safety net.
