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
- **Fix:** stop the loop as soon as a save fails, or use the retry prompt from section B (recommended).
- **Option asked about (2026-11-16): close Excel for you, saving your work first.** When the save fails:
  1. Attach to the running Excel (`win32com`) and find `CREATE_BOM.xlsx`.
  2. `Save()` it, so your edits are kept and no "save changes?" pop-up appears.
  3. `Close()` it.
  4. Save the results again. If that still fails, **stop the run**.
  - Proposed as a prompt: `[s] save your edits + close it and continue  [r] I closed it, retry  [q] stop`. Put it in `workbook.py` once, for all phases.
  - Limits:
    - In cell edit mode Excel rejects the call: retry a few times, then stop.
    - Your edits **in result sheets** are overwritten by the script. Sheet1 edits are kept, except the 4 result columns.
    - It can't close a lock held by another user / OneDrive / a read-only copy: falls back to [r] / [q].
    - It must never wait on a dialog such as Save As.
  - Status: `DONE` (prompt version)

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

## B. "Can it force-close the file and then save?"

Short answer: **yes, technically, but closing your Excel for you is risky.** When the save fails because the workbook is open, the options are:

| # | Option | How | Risk |
|---|---|---|---|
| 1 | **Stop the run** | save fails -> stop SAP work, print what wasn't saved | none; the work stops (fixes A1) |
| 2 | **Wait and retry** (recommended) | save fails -> `"CREATE_BOM.xlsx is open. Close it and press Enter to retry (or type q)"`, then try again. SAP work waits | none; needs you at the PC |
| 3 | **Backup file** | save fails -> write the results to `CREATE_BOM_unsaved_<time>.xlsx`, continue | results in two files, must be merged by hand; the next run doesn't know about them -> duplicates possible |
| 4 | **Force-close via Excel** | `win32com` attaches to the running Excel, finds the workbook and `Close(SaveChanges=False)`, then saves | **your unsaved edits in that workbook are lost**; Excel may show a dialog and hang the script; doesn't help if another user / OneDrive sync holds the lock |
| 5 | Kill Excel | `taskkill /IM excel.exe` | closes **all** workbooks, loses all unsaved work. Not recommended |

**Recommendation:** 2 (+ 1 when you type q). Optionally 4 only after asking: `"Close CREATE_BOM.xlsx in Excel without saving your changes there? (y/n)"`. Put it in `workbook.py` once, so all three phases behave the same.

---

## C. Plan - one script for all phases (`run_all.py`)  `PLANNED`

Yes, this is fine; the phases are already separate modules. Planned design (not built yet):

```text
python run_all.py [check] [slow] [from=2]
  1. phase 1  create_main.main()          -> stop if any BOM FAILED (ask: continue anyway?)
  2. wait     MLX_DATALAKE lag (~20 min) only if the routings of new materials are needed;
              phase 2 takes the new alt BOM from Sheet1, so usually no wait is needed
  3. phase 2  routing_main (task newbom)  -> fresh if the sheet exists, else normal
  4. phase 3  production_version          -> fresh / normal
  5. summary  counts per phase, list of FAILED rows
```
- `check` runs every phase in check mode (no SAP); `from=N` restarts at phase N.
- Each phase's `main()` should **return** its counts instead of only printing them, so `run_all.py` can decide whether to go on. That's a small change per phase.
- Fix A1 / A2 / A7 first; otherwise an unattended full run can silently lose results.

---

## D. Doc mismatches (log only, not fixed yet)

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

## E. Other notes from the review
- `routing_sf.load_routings()` is unused today. It's kept on purpose for a future standalone production version task (routing list from Snowflake instead of `Routing Result`).
- `bom_update_main.py` (CS02) saves its input with one sheet only. Never point it at the chain file.
- `docs/TODO.md` (TODO-1: search the first screen before scrolling in `bom.py`, base quantity for new BOMs) is still open and separate from this list.

---

## F. Plan - CA01 (create routing) later  `PLANNED`

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
