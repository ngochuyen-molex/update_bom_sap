# Workflow - create BOM -> assign components -> production version

Three phases, **one workbook**. Phase 1 writes its result into the source sheet (new result
columns on each component row, all other columns unchanged); phases 2 and 3 each write
their own result sheet.

> **IMPORTANT: close the workbook in Excel before every run.** The scripts write their
> results into it during the run; an open file is locked by Windows and the save fails.

```
CREATE_BOM.xlsx  (workbook.CHAIN_FILE)
├─ Sheet1               source (you) + phase 1 result columns   python create_main.py
├─ Routing Result       phase 2                                 python routing_main.py
└─ Prod Version Result  phase 3                                 python production_version.py
```

| Phase | Script | Reads | SAP | Snowflake | Writes | One row per |
|---|---|---|---|---|---|---|
| 1 Create BOM | `create_main.py` | `Sheet1` | CS01 | no | result columns in `Sheet1` | component (source row) |
| 2 Assign components | `routing_main.py` | `Sheet1` (STATUS SUCCESS) | CA02 | routings + op 0030 WC | `Routing Result` | routing x new alt BOM |
| 3 Production version | `production_version.py` | `Routing Result` | MM02 | no | `Prod Version Result` | routing x new alt BOM |

`Routing Result` and `Prod Version Result` have the same rows in the same order (one per
routing x new alt BOM). `Sheet1` has one row per component, so it is usually longer than
the number of BOMs.

---

## 0. Snowflake lag - why the chain doesn't take the new BOM from Snowflake

> **MLX_DATALAKE lags behind SAP by about 20 minutes (or more).** Anything created or
> changed in SAP - a new BOM, a component allocation, a production version - is not in
> Snowflake until then.

Consequence for the chain: right after phase 1, the **new alt BOM is not in Snowflake yet**.
If phase 2 took its alt BOMs and allocation from Snowflake (like the task "altbom" does),
the new alt BOM would simply be missing from the work list and nothing would be assigned
to it - without any error.

So each piece of data is taken from where it is reliable:

| Data | Changed by this chain? | Source in the chain |
|---|---|---|
| New alt BOM (number) | yes, created in phase 1 | `Sheet1`, column `Created Alt BOM` (rows with STATUS SUCCESS) |
| Component allocation of the new alt BOM | yes, phase 2 | SAP screen (every routing is opened; only empty operations are ticked) |
| Routings (group, GrC, description, deleted flag) | no, existed before | Snowflake |
| Op 0030 work center | no | Snowflake |
| Routing list for production versions | - | `Routing Result` sheet (no Snowflake) |

### Two ways to build the routing work list (same code, different source)

The routing assignment is used for two different jobs, so `routing_main.py` has two tasks.
Both use the same SAP steps, rules and result format; only the work list differs:

| | Task "newbom" (this chain) | Task "altbom" (existing BOMs) |
|---|---|---|
| Use when | BOMs were **just created** (phase 1) | checking / completing assignment of **existing** BOMs |
| Materials | `Sheet1` rows with STATUS SUCCESS | input file, column `Material` |
| Alt BOMs | only the **new** one, `Created Alt BOM` in `Sheet1` | **all** active alt BOMs, from Snowflake |
| Allocation check before SAP | none (not in Snowflake yet) | Snowflake: fully assigned entries are skipped |
| Routings + op 0030 WC | Snowflake | Snowflake |
| Command | `python routing_main.py` | `python routing_main.py altbom` |

Rules of thumb:
- Don't run the task "altbom" on materials whose BOM was created less than ~20 min ago:
  the new alt BOM would be missing (use the chain, or wait).
- Any check / validation based on Snowflake (e.g. "is everything assigned now?") should
  run **at least ~20 min after** the SAP run.
- The routings must exist before phase 1; a routing created in SAP minutes before the
  run would also be missing from Snowflake.

---

## 1. Commands

Same options in every phase:

| Option | Phase 1 | Phase 2 | Phase 3 |
|---|---|---|---|
| *(none)* | create every BOM not SUCCESS yet | build work list, run CA02 | build work list, run MM02 |
| `check` | validate, fill result columns in `Sheet1`, no SAP | write `Routing Result`, no SAP | write `Prod Version Result`, no SAP |
| `continue` (`retry`) | - (a normal run always continues) | process `PENDING` / `FAILED` of `Routing Result`, **no Snowflake** | process `PENDING` / `FAILED` of `Prod Version Result` |
| `fresh` | - | rebuild `Routing Result` (keeps SUCCESS rows) | rebuild from `Routing Result` (keeps SUCCESS rows) |
| `slow` | wait after every SAP action | same | same |

```powershell
cd C:\Users\nguyeb3\Downloads\SAP_automation\change_bom

python create_main.py check            # 1. validate the source
python create_main.py                  #    create BOMs (rerun = continue)

python routing_main.py check           # 2. routing list -> Routing Result
python routing_main.py                 #    assign (or: continue)
python routing_main.py continue        #    after a break / for FAILED rows

python production_version.py check     # 3. routing list from Routing Result
python production_version.py           #    create production versions
python production_version.py continue  #    after a break / for FAILED rows
```

A phase only works on the rows the previous phase finished. If phase 2 is continued later,
run phase 3 with `fresh` to pick up the routings that are now done (SUCCESS rows of
phase 3 are kept, nothing is created twice).

---

## 2. Phase 1 - create BOM (`create_main.py`)

**Input `Sheet1`** (merged cells: run `unmerge_excel.py` first), read by `reader.py`:

| Column | Required |
|---|---|
| `Parent Material` | yes |
| `Alternative BOM` | yes (alt BOM 1 = existing -> CHANGE, others -> CREATE, unless `Action` is set) |
| `Component after`, `BOM rate after`, `Unit after`, `Component Scrap after` | yes |
| `Base Qty after` (or `Base Qty`) | yes |
| `Action` | optional: CREATE / CHANGE |

Rows are grouped into one BOM per `Parent Material` + `Alternative BOM` and validated
(base quantity, quantity, unit, duplicate components).

**Output: result columns in `Sheet1`**, filled on **every component row** of the BOM
(same values on all rows of one BOM). Missing columns are added at the end; existing
`STATUS` / `MESSAGE` columns are reused. All other columns and the formatting stay as
they are.

| Column | Meaning |
|---|---|
| Created Alt BOM | alt BOM number SAP assigned in CS01 (used by phase 2) |
| STATUS | `PENDING` / `SUCCESS` / `SKIPPED` (Action CHANGE, not done here) / `FAILED` (validation or SAP error) |
| MESSAGE | SAP status bar text or error |
| Updated At | last change |

A BOM whose rows are SUCCESS is **never created again** (a rerun would make a duplicate
BOM). To mark a BOM that was created outside this script (e.g. by an older script), set
`STATUS = SUCCESS` and `Created Alt BOM` on its rows by hand.

---

## 3. Phase 2 - assign components (`routing_main.py`, task "newbom")

- Materials and **new alt BOM** (`Created Alt BOM`) from `Sheet1` rows with `STATUS = SUCCESS`.
- **All routings of all groups and group counters** of the material from Snowflake,
  **including deleted ones**: deleted routings (`DELKZ = X`) and trial routings
  (description contains `TRIAL`) are written as `SKIPPED` and not opened in SAP.
- Only the **newly created alt BOM** is assigned. Its allocation is not checked in
  Snowflake: a BOM created a few minutes ago is not in MLX_DATALAKE yet (~20 min lag),
  so every routing is opened and the screen decides. The routings themselves already
  existed, so the lag does not matter for them.
- Operation per component: packaging -> 0030, lubricant on op 0030 `YMHI` / `YMHI-1` -> 0030,
  rest -> 0020 (details: [`routing.md`](routing.md) 4.4).

**Output `Routing Result`**: Material, Group, GrC, Description, Alt BOM, Alt BOM Text,
Op 0030 WC, Unassigned (Snowflake) (empty in this task), STATUS, MESSAGE, Updated At.

All other alt BOMs of a material: use the separate task `python routing_main.py altbom`
([`routing.md`](routing.md)).

---

## 4. Phase 3 - production version (`production_version.py`)

- Routing list from **`Routing Result`** - no Snowflake in this phase.

| Routing Result row | Prod Version Result |
|---|---|
| no routing | `SKIPPED` (same message) |
| `Deleted routing` | `SKIPPED - Deleted routing` |
| description contains `TRIAL` | `SKIPPED - Trial routing, no production version` |
| `SUCCESS` or nothing to assign on screen | `PENDING` -> created |
| `PENDING` / `FAILED` | `SKIPPED - Waiting for routing assignment (STATUS ...)`; picked up by `fresh` later |

- New version = highest version < 5000 + 1; text = routing description + ` B` + alt BOM;
  longer than 40 characters -> ` - <material>` removed once, still too long -> `FAILED`.
- A production version that is SUCCESS is **never created again** (also not by `fresh`).
- `FAILED` with `Popup after Save` may already be saved in SAP: check in MM02 before `continue`.

**Output `Prod Version Result`**: Material, Group, GrC, Description, Alt BOM, Version Text,
New Version, STATUS, MESSAGE, Updated At.

### Future: standalone production version task

If a request ever needs **only** production versions (no create BOM / assignment before
it), there is no `Routing Result` sheet to read. That task should get the routing list
from **Snowflake** instead (`routing_sf.load_routings()`: routings with group, GrC,
description, deleted flag), apply the same skip rules (deleted, trial) and take the alt
BOM from its own input. Not built yet - only when such a request comes up.

---

## 5. Status values (all phases)

| STATUS | Meaning | Processed by |
|---|---|---|
| `PENDING` | to do | normal run, `continue` |
| `SUCCESS` | done and saved | never again |
| `SKIPPED` | not done on purpose; reason in MESSAGE | - (`fresh` re-evaluates "Waiting ...") |
| `FAILED` | error; text in MESSAGE | `continue` |

Every result sheet is saved during the run (phase 1 and 3 after every row, phase 2 every
5 rows) and when the run stops (Ctrl+C, crash). Saving writes a temp copy and replaces
the file, so a crash can't leave the workbook half-written.

---

## 6. Files

| File | Role |
|---|---|
| `workbook.py` | `CHAIN_FILE`, sheet names, shared column names, sheet order, safe write of result sheets |
| `create_main.py`, `reader.py`, `bom_create.py` | phase 1 |
| `routing_main.py`, `routing.py`, `routing_sf.py` | phase 2 (+ task altbom) |
| `production_version.py` | phase 3 |
| `models.py` | `BomHeader`, `Routing`, `ProductionVersion` |
| `connect.py` | SAP session, `slow` mode |
| `bom.py` | shared cleaning (`clean_component` = the one material / alt BOM cleaner), `press_enter`, `save_bom`, CS02 helpers |

Full inventory (incl. standalone tools and importance): [`archive/README.md`](../archive/README.md).
Known bugs, open decisions and plans: [`TODO_bug.md`](TODO_bug.md). CS02 search ideas: [`TODO.md`](TODO.md), [`note.md`](note.md).
