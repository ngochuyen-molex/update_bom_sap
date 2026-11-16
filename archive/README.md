# Archive + file inventory

Cleanup on 2026-11-16. Nothing was deleted: items no longer used were moved here.
Nothing in this folder is imported by the active code.

## 1. What is in this folder

| Item | Was | Why archived |
|---|---|---|
| `production_version_fix.py` | one-off fix (rename / create production versions of the first half-finished PV run) | the job is done; it is also broken: it reads `create_bom_result.xlsx`, which has no Group/Version columns (bug A3), and it imported `INPUT_FILE` / `RESULT_FILE` from `production_version.py`, which were removed. **Will not run from here.** |
| `docs/routing_altbom_plan.md` | design notes of the altbom task | historical; many details are outdated. The current behaviour is in `docs/routing.md` + `docs/workflow.md` |
| `removed_code.py` → `save_excel`, `prepare_result_columns`, `log_result`, `add_log`, `save_progress` | `excel_utils.py` | old single-sheet result helpers, replaced by `workbook.py`. `save_progress` writes the file with **one sheet only** -> would wipe `Routing Result` / `Prod Version Result` if used on the chain file |
| `removed_code.py` → `print_boms` | `reader.py` | debug print, never called (and fails on an empty quantity) |
| (deleted lines) `INPUT_FILE` / `RESULT_FILE` | `production_version.py` | only used by the fix script |

To bring something back: copy it back to its file. Don't run anything from this folder directly.

## 2. Inventory of the whole project

Importance: **High** = the chain breaks without it · **Medium** = needed for a task / tool you still use · **Low** = reference only.

### Chain (create BOM -> assign components -> production version)

| File | Role | Used by | Importance |
|---|---|---|---|
| `workbook.py` | chain file path, sheet names, shared column names, safe Excel writes | all 3 phases | **High** |
| `connect.py` | SAP session, `slow` mode | every SAP script | **High** |
| `bom.py` | shared cleaning (`clean_component`, `clean_text`, ...), `press_enter`, `save_bom`, CS02 helpers | phase 1, 2, 3, `main.py`, `update_baseqty.py` | **High** |
| `models.py` | `BomHeader`, `BomItem`, `Routing`, `ProductionVersion` | all phases | **High** |
| `create_main.py` | **phase 1 runner** (CS01) | you | **High** |
| `reader.py` | reads + validates Sheet1 | phase 1 | **High** |
| `bom_create.py` | CS01 screen steps | phase 1 | **High** |
| `routing_main.py` | **phase 2 runner** (CA02), tasks `newbom` + `altbom` | you | **High** |
| `routing.py` | assignment rules + CA02 screen steps | phase 2, phase 3 (`is_trial`, cell helpers) | **High** |
| `routing_sf.py` | Snowflake queries | phase 2 (`load_routing_list`, `load_assignment_worklist`). `load_routings` is unused today, kept for a future standalone PV task | **High** |
| `production_version.py` | **phase 3 runner** (MM02) | you | **High** |
| `docs/workflow.md`, `docs/routing.md` | how the chain / routing works | you | **High** |
| `docs/TODO.md` | known bugs, open decisions, plans | you | **High** |
| `docs/images/*.png` | CA02 screenshots used in `routing.md` | docs | Low |

### Standalone tools (not part of the chain, still work)

| File | Role | Notes | Importance |
|---|---|---|---|
| `excel_utils.py` | `load_excel` (main.py, update_baseqty.py, update_workcenter_efficiency.py), `unmerge_and_fill` (unmerge_excel.py) | trimmed to what is used | **Medium** |
| `unmerge_excel.py` | unmerges cells of a source file | points to `CHANGE_BOM_demo.xlsx`, not to `CREATE_BOM.xlsx`; change `INPUT_FILE` before using it for the chain | **Medium** |
| `main.py` | CS02 change BOM | saves its input file with **one sheet** only -> never point it at `CREATE_BOM.xlsx` | Medium |
| `update_baseqty.py` | change base quantity in CS02 | affected by bug A4 (Enter not pressed) | Medium |
| `update_workcenter_efficiency.py` | work center efficiency (Set MU APS.xlsx) | own save logic | Low |
| `TODO.md` (root), `note.md` | CS02 search ideas (TODO-1), notes on `bom.py` scrolling | still valid for `bom.py` | Low |
| `ca02.vbs`, `mm02.vbs`, `altbompopup.vbs` | SAP GUI recordings the Python steps were built from | reference for screen IDs | Low |

### Archived (this folder)

| File | Importance |
|---|---|
| `production_version_fix.py`, `removed_code.py`, `docs/routing_altbom_plan.md` | Low (history only) |
