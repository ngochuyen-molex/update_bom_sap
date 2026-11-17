# Archive + file inventory

Cleanup on 2026-11-16. Nothing was deleted: items no longer used were moved here.
Nothing in this folder is imported by the active code.

Inventory re-checked against the actual files on **2026-11-17**; entries for files that are
no longer in the repo are marked below.

## 1. What is in this folder

| Item | Was | Why archived |
|---|---|---|
| `production_version_fix.py` | one-off fix (rename / create production versions of the first half-finished PV run) | the job is done; it is also broken: it reads `create_bom_result.xlsx`, which has no Group/Version columns (bug A3), and it imported `INPUT_FILE` / `RESULT_FILE` from `production_version.py`, which were removed. **Will not run from here.** |
| `removed_code.py` → `save_excel`, `prepare_result_columns`, `log_result`, `add_log`, `save_progress` | `excel_utils.py` | old single-sheet result helpers, replaced by `workbook.py`. `save_progress` writes the file with **one sheet only** -> would wipe `Routing Result` / `Prod Version Result` if used on the chain file |
| `removed_code.py` → `print_boms` | `reader.py` | debug print, never called (and fails on an empty quantity) |
| (deleted lines) `INPUT_FILE` / `RESULT_FILE` | `production_version.py` | only used by the fix script |
| ~~`docs/routing_altbom_plan.md`~~ | design notes of the altbom task | **not in the repo** (2026-11-17): listed as archived here but the file is absent. Current behaviour is in `docs/routing.md` + `docs/workflow.md` |

To bring something back: copy it back to its file. Don't run anything from this folder directly.

## 2. Inventory of the whole project

Importance: **High** = the chain breaks without it · **Medium** = needed for a task / tool you still use · **Low** = reference only.

### Chain (create BOM -> assign components -> production version)

| File | Role | Used by | Importance |
|---|---|---|---|
| `README.md` (project root) | entry point: what this is, quick start, options, standalone tools | you | **High** |
| `config.py` | all input file paths + fixed SAP / Snowflake values, `SAVE_EVERY` | everything | **High** |
| `workbook.py` | sheet names, shared column names, safe Excel writes; re-exports `CHAIN_FILE` from `config.py` | all 3 phases | **High** |
| `connect.py` | SAP session, `slow` mode | every SAP script | **High** |
| `utils.py` | the cleaners (`clean_component`, `clean_item`, `clean_number`, `clean_text`, `numbers_equal`) | `bom.py` and all phases | **High** |
| `bom.py` | `press_enter`, `save_bom`, CS02 helpers; re-exports the cleaners from `utils.py` | phase 1, 2, 3, `bom_update_main.py`, `update_baseqty.py` | **High** |
| `models.py` | `BomHeader`, `BomItem`, `Routing`, `ProductionVersion` | all phases | **High** |
| `run_all.py` | runs phases 1, 2, 3 one after the other (`check`, `slow`, `from=N`) | you | **High** |
| `bom_create_main.py` | **phase 1 runner** (CS01) | you | **High** |
| `reader.py` | reads + validates Sheet1 | phase 1 | **High** |
| `bom_create.py` | CS01 screen steps | phase 1 | **High** |
| `routing_main.py` | **phase 2 runner** (CA02), tasks `newbom` + `altbom` | you | **High** |
| `routing.py` | assignment rules + CA02 screen steps | phase 2, phase 3 (`is_trial`, cell helpers) | **High** |
| `routing_sf.py` | Snowflake queries | phase 2 (`load_routing_list`, `load_assignment_worklist`). `load_routings` is unused today, kept for a future standalone PV task | **High** |
| `production_version.py` | **phase 3 runner** (MM02) | you | **High** |
| `docs/workflow.md`, `docs/routing.md` | how the chain / routing works | you | **High** |
| `docs/TODO_bug.md` | known bugs, open decisions, plans | you | **High** |
| `docs/images/*.png` | CA02 screenshots used in `routing.md` | docs | Low |

### Standalone tools (not part of the chain, still work)

| File | Role | Notes | Importance |
|---|---|---|---|
| `unmerge_excel.py` | unmerges cells of a source file (`unmerge_and_fill` from `workbook.py`) | points to `UNMERGE_FILE` (`CHANGE_BOM_demo.xlsx`), not to `CREATE_BOM.xlsx`; change it in `config.py` before using it for the chain | **Medium** |
| `bom_update_main.py` | CS02 change BOM (component / quantity), input `CHANGE_BOM_FILE` | saves its input file with **one sheet** only -> never point it at `CREATE_BOM.xlsx` | Medium |
| `update_baseqty.py` | change base quantity in CS02, input `BASE_QTY_FILE` | affected by bug A4 (Enter not pressed) | Medium |
| `update_workcenter_efficiency.py` | work center efficiency in CR02, input `WORKCENTER_FILE` | own save logic, single sheet | Low |
| `docs/change_bom_rate.md` | CS02 change algorithm and row matching | describes `bom_update_main.py` | Low |
| `docs/TODO.md` | CS02 search ideas (TODO-1), refactor notes | still valid for `bom.py` | Low |

Gone from the repo (checked 2026-11-17), still named in older notes:
`excel_utils.py` (merged into `workbook.py`), `create_main.py` (now `bom_create_main.py`),
`main.py` (now `bom_update_main.py`), `docs/note.md` (CS02 scrolling background, cited by
`TODO.md` and by comments in `bom.py` / `bom_update_main.py`), and the
`ca02.vbs` / `mm02.vbs` / `altbompopup.vbs` SAP GUI recordings.

### Archived (this folder)

| File | Importance |
|---|---|
| `production_version_fix.py`, `removed_code.py` | Low (history only) |
