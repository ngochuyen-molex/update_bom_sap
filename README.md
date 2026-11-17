# change_bom - SAP BOM / routing / production version automation

Python SAP GUI scripting (`win32com`) for plant **1901**: create BOMs (CS01), assign BOM
components to routing operations (CA02) and create production versions (MM02), driven by
one Excel workbook. Some data (routings, work centers, current allocation) is read from
Snowflake (`MLX_DATALAKE.MLX_SAP_ECC`); everything written goes through the SAP GUI.

> **Close the workbook in Excel before every run.** The scripts write their results into it
> while they run; an open file is locked by Windows and the save fails. If that happens the
> SAP changes are already made - close the file and run `continue`.

## The chain

Three phases, **one workbook** (`config.CHAIN_FILE`, today `CREATE_BOM.xlsx`):

```
CREATE_BOM.xlsx
├─ Sheet1               source (you) + phase 1 result columns   bom_create_main.py
├─ Routing Result       phase 2                                  routing_main.py
└─ Prod Version Result  phase 3                                  production_version.py
```

| Phase | Script | SAP | Snowflake | Writes |
|---|---|---|---|---|
| 1 Create BOM | `bom_create_main.py` | CS01 | no | result columns in `Sheet1` |
| 2 Assign components | `routing_main.py` | CA02 | routings + op 0030 WC | `Routing Result` |
| 3 Production version | `production_version.py` | MM02 | no | `Prod Version Result` |

Each phase only works on the rows the previous phase finished. A row that is `SUCCESS` is
never processed again, so a rerun is always safe.

## Quick start

```powershell
cd C:\Users\nguyeb3\Downloads\SAP_automation\change_bom

python run_all.py check        # all three phases, no SAP: build + validate the work lists
python run_all.py              # all three phases in SAP, one after the other
```

Or run the phases one by one (same thing, more control):

```powershell
python bom_create_main.py check         # 1. validate the source
python bom_create_main.py              #    create BOMs (a rerun continues)

python routing_main.py check           # 2. routing work list -> Routing Result
python routing_main.py                 #    assign components
python routing_main.py continue        #    after a break / for FAILED rows

python production_version.py check     # 3. routing list from Routing Result
python production_version.py           #    create production versions
python production_version.py continue  #    after a break / for FAILED rows
```

## Options

| Option | Meaning |
|---|---|
| *(none)* | do the work in SAP |
| `check` | build / validate and write the result sheet, **no SAP** |
| `continue` (`retry`) | process `PENDING` / `FAILED` rows of the result sheet (phase 2: no Snowflake) |
| `fresh` | rebuild the work list although a result sheet exists (`SUCCESS` rows are kept) |
| `slow` | wait after every SAP action - use when a screen is missed |
| `from=2` \| `from=3` | `run_all.py` only: start at that phase |
| `altbom` | `routing_main.py` only: the second task (see below) |
| `test` | `production_version.py` only: one hard-coded version, no Excel |

Not every option applies to every phase - the full matrix is in
[`docs/workflow.md`](docs/workflow.md) section 1.

## Status values

| STATUS | Meaning | Processed by |
|---|---|---|
| `PENDING` | to do | normal run, `continue` |
| `SUCCESS` | done and saved | never again |
| `SKIPPED` | not done on purpose; reason in MESSAGE | - |
| `FAILED` | error; text in MESSAGE | `continue` |

Result sheets are saved during the run and when it stops (Ctrl+C, crash). Saving writes a
temp copy and replaces the file, so a crash cannot leave the workbook half-written.

## Two routing tasks

`routing_main.py` has the same SAP logic for two different work lists, because
**Snowflake lags SAP by about 20 minutes**:

| | `newbom` (default, the chain) | `altbom` |
|---|---|---|
| Use when | BOMs were just created in phase 1 | checking / completing **existing** BOMs |
| Alt BOMs | only the new one, from `Sheet1` | all active ones, from Snowflake |
| Command | `python routing_main.py` | `python routing_main.py altbom` |

Don't run `altbom` on BOMs created less than ~20 min ago - the new alt BOM would be missing
from the work list with no error. Details: [`docs/workflow.md`](docs/workflow.md) section 0.

## Standalone tools (not part of the chain)

| Script | Does | Input (in `config.py`) |
|---|---|---|
| `bom_update_main.py` | change component / quantity in CS02 | `CHANGE_BOM_FILE` |
| `update_baseqty.py` | change the header base quantity in CS02 | `BASE_QTY_FILE` |
| `update_workcenter_efficiency.py` | work center efficiency / utilization in CR02 | `WORKCENTER_FILE` |
| `unmerge_excel.py` | unmerge cells and fill a source file down | `UNMERGE_FILE` |

> `bom_update_main.py` and `update_workcenter_efficiency.py` save their input file with
> **one sheet only**. Never point them at the chain workbook - it would wipe
> `Routing Result` and `Prod Version Result`.

## Configuration

All paths and fixed SAP values are in [`config.py`](config.py) - change them there and
nowhere else: the input files, `SAP_CONNECTION_NAME` (`M04 [SSO]`), `PLANT` (`1901`),
`BOM_USAGE`, the Snowflake connection name from `connections.toml`, and `SAVE_EVERY`.

Needs: SAP Logon with **SAP GUI scripting enabled**, `pywin32`, `pandas`, `openpyxl`,
`snowflake-connector-python`.

## Documentation

| Doc | Content |
|---|---|
| [`docs/workflow.md`](docs/workflow.md) | the chain in full: phases, columns, Snowflake lag, status values |
| [`docs/routing.md`](docs/routing.md) | CA02 component assignment: rules, screens, both tasks |
| [`docs/change_bom_rate.md`](docs/change_bom_rate.md) | CS02 change algorithm and row matching |
| [`docs/TODO_bug.md`](docs/TODO_bug.md) | known bugs, open decisions, plans |
| [`docs/TODO.md`](docs/TODO.md) | CS02 search ideas and refactor notes |
| [`archive/README.md`](archive/README.md) | file inventory of the whole project + what was archived |
