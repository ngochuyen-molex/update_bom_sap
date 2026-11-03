# CA02 - Assign new BOM to routings

Assigns the components of the newly created BOM (from `create_main.py`) to
operation `0020` in every routing of the material, using CA02.

| File | Role |
|---|---|
| `routing_main.py` | Runner. Reads the Excel, loops materials, writes results. **Run this.** |
| `routing.py` | SAP screen steps (CA02). `python routing.py` tests one material. |
| `routing_sf.py` | Optional Snowflake check (`CHECK_SNOWFLAKE`). |
| `models.py` | `RoutingTask`, `RoutingGroup` dataclasses. |

Input: result file of `create_main.py` (rows with `STATUS = SUCCESS`).
Output: `..._routing.xlsx` with `ROUTING STATUS`, `ROUTING MESSAGE`, `ROUTINGS DONE`.

---

## Screen 1 - CA02 initial screen

Enter material and plant, press Enter.

| Field | Element ID | Code |
|---|---|---|
| Material | `wnd[0]/usr/ctxt[0]` | `open_routing()` |
| Plant | `wnd[0]/usr/ctxt[1]` | `open_routing()` |

---

## Screen 2 - Popup "Entry Option: Group"

![Entry Option: Group popup](images/ca02_group_popup.png)

Shown when the material has routings in **more than one group**. Lists every
group and group counter (GrC) of the material.

- Table: `wnd[1]/usr/tblSAPLCPSLTCTRL_2110`
- **Group** is only filled on the first row of each group. Blank = same group as the row above.
- In this example:

  | Popup row | Group | GrC |
  |---|---|---|
  | 0 | M0164412 | 22 |
  | 1 | (blank -> M0164412) | 27 |
  | 2 | (blank -> M0164412) | 30 |
  | 3 | M5016441 | 13 |

  is read as:
  ```
  M0164412 (popup row 0): 22, 27, 30
  M5016441 (popup row 3): 13
  ```

| Column | Content | Constant in `routing.py` |
|---|---|---|
| `ctxt[0,row]` | Group (PLNNR) | `POPUP_GROUP_COL = 0` (verified) |
| `txt[1,row]` | GrC (PLNAL) | `POPUP_COUNTER_COL = 1` (verified) |
| `txt[2,row]` | Task list description | `POPUP_DESC_COL = 2` (from recording) |

Code:
- `read_group_popup()` - reads all rows into `RoutingGroup` objects (stops at first blank row).
- `select_popup_row()` - double-clicks the first row of a group to open it.

If the material has only **one group**, SAP skips this popup and goes straight to Screen 3.

---

## Screen 3 - Change Routing: Overview (Header overview)

![Change Routing: Overview](images/ca02_header_overview.png)

After selecting a group in Screen 2, only the routings (GrC) **of that group** are shown.

- Table: `wnd[0]/usr/tblSAPLCPDITCTRL_1200`

| Column | Content | Constant in `routing.py` |
|---|---|---|
| `txt[0,row]` | GrC | `HEADER_COUNTER_COL = 0` (verified) |
| `txt[1,row]` | Task list description | `HEADER_DESC_COL = 1` (from recording) |

Code:
- `get_header_counters()` - builds `{GrC: row}`, e.g. `{22: 0, 27: 1, 30: 2}`.
- `open_header_row()` - double-clicks a routing to open it.
- `back_to_header()` - Back (`wnd[0]/tbar[0]/btn[3]`) x3 returns here after each routing.

Each GrC from Screen 2 is matched to its row here. A GrC missing on this screen = error.

---

## Screen 4 - BOM selection popup (`btn[7]`)

*No screenshot yet.*

Opened with `wnd[0]/tbar[1]/btn[7]`. The new BOM is double-clicked at
`wnd[1]/usr/sub/1/txt[1,8]` (row 8, from recording, `BOM_ROW` in `routing_main.py`).

Code: `open_component_allocation()`

---

## Screen 5 - Component allocation

*No screenshot yet.*

- Table: `wnd[0]/usr/tblSAPLCMDITCTRL_1000`
- Column `txt[10,row]` = operation the component is already assigned to (`COMP_OPERATION_COL`).

Rules (`select_unassigned_components()`):
- Stop at the first blank row (columns 0-9 all empty).
- `txt[10]` empty -> select the row.
- `txt[10]` has a value (`0030`, `0040`, ...) -> leave it.

Assign (`assign_to_operation()`):
1. `wnd[0]/tbar[1]/btn[5]`
2. `wnd[1]/usr/txt[0]` = `0020`
3. `wnd[1]/tbar[0]/btn[0]` (OK), then Enter

---

## Flow

```
for each material (STATUS = SUCCESS in Excel):
    CA02 -> Screen 1 -> Screen 2: read all groups + GrC
    (optional) compare with Snowflake, log differences

    for each group:
        CA02 -> Screen 1 -> Screen 2: double-click group's first row
        Screen 3: read {GrC: row}
        for each GrC of the group:
            Screen 3: open GrC
            Screen 4: pick BOM (row 8)
            Screen 5: select unassigned rows, assign 0020
            Back x3 -> Screen 3
        Save (wnd[0]/tbar[0]/btn[11])

    write STATUS / MESSAGE / ROUTINGS DONE to Excel
```

Saved once per group: each group is a separate CA02 entry, and `/NCA02` discards unsaved changes.

---

## Open points

- Column indices of Screens 2 and 3 are verified with the Scripting Tracker.
- Component allocation blank-row check (columns 0-9 empty) is not verified.
- BOM row in Screen 4 is fixed at 8.
- Screenshots for Screens 1, 4 and 5 still to add (`docs/images/`).
