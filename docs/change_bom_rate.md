# Logic Notes

## 0. Main algorithm (`main.py`)

```text
Connect to SAP (connect.py: get_session)
Open transaction CS02

Read Excel (INPUT_FILE)
Reset result columns: STATUS = "", MESSAGE = "", LOG = ""

FOR each row in dataframe:
    TRY:
        STATUS = "PROCESSING"

        # 1. Skip blank rows
        IF Parent Material, Alternative BOM and Item are all blank:
            STATUS = "SKIPPED", MESSAGE = "Blank row"
            CONTINUE

        # 2. Check if anything needs changing
        IF Component before == Component after
           AND BOM rate before == BOM rate after:
            STATUS = "SKIPPED", MESSAGE = "Component and Qty unchanged"
            CONTINUE

        # 3. Open BOM
        Open BOM (Parent Material, Alternative BOM, usage 1, plant 1901)

        # 4. Find the BOM row                          -> see section 1
        IF Item is filled:
            row = Find row by Item number
        ELSE:
            row = Find row by Component before
            (error if not found, or if found on more than one line)
        # Both searches scroll safely               -> see section 2
        # and return a visible row index            -> see section 3

        # 5. Validate current SAP data
        Read current Component from SAP
        IF SAP Component != Excel Component before:
            raise error "Component mismatch"

        (Qty check: SAP Qty != Excel BOM rate before
         -> currently DISABLED / commented out in bom.py)

        # 6. Update
        Update Component  = Component after
        Update Qty        = BOM rate after
        Press Enter twice (confirm SAP messages)
        Wait 2 seconds
        Save BOM

        STATUS = "SUCCESS", MESSAGE = "BOM saved successfully"

    EXCEPT any error (item not found, component mismatch, SAP error ...):
        STATUS = "FAILED", MESSAGE = error text
        (continue with next row)

    FINALLY:
        Save Excel progress (overwrites INPUT_FILE after every row)
        IF save fails: print "Make sure the Excel file is closed."
```

Notes:
- Every step writes a line to the `LOG` column, so you can see where a row stopped.
- Errors are not checked step by step. Any error raised during the row jumps to `EXCEPT`, which marks the row FAILED, and the script moves on to the next row.
- Because Excel is saved after every row, progress is kept even if the script stops halfway.
- SAP is not closed at the end of the run.

## 1. Finding the BOM row: by Item or by Component

In `main.py`, each Excel row is matched to a SAP BOM line:

- If `Item` has a value, it searches by item number (`find_item_row()` in `bom.py`).
- If `Item` is empty or NaN, it searches by `Component before` (`find_component_row()` in `bom.py`).

`clean_item()` turns a blank or NaN Item into `""`, and an empty string counts as false, so the script falls back to component search.

| Item value in Excel | `clean_item()` result | Search used |
|---|---|---|
| Empty cell / spaces | `""` | Component |
| `10` or `10.0` | `"0010"` | Item |
| `1` | `"0010"` (single digit x10) | Item |
| `0` | `"0000"` | Item (will usually fail) |

Notes:
- Only the BOM that is currently open is searched. Different materials can share the same component without any problem.
- If `Component before` appears on more than one line of the same BOM, the row fails on purpose so that the wrong line is never changed. Fill in `Item` for those rows.

## 2. SAP table scrolling and invalid references (`scroll_bom_table()`)

### Problem
Python does not hold the SAP table itself; it holds a COM reference (a link) to the table on SAP's screen. When the script scrolls, SAP sends a request to the server and redraws the screen. Links grabbed before the scroll (for example `table.VerticalScrollbar`) can then point to screen objects SAP has already thrown away. Using them can make SAP GUI fail:

- `The remote procedure call failed` on the current row, and
- `The RPC server is unavailable` on every following row, because the session is dead.

### Why only component search failed
- `find_item_row()` stops as soon as it finds the item. When the item is on the first screen, it never reuses an old link.
- `find_component_row()` must read every page to check for duplicates, so it always scrolls more than once and reused the old link. That caused the crash.

### Fix
`scroll_bom_table()` calls `session.findById(...)` before every scroll and every read, so it always uses a fresh link to the table as it exists now:

```python
def scroll_bom_table(session, position):
    get_bom_item_table(session).VerticalScrollbar.Position = position
    return int(get_bom_item_table(session).VerticalScrollbar.Position)
```

Both search functions use this helper. Reading cells (`get_item`, `get_component`) was already safe because each call looks the cell up freshly with `findById`.

## 3. Absolute row vs visible row (returning to the matched row)

The SAP table only shows a few rows on screen at a time, and the script can only read or edit rows that are on screen.

- **Absolute row**: the row's position in the whole BOM (0, 1, 2, ...).
- **Visible row**: the row's position on the screen right now (0 = top of the screen).

Functions such as `update_component()` write to `ctxt[2,{row_idx}]`, where `row_idx` must be a **visible** row.

### Example: a BOM with 20 rows, 8 fit on screen

| Page | Scroll position | Absolute rows shown |
|---|---|---|
| 1 | 0 | 0-7 |
| 2 | 8 | 8-15 |
| 3 | 12 | 12-19 (last page is aligned to the bottom) |

Suppose the component is at absolute row 10:

1. The scan finds it on page 2 and records `(absolute_row=10, page_start=8)`.
2. The scan continues to page 3 (to check for duplicates), so row 10 is no longer on screen.
3. `scroll_bom_table(session, 8)` scrolls back and returns the position SAP actually landed on (8).
4. Visible row = 10 - 8 = **2**, the third row on screen.
5. That visible row is returned to `main.py` and used for validation and the update.

```python
visible_row = absolute_row - scroll_bom_table(session, page_start)
```

The script uses the position SAP actually lands on, not the requested one, because SAP does not always land exactly where you ask (especially near the end of the table).

### Safety checks afterwards
- If `visible_row` is not within the rows on screen, it raises an error.
- It reads the component at `visible_row` again; if it does not match, it raises "No update was made". This way a wrong row is never changed.
