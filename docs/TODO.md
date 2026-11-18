# TODO

## TODO-1: Search the first screen first, scroll only if not found

**Status:** Not started (idea / intention)
**Files:** `bom.py` -> `find_item_row()`, `find_component_row()`, `scroll_bom_table()`
**Background:** see `change_bom_rate.md`, sections 1-3. (Older notes and code comments cited
a `note.md` for this; that file is **not in the repo** (checked 2026-11-17), but
`change_bom_rate.md` sections 1, 2 and 3 cover exactly the same three topics in the same
order - item vs component search, scrolling / invalid references, absolute vs visible row -
so the old section numbers still apply.)

### Goal
Search the rows already visible on screen first, without scrolling, like the
original code did. Only if the Item / Component is not found there,
scroll page by page using `scroll_bom_table()` (the current behaviour).

Expected benefits:
- Fewer scroll actions, so it is faster and there are fewer chances for SAP GUI errors.
- The original simple logic is used for small BOMs, which are most of them.

### Current behaviour
Both functions always start by scrolling to position 0 and then go page by page.
`find_component_row()` always scans every page, because it must check for
duplicates.

### Original code (for reference)
```python
def find_item_row(session, target_item):
    table = get_bom_item_table(session)
    target_item = clean_item(target_item)

    for row_idx in range(table.RowCount):
        current_item = get_item(session, row_idx)
        if current_item == target_item:
            return row_idx

    raise ValueError(f"Item {target_item} not found in BOM table.")
```

### Things to watch out for (why it is not done yet)
1. **`RowCount` vs `VisibleRowCount`:** the original loop goes up to `RowCount`
   (all rows in the BOM), but SAP cell IDs like `txt[0,row_idx]` only exist for
   rows on screen (0 .. `VisibleRowCount - 1`). With a long BOM, the original
   code would fail or read the wrong rows. The first-screen pass must loop only
   to `VisibleRowCount`.
2. **The table may already be scrolled:** after an earlier action the table might
   not be at the top. Make sure the "first screen" really starts at row 0
   (check `VerticalScrollbar.Position == 0`), or accept that whatever is on
   screen is searched first.
3. **Duplicate check for components:** finding the component on the first screen
   does not prove it is unique. The rows below could contain it again.
   Options:
   - a) Return the first match on screen (faster, but it could pick the wrong
     line if there is a duplicate further down).
   - b) If the whole BOM fits on one screen (`RowCount <= VisibleRowCount`),
     no scrolling is needed and the duplicate check is complete. Only scroll
     when the BOM is longer than one screen.  <- recommended
4. **Never reuse old table/scrollbar references after scrolling.** Always use
   `scroll_bom_table()` (see `change_bom_rate.md`, section 2).
5. The function must still return a **visible** row index (see `change_bom_rate.md`,
   section 3).

### Suggested design
```text
find_row(target):
    table = get_bom_item_table(session)

    # Pass 1: no scrolling, current screen only
    FOR i in 0 .. VisibleRowCount - 1:
        IF row i matches target: remember it

    IF item search AND found:                  return i
    IF component search AND found
       AND whole BOM fits on one screen:       check duplicates, return i

    # Pass 2: fall back to the existing page-by-page scroll search
    return existing scrolling logic (scroll_bom_table)
```

### Test after implementing
- A BOM where the target is on the first screen (should not scroll).
- A BOM where the target is further down (must scroll and still find it).
- A component that appears twice (must still fail with "matches multiple BOM rows").
- Run on demo data first and back up the Excel file.

### Prompt to reuse with an AI tool
> In `change_bom/bom.py`, implement TODO-1 from `change_bom/docs/TODO.md`:
> change `find_item_row()` and `find_component_row()` to search the rows on the
> current screen first without scrolling, and only fall back to the existing
> scroll search (`scroll_bom_table()`) if not found. Follow the
> "Things to watch out for" and "Suggested design" sections, keep returning a
> visible row index, and read `change_bom/docs/change_bom_rate.md` sections 1-2 for
> background (`note.md` is referenced in older notes but is not in the repo).


## TODO-3: One BOM entry point, action chosen per row in the input Excel

**Status:** Not started (needs the request form below to be agreed first)
**Files today:** `bom_create_main.py` (CS01), `bom_update_main.py` (CS02 component/qty),
`update_baseqty.py` (CS02 header base qty), `reader.py` (`decide_action`, `validate_boms`)

### Problem
Three scripts do the same frame (read Excel, loop, log STATUS/MESSAGE, save, resume) and
differ only in the SAP action. They use different input files, different Excel helpers and
different result handling:

| Script | Input file | Excel helper | Result written |
|---|---|---|---|
| `bom_create_main.py` | `CREATE_BOM.xlsx` (chain file) | `workbook.py` (safe save, resume) | yes, Sheet1 columns |
| `bom_update_main.py` | `CHANGE_BOM_list.xlsx` | own `add_log` / `save_progress` | yes, overwrites file, no lock handling |
| `update_baseqty.py` | `create_bom_final.xlsx` | `workbook.load_excel` | **no** (only printed) |

Known issues this fixes:
- `update_baseqty.py` groups by `Base Qty after`, so one BOM with two different base qty
  values is updated twice without warning.
- `update_baseqty.py` has no result file: after a run you cannot see which BOMs failed.
- `workbook.load_excel` (old `excel_utils.load_excel`) reads without `dtype=str`, so values
  like `0004` / leading-zero materials can lose zeros. Switch to `workbook.read_sheet` here
  and remove `load_excel` when the old scripts are gone.
- `reader.decide_action` defaults to CHANGE when alt BOM = 1 and Action is empty. A blank
  cell silently decides what is posted to SAP.
- **Open question, still unanswered:** what base quantity a newly created BOM should get.
  Today CREATE takes `Base Qty after` from the request form; it is not derived or checked
  against the routing. Decide this before CREATE and the base-qty change share one path.

### Target design
- One script / CLI command for all BOM work: `python cli.py bom <file>`.
- `Action` column is **required**: `CREATE` or `CHANGE`. Blank or anything else = row error
  before SAP is touched (remove the alt-BOM-1 default).
- CHANGE does only what is filled in the "after" columns:
  - `Base Qty after` filled -> update header base qty (replaces `update_baseqty.py`)
  - `Component after` / `BOM rate after` filled -> update the item (today's `bom_update_main.py`)
  - both filled -> both, one save
- Dispatch by dict, not if/elif:
  ```python
  BOM_ACTIONS = {"CREATE": create_bom, "CHANGE": change_bom}
  ```
- Results always go through `workbook.py` (STATUS / MESSAGE / LOG / Updated At, resume with
  `continue`, retry FAILED, lock handling).

### Pre-validation rules (in `validate_boms`, before SAP)
- Action missing / invalid -> error.
- CREATE: base qty required, at least one item, every item has component + quantity + unit,
  no duplicate component in the same BOM.
- CHANGE: at least one "after" value filled, otherwise error "nothing to change".
- CHANGE item: `Item` or `Component before` must be filled to find the line
  (see `change_bom_rate.md` section 1). `Component before` is checked against SAP before changing.
- Same Parent Material + Alt BOM with different `Base Qty after` values -> error.
- Same Parent Material + Alt BOM with mixed actions (CREATE and CHANGE) -> error.

### Request form: what requesters must fill in (one row per BOM item)

| Column | CREATE | CHANGE (item) | CHANGE (base qty only) | Notes |
|---|---|---|---|---|
| Action | required | required | required | `CREATE` or `CHANGE` |
| Parent Material | required | required | required | on every row (no merged cells, or run unmerge first) |
| Alternative BOM | required | required | required | for CREATE: the alt BOM that should be created |
| Item | - | recommended | - | needed when `Component before` appears twice in the BOM |
| Component before | - | required if Item blank | - | must match SAP, otherwise the row fails |
| BOM rate before | - | recommended | - | for checking (qty check currently disabled in bom.py) |
| Component after | required | optional | - | blank = keep component |
| BOM rate after | required | optional | - | blank = keep quantity |
| Unit after | required | - | - | |
| Component Scrap after | optional | - | - | blank = 0 |
| Base Qty after | required | optional | required | same value on every row of one BOM |

General rules for the form:
- Plant 1901, BOM usage 1 (fixed). Requests for other plants/usages need a code change.
- Materials as text (keep leading zeros), no formulas, no merged cells if possible.
- One Excel file per request; do not rename the column headers.

### Test after implementing
- One file with a CREATE BOM, a CHANGE item, a CHANGE base qty only, and an invalid row.
- Invalid rows must be reported by pre-validation and nothing posted for them.
- `continue` after stopping halfway only processes PENDING / FAILED rows.
- Run on demo data first and back up the Excel file.

### Prompt to reuse with an AI tool
> Implement TODO-3 from `change_bom/docs/TODO.md`: merge `bom_create_main.py`,
> `bom_update_main.py` and `update_baseqty.py` into one BOM entry point that reads the
> request-form columns, validates them with the listed rules, dispatches per row with
> `BOM_ACTIONS`, and writes results through `workbook.py`. Keep the CS02 item-finding logic
> unchanged (described in `change_bom_rate.md` sections 1-3; older notes cite a `note.md`
> that is not in the repo).


## TODO-4: Folder layout per SAP object

**Status:** Not started (do after TODO-3)
**Done 2026-11-16:** `config.py` (paths, plant, usage, connections, SAVE_EVERY), `utils.py`
(clean_* + numbers_equal, `bom.py` re-exports them), `excel_utils.py` merged into `workbook.py`
and deleted. Still in the flat folder; the folder move below is not done.

### Target layout
```
change_bom/
  config.py              # paths, plant, BOM usage, Snowflake connection name, SAVE_EVERY
  cli.py                 # one entry: bom | routing newbom | routing altbom | pv | workcenter | unmerge
  common/
    utils.py             # clean_text / clean_component / clean_number / clean_item, now()
    workbook.py          # all Excel I/O (+ unmerge_and_fill from excel_utils.py)
    runner.py            # generic run loop: todo filter, try/except per row, save every N, summary
    models.py
  sap/
    connect.py
    gui.py               # set_text, press_enter, status bar, find_row, double_click
  bom/        screens.py (CS01+CS02), reader.py, actions.py
  routing/    screens.py (routing.py), queries.py (routing_sf.py), worklists.py
  production_version/  screens.py, worklists.py
  workcenter/ efficiency.py
```

### Rules
- Lower layers never import upper ones: `common` and `sap` know nothing about BOM/routing.
- `bom/reader.py` must not import from SAP screen code (today it imports `clean_*` from `bom.py`).
- `routing_sf.py` queries: plant as a parameter, shared "latest change state / not deleted"
  filters written once.
- Move one folder at a time and run each phase on demo data after every move.