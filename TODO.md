# TODO

## TODO-1: Search the first screen first, scroll only if not found

**Status:** Not started (idea / intention)
**Files:** `bom.py` -> `find_item_row()`, `find_component_row()`, `scroll_bom_table()`
**Background:** see `note.md`, sections 1-3

### Goal
Search the rows already visible on screen first, without scrolling, like the
original code did. Only if the Item / Component is not found there,
scroll page by page using `scroll_bom_table()` (the current behaviour).

Expected benefits:
- Fewer scroll actions, so it is faster and there are fewer chances for SAP GUI
  errors.
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
   `scroll_bom_table()` (see `note.md`, section 2).
5. The function must still return a **visible** row index (see `note.md`,
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
> In `change_bom/bom.py`, implement TODO-1 from `change_bom/TODO.md`:
> change `find_item_row()` and `find_component_row()` to search the rows on the
> current screen first without scrolling, and only fall back to the existing
> scroll search (`scroll_bom_table()`) if not found. Follow the
> "Things to watch out for" and "Suggested design" sections, keep returning a
> visible row index, and read `change_bom/note.md` for background.
