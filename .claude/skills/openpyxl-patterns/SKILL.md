---
name: openpyxl-patterns
description: openpyxl best practices for reading large Excel files efficiently. Activate when working with Excel I/O.
origin: project
---

# openpyxl Best Practices

## When to Activate

- Reading or writing Excel files with openpyxl
- Optimizing spreadsheet processing performance
- Handling merged cells, styles, or number formats

## Read-Only Mode for Large Files

```python
# Use read_only=True for profiling passes — streams rows, low memory
wb = openpyxl.load_workbook(path, read_only=True)
for row in ws.iter_rows():
    # process row — cells are read on demand
    pass
wb.close()  # MUST close read-only workbooks

# Switch to normal mode only for cells you need to keep
wb = openpyxl.load_workbook(path, data_only=False)
```

## Style Object Interning

openpyxl interns style objects — cells with identical styles share the same object.

```python
# FAST: Use id() for style fingerprinting
style_key = (id(cell.font), id(cell.border), id(cell.fill))

# SLOW: Don't build hashable tuples from attributes
style_key = (cell.font.bold, cell.font.italic, ...)  # Avoid this
```

## Merged Cell Map

```python
# Build once, use everywhere — O(1) lookups
def build_merged_cell_map(sheet):
    merged_map = {}
    for merge_range in sheet.merged_cells.ranges:
        start = merge_range.start_cell.coordinate
        for cell in merge_range.cells:
            coord = f"{get_column_letter(cell[1])}{cell[0]}"
            merged_map[coord] = start
    return merged_map
```

## Number Format Strings

```python
# Access via cell.number_format — returns string like "#,##0.00"
nfs = cell.number_format  # "General" if unset

# cell.data_type gives: 's' (string), 'n' (numeric), 'b' (bool),
# 'd' (datetime), 'e' (error), 'f' (formula)
```

## Performance Tips

- `sheet.max_row` / `sheet.max_column` — use instead of iterating to find bounds
- `get_column_letter(col_idx)` / `column_index_from_string(letter)` for conversions
- `sheet.cell(row=r, column=c)` is faster than `sheet['A1']` for known coordinates
- Avoid `sheet.iter_rows(values_only=True)` when you need cell objects (styles, formats)
