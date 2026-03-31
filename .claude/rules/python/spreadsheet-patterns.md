---
paths:
  - "spreadsheet_llm/**/*.py"
---
# Spreadsheet Processing Patterns

## Merged Cell Handling
- ALWAYS build a merged cell map (`{coord: start_coord}`) before processing cells
- NEVER iterate `sheet.merged_cells.ranges` per-cell — use pre-built dict for O(1) lookups
- When reading a merged cell's value, always redirect to the start cell of the merge range

## Style Fingerprinting
- Use `id(cell.font)` + `id(cell.border)` + `id(cell.fill)` as style fingerprint
- openpyxl interns style objects — `id()` is free and unique per distinct style
- Do NOT build custom hashable tuples from style attributes for comparison

## Candidate Generation
- Cap rectangular candidates to prevent O(n^4) combinatorial blowup
- Use sweep-line approach: consecutive boundary pairs × active column boundaries
- Sort by populated cell count before capping

## Edge Cases to Always Handle
- Empty sheets (max_row=0 or single None cell)
- Adjacent tables with no gap row (detect via style/data-type transitions)
- Merged headers spanning full width (merge flag creates boundary)
- Sheets with no formatted headers (fall back to text-first-row heuristic)
- Very large sheets (10K+ rows) — stream, chunk, never materialize all cells

## Testing
- Test with both small (10-row) and large (10K-row) sheets
- Test sheets with 0, 1, 2, and many tables
- Test adjacent tables, nested tables, and tables with merged headers
