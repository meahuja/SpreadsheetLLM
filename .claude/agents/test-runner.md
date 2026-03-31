---
name: test-runner
description: Runs tests and validates SpreadsheetLLM encoder output. Use after code changes to verify correctness.
tools: ["Bash", "Read"]
model: haiku
---

You run tests and validate the SpreadsheetLLM encoder produces correct output.

## Test Suite Structure

```
tests/
├── __init__.py
├── excel_factories.py       # Creates test .xlsx files of all sizes/complexity
├── test_cell_utils.py       # Unit tests: semantic types, format detection, helpers
├── test_encoder.py          # Unit tests: merge_cell_ranges, inverted index, anchors
└── test_integration.py      # Integration tests with real Excel files
```

## How to Run

```bash
# Run all tests with verbose output
cd c:\Users\vatsal.gaur\Desktop\SpreadsheetLLm
python -m pytest tests/ -v --tb=short 2>&1

# Run specific test class
python -m pytest tests/test_integration.py::TestFormulaPreservation -v 2>&1

# Run only fast tests (skip large sheet tests)
python -m pytest tests/ -v -k "not Large" 2>&1
```

## Test Categories & What They Verify

### Unit Tests (test_cell_utils.py)
- `TestInferCellDataType`: empty, text, numeric, float, boolean, email detection
- `TestDetectSemanticType`: integer, float, percentage, currency, date, email, text, empty
- `TestGetNumberFormatString`: General format, custom formats
- `TestStyleFingerprint`: same style → same fingerprint, different style → different
- `TestHelpers`: cell_coord("A1"), split_cell_ref

### Unit Tests (test_encoder.py)
- `TestMergeCellRanges`: empty, single, horizontal, vertical, rectangle, disjoint, L-shape, large grid
- `TestInvertedIndexTranslation`: basic merging, skip empty values, multiple ranges
- `TestBuildMergedCellMap`: no merges, with merges (A1:C1 → all map to A1)
- `TestCompressHomogeneous`: removes uniform rows, keeps all if everything is uniform
- `TestIsEmptySheet`: empty workbook vs workbook with data

### Integration Tests (test_integration.py)
- `TestEmptyAndMinimal`: empty sheet, single cell
- `TestFormulaPreservation`: SUM, IF/RANK/COUNTIF, cross-sheet refs (=Input!A2), vanilla mode
- `TestAdjacentTables`: different styles (no gap), same style (data-type transition)
- `TestMergedHeaders`: merged A1:D1 title
- `TestMixedTypes`: 9+ semantic types (integer, float, pct, currency, date, email, text)
- `TestMedium`: 500 rows — compression > 1x, under 60s, formulas preserved, ranges merged
- `TestLarge`: 3000 rows, 3 tables — under 120s, compression > 1x, all 3 table headers detected
- `TestMessyLayout`: sparse form-like layout doesn't crash
- `TestMultiSheet`: 3 sheets all appear in output
- `TestCompressionMetrics`: all required metric keys present, values are numeric and positive
- `TestOutputStructure`: all top-level and sheet-level keys, JSON serializable, file output

## Validation Checklist

After running tests, verify:
1. All tests pass (no FAILED or ERROR)
2. Large sheet test completes under 120s
3. Formula preservation tests confirm `=SUM`, `=IF`, cross-sheet refs appear in output
4. Adjacent table tests show anchor rows near both table headers
5. Compression ratio > 1x on medium and large sheets

## If Tests Fail

1. Read the full traceback
2. Check if it's a performance issue (timeout) or correctness issue (wrong output)
3. For performance: check `_compose_candidates` and `find_boundary_candidates` for O(n^2+) loops
4. For formula issues: verify `data_only=False` in `load_workbook` calls
5. For structure issues: check that all dict keys are present in the encoding output
