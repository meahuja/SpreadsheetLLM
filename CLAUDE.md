# SpreadsheetLLM — Project Reference for Claude

## What this project is

A **C# .NET port** of the SpreadsheetLLM research paper (arXiv:2407.09025).  
Compresses Excel spreadsheets into compact representations suitable for LLM input.  
The original Python implementation lives in `spreadsheet_llm/`; the .NET port is in `SpreadsheetLLM.Core/`.

---

## Directory layout

```
SpreadsheetLLM-main/
├── SpreadsheetLLM.Core/          ← .NET library (netstandard2.0, C# 9)
│   ├── SheetCompressor.cs        ← Main 3-stage pipeline
│   ├── ExcelReader.cs            ← ClosedXML-based xlsx reader
│   ├── CellUtils.cs              ← Cell type/format/semantic classification
│   ├── ChainOfSpreadsheet.cs     ← QA pipeline (Anthropic/OpenAI backends)
│   ├── VanillaEncoder.cs         ← Baseline row-major encoder
│   ├── IsExternalInit.cs         ← Polyfill for C# 9 init on netstandard2.0
│   └── Models/
│       ├── CellData.cs           ← Cell value, formula, style snapshot
│       ├── WorksheetSnapshot.cs  ← 2-D grid, 1-based row/col indexing
│       ├── SheetEncoding.cs      ← JSON output schema
│       └── CompressionMetrics.cs ← Per-sheet and overall metrics
├── SpreadsheetLLM.TestRunner/    ← Console test runner (net9.0)
│   └── Program.cs                ← Generates 10 sample sheets & encodes them
├── spreadsheet_llm/              ← Python reference (encoder.py, cell_utils.py…)
├── tests/                        ← Python unit tests
├── plain_adjacent.xlsx           ← Sample Excel file
└── CLAUDE.md                     ← This file
```

---

## Build & run

```powershell
# Build
& "C:\Program Files\dotnet\dotnet.exe" build SpreadsheetLLM.Core/SpreadsheetLLM.Core.csproj -c Release

# Run test runner (generates + encodes 10 sample workbooks)
& "C:\Program Files\dotnet\dotnet.exe" run --project SpreadsheetLLM.TestRunner -c Release

# JSON output lands in: SpreadsheetLLM.TestRunner/bin/Release/net9.0/test_output/
```

---

## NuGet dependencies

| Project | Package | Version | Purpose |
|---------|---------|---------|---------|
| Core | ClosedXML | 0.102.2 | Read/write .xlsx |
| Core | System.Text.Json | **8.0.5** | JSON serialisation (7.0.0 had CVE) |
| TestRunner | ClosedXML | 0.102.2 | Create sample .xlsx in tests |

---

## The 3-stage pipeline (`SheetCompressor.Encode`)

### Stage 1 — Structural-Anchor-Based Extraction (§3.2)
Identifies table boundaries via row/column fingerprints, header detection, and
data-type transitions. Runs NMS to select the best candidate regions, then
expands each anchor by k rows/cols (default k=2) and drops homogeneous regions.

Key constants: `MaxCandidates=200`, `MaxBoundaryRows=100`, `HeaderThreshold=0.6`,
`SparsityThreshold=0.10`, `NmsIouThreshold=0.5`.

### Stage 2 — Inverted-Index Translation (§3.3)
Builds a `value → [cell-refs]` map for kept cells, then merges adjacent refs
into ranges (`A1:A3`). Result: `Dictionary<string, List<string>>`.

### Stage 3 — Data-Format-Aware Aggregation (§3.4)
Groups cells by `{type, nfs}` key where type is one of the 9 semantic categories
(integer, float, currency, percentage, scientific, date, time, year, email) and
nfs is the raw Excel number-format string.

---

## JSON output shape

```json
{
  "file_name": "example.xlsx",
  "sheets": {
    "Sheet1": {
      "structural_anchors": { "rows": [1,2,6], "columns": ["A","B","C"] },
      "cells": { "Revenue": ["B2:B5"], "=SUM(B2:B5)": ["B6"] },
      "formats": { "{\"type\":\"currency\",\"nfs\":\"$#,##0.00\"}": ["B2:B5"] },
      "numeric_ranges": { "{\"type\":\"integer\",\"nfs\":\"General\"}": ["C2:C10"] }
    }
  },
  "compression_metrics": { "sheets": {…}, "overall": {…} }
}
```

---

## Bugs fixed (all confirmed passing after fix)

| # | File | Error | Root cause | Fix |
|---|------|-------|-----------|-----|
| 1 | Models/*.cs | `CS0518 IsExternalInit` | `init` properties (C# 9) need a polyfill on netstandard2.0 | Added `SpreadsheetLLM.Core/IsExternalInit.cs` |
| 2 | Core.csproj | `NU1903` vulnerability | `System.Text.Json 7.0.0` has high-severity CVE | Upgraded to `8.0.5` |
| 3 | ExcelReader.cs:134,137 | `CS1061 XLColor.IsEmpty` | `XLColor` in ClosedXML 0.102.2 has no `IsEmpty` property | Removed guard; colour `.ToString()` always succeeds |
| 4 | ChainOfSpreadsheet.cs:249 | `CS1503 Split overload` | `Split(char, StringSplitOptions)` only in .NET 5+; not on netstandard2.0 | Changed to `Split(new char[]{ ' ' }, StringSplitOptions.RemoveEmptyEntries)` |

## Test results (11/11 passing)

Run with: `dotnet run --project SpreadsheetLLM.TestRunner -c Release`

| Test | Sheets | Original tokens | Final tokens | Ratio |
|------|--------|-----------------|--------------|-------|
| Simple table | 1 | 310 | 935 | 0.33x |
| Multi-table sheet | 1 | 260 | 725 | 0.36x |
| Merged cells | 1 | 263 | 669 | 0.39x |
| Formulas | 1 | 264 | 684 | 0.39x |
| Dates & currency | 1 | 265 | 648 | 0.41x |
| All-numeric | 1 | 599 | 1001 | 0.60x |
| Mixed formats | 1 | 402 | 1706 | 0.24x |
| **Large sheet (50r×10c)** | 1 | **7675** | **3794** | **2.02x** |
| Sparse | 1 | 121 | 522 | 0.23x |
| Multi-sheet workbook | 3 | 679 | 1968 | 0.35x |
| plain_adjacent.xlsx | 1 | 713 | 1432 | 0.50x |

> **Note on small-sheet ratios < 1**: The JSON metadata overhead (structural anchors, type keys, format groups) dominates for small sheets. Compression becomes beneficial on larger, repetitive workbooks (the paper reports ~25x on real-world spreadsheets with many repeated patterns). The large sheet test already shows 2x even with generated data.

---

## Key classes quick-reference

### `SheetCompressor` (public API)
```csharp
var encoding = new SheetCompressor().Encode("path/to/file.xlsx", k: 2);
// or from pre-loaded snapshots:
var encoding = new SheetCompressor().Encode(snapshots, "filename.xlsx", k: 2);
```

### `ExcelReader` (static)
```csharp
WorksheetSnapshot[] sheets = ExcelReader.ReadWorkbook("path.xlsx");
```

### `CellUtils` (static helpers)
- `InferCellDataType(cell)` → `"empty"|"text"|"numeric"|"boolean"|"datetime"|"email"|"error"|"formula"`
- `DetectSemanticType(cell)` → 9 semantic types from paper
- `GetStyleFingerprint(cell)` → stable int hash for boundary detection
- `CellCoord(row, col)` → `"A1"` (1-based)
- `ColumnLetter(col)` / `ColumnNumber(colLetter)` → conversion helpers
- `SplitCellRef("AB12")` → `("AB", 12)`

### `WorksheetSnapshot`
- `Cells[row-1][col-1]` — zero-based backing array
- `GetCell(row, col)` — **1-based**, bounds-checked, returns `null` for empty

### `ChainOfSpreadsheet` (QA)
Needs env vars `ANTHROPIC_API_KEY` or `OPENAI_API_KEY`.  
Backends: `"anthropic"` (claude-sonnet), `"openai"` (gpt-4), `"placeholder"` (mock).

---

## Style / conventions

- All row/column indices are **1-based** everywhere (matching Excel).  
- Formulas stored as `"=SUM(B2:B5)"` (with leading `=`).  
- `init` properties used on model classes — polyfill required for netstandard2.0.  
- Pre-compiled `static readonly Regex` fields in CellUtils for performance.
