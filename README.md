# SpreadsheetLLM

A Python + C# implementation of the **SpreadsheetLLM** framework from the paper:

> *SpreadsheetLLM: Encoding Spreadsheets for Large Language Models*
> arXiv:2407.09025 — Yuzhang Tian et al., Microsoft Research 2024

## What is SpreadsheetLLM?

Spreadsheets have complex 2D layouts, flexible formatting, and large cell counts that overwhelm LLM context windows. The **SheetCompressor** framework achieves ~25x token compression while preserving structural understanding, reaching **78.9% F1** on table detection (12.3% above SOTA).

---

## Project Structure

```
SpreadsheetLLm/
├── spreadsheet_llm/
│   ├── encoder.py          # SheetCompressor 3-stage pipeline
│   ├── cell_utils.py       # Cell type/format/semantic classification
│   ├── cos.py              # Chain-of-Spreadsheet QA pipeline
│   └── vanilla.py          # Vanilla baseline encoding
├── SpreadsheetLLM.Core/    # C# .NET port (for WPF/VSTO integration)
│   ├── SheetCompressor.cs
│   ├── CellUtils.cs
│   ├── ExcelReader.cs
│   ├── ChainOfSpreadsheet.cs
│   └── Models/
├── tests/
│   ├── test_cell_utils.py
│   ├── test_encoder.py
│   └── test_integration.py
├── cli.py                  # CLI entry point
├── demo.py                 # Demo script
└── requirements.txt
```

---

## Python Setup

```bash
pip install -r requirements.txt
```

**Requirements**: Python 3.9+, `openpyxl>=3.1.0`

---

## Usage

### CLI

```bash
# Encode a spreadsheet (SheetCompressor pipeline)
python cli.py encode input.xlsx

# Save output to JSON
python cli.py encode input.xlsx -o output.json -k 2

# Vanilla baseline encoding
python cli.py encode input.xlsx --vanilla

# Chain-of-Spreadsheet QA
python cli.py qa input.xlsx "What is the total revenue?"
python cli.py qa input.xlsx "Which student scored highest?" --provider anthropic
```

### Python API

```python
from spreadsheet_llm.encoder import encode_spreadsheet

result = encode_spreadsheet("my_file.xlsx", k=2)
print(result["sheets"]["Sheet1"]["cells"])
print(result["compression_metrics"]["overall"]["overall_ratio"])
```

### Chain-of-Spreadsheet QA

```python
from spreadsheet_llm.encoder import encode_spreadsheet
from spreadsheet_llm.cos import identify_table, generate_response

encoding = encode_spreadsheet("data.xlsx")
table_range = identify_table(encoding, "What is total revenue?", provider="anthropic")
answer = generate_response(encoding["sheets"]["Sheet1"], "What is total revenue?")
```

Set `ANTHROPIC_API_KEY` or `OPENAI_API_KEY` environment variable for real LLM inference. Without either, a placeholder response is returned.

---

## The 3-Stage SheetCompressor Pipeline

### Stage 1 — Structural-Anchor-Based Extraction (Section 3.2)
Identifies table boundaries via row/column fingerprint diffing, header detection, and data-type transitions. Uses IoU-based NMS to select surviving table candidates. Expands anchors by k-neighborhood (default k=2).

### Stage 2 — Inverted-Index Translation (Section 3.3)
Maps cell values → cell ranges (lossless). Repeated values in adjacent cells are merged into ranges like `A1:A5`. Formulas are preserved as strings (e.g. `=SUM(B2:B5)`).

### Stage 3 — Data-Format-Aware Aggregation (Section 3.4)
Groups cells by semantic type (integer, float, currency, date, etc.) and number format string into contiguous rectangular regions.

**Compression ratios**: Stage 1 ~2-4x, Stage 2 ~5-15x, Stage 3 up to ~25x overall.

---

## C# .NET Port

Located in `SpreadsheetLLM.Core/` — a `netstandard2.0` class library for use in WPF or VSTO Excel add-ins.

**NuGet dependencies**: `ClosedXML`, `System.Text.Json`

```csharp
var compressor = new SheetCompressor();
var result = compressor.Encode("C:\\path\\to\\file.xlsx", k: 2);
var json = JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
```

---

## JSON Output Schema

```json
{
  "file_name": "myfile.xlsx",
  "sheets": {
    "Sheet1": {
      "structural_anchors": { "rows": [1, 2, 6, 7, 12], "columns": ["A","B","C","D"] },
      "cells": {
        "Fruit": ["A1"],
        "=B2*C2": ["D2"],
        "=SUM(D2:D5)": ["D6"]
      },
      "formats": {
        "{\"nfs\":\"General\",\"type\":\"text\"}": ["A1:D1", "A7:D7"],
        "{\"nfs\":\"General\",\"type\":\"integer\"}": ["C2:C5", "B8:C11"]
      },
      "numeric_ranges": {
        "{\"nfs\":\"General\",\"type\":\"integer\"}": ["C2:C5", "B8:C11"]
      }
    }
  },
  "compression_metrics": {
    "overall": {
      "original_tokens": 808,
      "final_tokens": 420,
      "overall_ratio": 1.92
    }
  }
}
```

---

## Running Tests

```bash
python -m pytest tests/ -v
```

---

## LLM Provider Setup

| Provider | Environment Variable | Model Used |
|----------|---------------------|------------|
| Anthropic | `ANTHROPIC_API_KEY` | claude-sonnet-4-20250514 |
| OpenAI | `OPENAI_API_KEY` | gpt-4 |
| Placeholder | _(none)_ | Demo responses |

---

## Reference

- **Paper**: [arXiv:2407.09025](https://arxiv.org/abs/2407.09025)
- **Reference repo**: [kingkillery/Spreadsheet_LLM_Encoder](https://github.com/kingkillery/Spreadsheet_LLM_Encoder)
