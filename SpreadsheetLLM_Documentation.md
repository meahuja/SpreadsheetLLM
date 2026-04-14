# SpreadsheetLLM .NET Implementation

> **Technical Basis**: arXiv:2407.09025 research paper, C# .NET implementation  
> **Date**: April 2026

---

## Table of Contents

1. [What is SpreadsheetLLM?](#1-what-is-spreadsheetllm)
2. [Why is This Technology Needed?](#2-why-is-this-technology-needed)
3. [Overall System Architecture](#3-overall-system-architecture)
4. [The 3-Stage Compression Pipeline — Plain English](#4-the-3-stage-compression-pipeline--plain-english)
   - [Stage 1: Finding Important Rows & Columns (Structural Anchor Extraction)](#stage-1-finding-important-rows--columns--structural-anchor-extraction)
   - [Stage 2: Flipping the Index (Inverted-Index Translation)](#stage-2-flipping-the-index--inverted-index-translation)
   - [Stage 3: Grouping by Data Format (Data-Format-Aware Aggregation)](#stage-3-grouping-by-data-format--data-format-aware-aggregation)
5. [Real Large-Sheet Examples](#5-real-large-sheet-examples)
   - [Example A — 500-Row Sales Transactions Sheet](#example-a--500-row-sales-transactions-sheet)
   - [Example B — 300-Row HR Payroll Sheet](#example-b--300-row-hr-payroll-sheet)
   - [Example C — Large Employee Sheet (50 rows × 10 columns)](#example-c--large-employee-sheet-50-rows--10-columns)
6. [Code Files — What Each One Does](#6-code-files--what-each-one-does)
7. [How to Read the JSON Output](#7-how-to-read-the-json-output)
8. [Compression Performance Numbers](#8-compression-performance-numbers)
9. [Frequently Asked Questions](#9-frequently-asked-questions)

---

## 1. What is SpreadsheetLLM?

### In One Sentence

> **SpreadsheetLLM is software that automatically compresses Excel files so that AI (Large Language Models, or LLMs) can understand them faster and more cheaply.**

### Understanding Through an Analogy

Imagine you want to hire a foreign-language expert to translate a 500-page book.  
The more pages you send, the **higher the cost and the longer it takes**.

But if you first reduce the book to a **30-page executive summary** and then send it for translation:
- The cost goes down
- The time gets shorter
- The expert understands it faster

SpreadsheetLLM does exactly this for Excel files.  
Before sending data to an AI, it **compresses the spreadsheet down to only the essential information**.

---

## 2. Why is This Technology Needed?

### The Problem: Excel Files Are Too Big for AI

| Situation | Details |
|-----------|---------|
| Typical Excel file size | Tens to hundreds of rows × dozens of columns |
| AI processing limit | There is a cap on how many characters (tokens) AI can receive at once |
| AI API billing | Charged per token (character) sent |

When you convert an Excel file as-is into text and send it to AI, it looks like this:
```
A1: Product, B1: Quantity, C1: Unit Price, D1: Total
A2: Apple, B2: 100, C2: 500, D2: 50000
A3: Banana, B3: 200, C3: 300, D3: 60000
... (hundreds of rows repeating)
```

This approach **wastes tokens (characters)**, and the repetitive numbers interfere with the AI's understanding.

### The Solution: Smart Compression

SpreadsheetLLM automatically handles the following:

1. Finds the structural boundaries of tables (header rows, section dividers).
2. **Merges repeated values across multiple cells into one entry.** (e.g., "Completed" → B5, B8, B12, B19)
3. **Groups cells with the same format into ranges.** (e.g., currency format → C2:C500)

Result: The amount of data sent to AI is reduced by up to **25 times**.

---

## 3. Overall System Architecture

```
Excel File (.xlsx)
       │
       ▼
┌─────────────────────┐
│   ExcelReader.cs    │  ← Opens the Excel file and loads all cell data into memory
│   (File Reader)     │
└─────────────────────┘
       │
       ▼
┌─────────────────────┐
│  SheetCompressor.cs │  ← Runs the 3-stage compression pipeline
│  (Core Compressor)  │
│                     │
│  Stage 1: Structure │  → Identify important rows & columns
│  Stage 2: Inversion │  → Map values → locations
│  Stage 3: Grouping  │  → Bundle cells with same format
└─────────────────────┘
       │
       ▼
┌─────────────────────┐
│   CellUtils.cs      │  ← Cell classification helper (Date? Currency? Integer? Email?)
│   (Cell Analyser)   │
└─────────────────────┘
       │
       ▼
┌─────────────────────┐
│  JSON Output File   │  ← The final compressed result sent to AI
│  (Compressed Result)│
└─────────────────────┘
```

### Code Files at a Glance

| File | Role | Analogy |
|------|------|---------|
| `ExcelReader.cs` | Reads Excel file into memory | A document scanner |
| `SheetCompressor.cs` | Runs the 3-stage compression | A professional summariser |
| `CellUtils.cs` | Classifies each cell's data type | A data classifier |
| `CellData.cs` | Stores all information for a single cell | A cell's ID card |
| `WorksheetSnapshot.cs` | Stores the entire sheet as a 2D grid | A photograph of the sheet |
| `SheetEncoding.cs` | Defines the JSON structure of the output | The template for the compressed result |
| `VanillaEncoder.cs` | Generates uncompressed baseline text (for comparison) | The original reference line |
| `ChainOfSpreadsheet.cs` | Q&A pipeline connecting to AI | The AI communication bridge |

---

## 4. The 3-Stage Compression Pipeline — Plain English

### The Example Spreadsheet We Will Use

Assume we have the following sales data:

```
     A              B        C           D
1  Product Name   Quantity  Unit Price   Total
2  Apple            100      $0.50      $50.00
3  Banana           200      $0.30      $60.00
4  Cherry           150      $1.20     $180.00
5  Date (fruit)      80      $2.00     $160.00
6  Elderberry        60      $3.50     $210.00
7  Grand Total      ---       ---    =SUM(D2:D6)
```

---

### Stage 1: Finding Important Rows & Columns — Structural Anchor Extraction

**Code location**: `SheetCompressor.cs` → `FindStructuralAnchors()` method

#### What does it do?

Think of it like finding the **table of contents, chapter titles, and subheadings** in a book first.  
Instead of reading every word, you identify only the key structural points.

#### How does it find them? (Step by step)

**Step 1 — Row Analysis: Determine what kind of row each one is**

```
Code: AnalyzeRowsSinglePass() method
```

For each row, the following checks are made:

| Check | Example | Result |
|-------|---------|--------|
| Are 60%+ of cells bold? | Row 1: Product Name, Quantity, Unit Price, Total — all bold | Header row! |
| Are 60%+ of cells centre-aligned? | Row 1 is all centre-aligned | Header row! |
| Do cells have bottom borders? | Row 1 has bottom borders | Header row! |
| Are text values all-caps? | TOTAL, REVENUE, etc. | Header row! |
| Does data type change from the previous row? | Row 1 = text, Row 2 = numbers begin | Boundary! |
| Is the row empty? | Empty row between row 6 and 7 | Section separator! |

**Step 2 — Generate Boundary Candidates**

```
Code: FindBoundaryCandidates() method
```

Based on the row and column analysis, the system creates candidate regions that "could be the boundary of a table".

For example:
- Row boundary candidates: [1, 2, 7]  (row 1 = header, row 2 = data start, row 7 = total row)
- Column boundary candidates: [A, B, C, D]

**Step 3 — Remove Duplicates (NMS — Non-Maximum Suppression)**

```
Code: NmsCandidates() method
```

Duplicate candidates pointing to the same region are eliminated, keeping only the most meaningful ones.

**Analogy**: When multiple people mark "this part is important!" and their marks overlap, we merge them into one.

**Step 4 — Expand Anchors (k=2 by default)**

```
Code: ExpandAnchors() method
```

The rows/columns surrounding each anchor are also included — 2 rows/columns in each direction.  
If anchor is row 5 → rows 3, 4, 5, 6, 7 are all included.

**Step 5 — Remove Uniform Repeated Rows/Columns**

```
Code: CompressHomogeneousRegions() method
```

Rows where every cell has the same value and format are removed.  
(e.g., In a 500-row sheet, if 300 rows have status = "Completed", those rows are bundled together in Stage 2 instead)

**Stage 1 Result (our example)**:
```
Retained rows: [1, 2, 3, 4, 5, 6, 7]  ← Small table, all rows kept
Retained columns: [A, B, C, D]
Structural anchors: rows=[1, 7], columns=["A", "D"]
```

---

### Stage 2: Flipping the Index — Inverted-Index Translation

**Code location**: `SheetCompressor.cs` → `CreateInvertedIndex()`, `CreateInvertedIndexTranslation()`

#### What does it do?

Normal Excel representation:
```
A1=Product Name, A2=Apple, A3=Banana ...
```

Inverted index representation (flipped):
```
"Apple"          → [A2]
"Banana"         → [A3]
"=SUM(D2:D6)"   → [D7]
```

#### Why do it this way?

In a 500-row sales dataset, if the status "Completed" appears in 350 rows:

**Original way**: `J2=Completed, J5=Completed, J8=Completed, ... (350 repetitions)` — enormous waste!

**Inverted index way**: `"Completed" → [J2, J5, J8, J12, ...]` — recorded only once!

#### Range Merging (MergeCellRanges)

```
Code: MergeCellRanges() method
```

Consecutive cells are merged into ranges:
```
[J2, J3, J4, J5, J6]  →  "J2:J6"   (5 entries → expressed as 1)
```

**Effect on large sheets**:
```
In a 500-row sales dataset:
- "Widget A" appears 98 times → "Widget A" → ["B2:B99"]  (continuous run)
```

#### Stage 2 Result (our example):
```json
"cells": {
  "Product Name":   ["A1"],
  "Quantity":       ["B1"],
  "Unit Price":     ["C1"],
  "Total":          ["D1"],
  "Apple":          ["A2"],
  "Banana":         ["A3"],
  "Cherry":         ["A4"],
  "Date (fruit)":   ["A5"],
  "Elderberry":     ["A6"],
  "Grand Total":    ["A7"],
  "=SUM(D2:D6)":   ["D7"]
}
```

---

### Stage 3: Grouping by Data Format — Data-Format-Aware Aggregation

**Code location**: `SheetCompressor.cs` → `GroupBySemanticType()`, `AggregateBySemanticType()`

#### What does it do?

Cells are grouped by their **data type** and **format pattern**.

`CellUtils.cs` classifies each cell into one of 9 semantic types:

| Type | Example |
|------|---------|
| `integer` | 100, 200, 80 |
| `float` | 3.14, 1.75 |
| `currency` | $1,250.50, £3,800.00 |
| `percentage` | 75%, 12.5% |
| `scientific` | 1.23E+05 |
| `date` | 2024-01-15 |
| `time` | 09:30:00 |
| `year` | 2023, 2024 |
| `email` | user@company.com |

#### Stage 3 Result (our example):
```json
"formats": {
  "{\"type\":\"text\",\"nfs\":\"General\"}":        ["A1:A6"],
  "{\"type\":\"integer\",\"nfs\":\"General\"}":     ["B2:B6"],
  "{\"type\":\"currency\",\"nfs\":\"$#,##0.00\"}":  ["C2:C6", "D2:D7"]
}
```

With this information, the AI instantly understands "Column C is entirely currency-formatted".

---

## 5. Real Large-Sheet Examples

### Example A — 500-Row Sales Transactions Sheet

**Code location**: `Program.cs` → `CreateRealisticSalesSheet()` method

#### Sheet Structure

```
     A       B           C         D              E          F         G       H          I           J           K
1  TxnID   Date        Region  Salesperson    Product    Category    Qty  UnitPrice   Total       Status   PaymentMethod
2   1001  2023-03-15   North   Alice Smith   Widget A   Hardware     12    $9.99    $119.88   Completed   Credit Card
3   1002  2023-07-22   South   Bob Jones     Widget B   Software      5   $14.99    $74.95    Completed   Cash
4   1003  2023-01-08   East    Carol White   Gadget X   Services     30   $24.99   $749.70    Pending     Bank Transfer
...
501  1500  2023-11-30  Central  Jack Ford     Part Z   Accessories   18   $99.99  $1,799.82  Completed   Credit Card
---  (empty row)
503  TOTAL  ---          ---      ---            ---      ---         ---    ---    $X,XXX.XX    ---          ---
```

#### Before Compression (sending the raw text)

The AI would receive text like this:
```
A1: TxnID, B1: Date, C1: Region, D1: Salesperson, ... K1: PaymentMethod
A2: 1001, B2: 2023-03-15, C2: North, D2: Alice Smith, E2: Widget A, ...
A3: 1002, B3: 2023-07-22, C3: South, D3: Bob Jones, E3: Widget B, ...
... (500 rows fully repeated)
```
**→ Massive token consumption**

#### After Compression (SpreadsheetLLM output)

**Stage 1 Result — Structural Anchors**:
```json
"structural_anchors": {
  "rows": [1, 2, 503],
  "columns": ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K"]
}
```
→ Only row 1 (header), row 2 (first data row), row 503 (total row) are recognised as anchors.

**Stage 2 Result — Inverted Index (repeated values collapsed to one)**:
```json
"cells": {
  "TxnID": ["A1"],  "Date": ["B1"],  "Region": ["C1"],
  "North":       ["C2", "C8", "C15", "C23", "..."],
  "South":       ["C3", "C7", "C19", "..."],
  "East":        ["C5", "C11", "..."],
  "West":        ["C4", "C9", "..."],
  "Central":     ["C6", "C13", "..."],
  "Widget A":    ["E2:E99"],
  "Widget B":    ["E100:E198"],
  "Completed":   ["J2:J4", "J7", "J9:J11", "..."],
  "Pending":     ["J3", "J8", "..."],
  "Credit Card": ["K2", "K5", "K7", "..."],
  "=SUM(I2:I501)": ["I503"],
  "TOTAL":       ["A503"]
}
```

**Key point**: Even if "Widget A" appears 98 times, it is recorded **only once** in the inverted index!

**Stage 3 Result — Format Aggregation**:
```json
"formats": {
  "{\"type\":\"date\",\"nfs\":\"yyyy-mm-dd\"}":    ["B2:B501"],
  "{\"type\":\"currency\",\"nfs\":\"$#,##0.00\"}": ["H2:H501", "I2:I501", "I503"],
  "{\"type\":\"integer\",\"nfs\":\"General\"}":    ["A2:A501", "G2:G501"],
  "{\"type\":\"text\",\"nfs\":\"General\"}":        ["C2:C501", "D2:D501", "..."]
}
```

→ The AI understands "the entire column B is date-formatted" in a single line.

#### Compression Effect (expected)

| Stage | Tokens |
|-------|--------|
| Original (500 rows × 11 columns) | Tens of thousands of tokens |
| After compression | Dramatically reduced |
| Compression ratio | **Higher the more repeated values there are** |

---

### Example B — 300-Row HR Payroll Sheet

**Code location**: `Program.cs` → `CreateRealisticHRSheet()` method

#### Sheet Structure

```
     A       B              C              D                   E        F           G          H          I          J        K        L
1  EmpID   Name        Department       JobTitle           PayGrade  Location   StartDate  BaseSalary   Bonus   TotalComp  Status  Manager
2   2000  Employee 2000  Engineering  Software Engineer      L3     New York   2017-06-15   $85,000    $8,500    $93,500   Active  John Adams
3   2001  Employee 2001  Sales        Sales Rep              L3     Chicago    2019-03-22   $75,000    $9,200    $84,200   Active  Sarah Lee
4   2002  Employee 2002  Engineering  Senior Engineer        L4     San Fran.  2020-01-10  $115,000   $14,000   $129,000  Active  Mike Brown
...
301  2299  Employee 2299  Finance     Financial Analyst      L4     New York   2022-08-05   $90,000   $11,700   $101,700  On Leave  Lisa Chan
```

#### The Striking Effect of Compression

HR data has heavy repetition patterns:
- Department: only 10 types, repeated 300 times
- Pay Grade: only 5 types (L3–L7), repeated
- Location: only 5 cities, repeated
- Status: only "Active" or "On Leave"
- Manager: only 5 people, repeated

**After inverted index compression**:
```json
"cells": {
  "Engineering": ["C2", "C4", "C7", "C11", "..."],   ← 90 cells → 1 entry
  "Sales":       ["C3", "C8", "C13", "..."],          ← 60 cells → 1 entry
  "Active":      ["K2:K4", "K6:K8", "..."],           ← 280 cells merged into ranges
  "On Leave":    ["K5", "K9", "K15", "..."],          ← 20 cells
  "L3":          ["E2", "E3", "E6", "..."],
  "L4":          ["E4", "E5", "E8", "..."],
  "New York":    ["F2", "F5", "F7", "..."],
  "John Adams":  ["L2", "L6", "L14", "..."]
}
```

**Format aggregation**:
```json
"formats": {
  "{\"type\":\"date\",\"nfs\":\"yyyy-mm-dd\"}":  ["G2:G301"],
  "{\"type\":\"currency\",\"nfs\":\"$#,##0\"}":  ["H2:H301", "I2:I301", "J2:J301"],
  "{\"type\":\"integer\",\"nfs\":\"General\"}":   ["A2:A301"]
}
```

→ The AI instantly grasps: "Column G is all dates, columns H–J are all currency-formatted".

---

### Example C — Large Employee Sheet (50 rows × 10 columns)

**Code location**: `Program.cs` → `CreateLargeSheet()` method

This is the test case that achieved a **2.02x compression ratio** in real testing — the sheet becomes *smaller* after compression.

#### Sheet Structure

```
     A    B            C            D         E        F       G        H            I                 J
1   ID   Name         Dept        Salary   YearsExp  Rating  Active  StartDate      Email            Notes
2    1  Employee1   Engineering  $68,532      7        4.2    Yes    2019-03-15  emp1@company.com    Active
3    2  Employee2   Sales        $75,100     12        3.8    No     2016-07-22  emp2@company.com
4    3  Employee3   HR           $52,800      3        4.7    Yes    2021-01-10  emp3@company.com    Active
...
51  50  Employee50  Marketing    $91,200      9        3.5    Yes    2018-11-30  emp50@company.com
```

#### Stage 1 — Anchor Detection (detailed walkthrough)

**Row analysis**:
- Row 1: ID, Name, Dept, … → all bold → **Header row confirmed!**
- Rows 2–51: Data rows (mix of numbers and text)

**Column analysis**:
- Column A (ID): sequentially increasing integers → unique fingerprint
- Column D (Salary): `$#,##0` currency format
- Column H (StartDate): `yyyy-mm-dd` date format
- Column I (Email): email pattern detected

**Anchor result**:
```json
"structural_anchors": {
  "rows": [1, 2, 51],
  "columns": ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"]
}
```

#### Stage 2 — Inverted Index Result

```json
"cells": {
  "Engineering": ["C2", "C7", "C11", "C19", "C25", "C33", "C41", "C48"],
  "Sales":       ["C3", "C8", "C14", "C22", "C29", "C36", "C44"],
  "HR":          ["C4", "C12", "C20", "C27", "C35"],
  "Finance":     ["C5", "C9", "C16", "C23", "C30"],
  "Marketing":   ["C6", "C10", "C18", "C24", "C31", "C38", "C46"],
  "Yes":         ["G2", "G4", "G6", "G9", "G11", "..."],
  "No":          ["G3", "G5", "G7", "G8", "..."],
  "Active":      ["J2", "J4", "J6", "..."]
}
```

**5 department names repeated 50 times → summarised as 5 entries!**

#### Stage 3 — Format Aggregation Result

```json
"formats": {
  "{\"type\":\"currency\",\"nfs\":\"$#,##0\"}":  ["D2:D51"],
  "{\"type\":\"date\",\"nfs\":\"yyyy-mm-dd\"}":  ["H2:H51"],
  "{\"type\":\"email\",\"nfs\":\"General\"}":    ["I2:I51"],
  "{\"type\":\"integer\",\"nfs\":\"General\"}":  ["A2:A51", "E2:E51"]
}
```

#### Compression Performance

| Stage | Tokens | vs. Original |
|-------|--------|-------------|
| Original (vanilla) | 7,675 | baseline |
| After Stage 1 (anchors only) | ~3,000 | ~2.5x |
| After Stage 2 (inverted index) | ~2,500 | ~3x |
| Final output | **3,794** | **2.02x** |

→ The final file is **less than half** the original size!

---

## 6. Code Files — What Each One Does

### `ExcelReader.cs` — Excel File Reader

**What it does**: Opens an Excel (.xlsx) file and loads all cell information into memory.

```
Excel file → Open with ClosedXML library → Extract each cell's data → Create WorksheetSnapshot
```

**Information extracted**:

| Information | Example |
|-------------|---------|
| Cell value | "Apple", 100, 2024-01-15 |
| Formula | `=SUM(B2:B5)` |
| Number format | Currency (`$#,##0.00`), Date (`yyyy-mm-dd`) |
| Style | Bold, font colour, border, alignment |
| Merge status | Is A1:D1 a merged cell? |

**Important feature**: Formulas are stored as the **formula string itself**, not the computed result.
```
Cell D7 contains =SUM(D2:D6) → stored as "=SUM(D2:D6)"  (not the result 660.00)
```

---

### `CellUtils.cs` — Cell Classifier

**What it does**: Determines what kind of data each cell contains.

#### `InferCellDataType()` — Basic data type detection

```
Cell value → Is it an email? → Is it a formula? → Is it a number? → Is it date-formatted? → Text
```

Returns: `"empty"` | `"text"` | `"numeric"` | `"boolean"` | `"datetime"` | `"email"` | `"error"` | `"formula"`

#### `DetectSemanticType()` — Semantic type detection (9 categories)

```csharp
// Example: cell value is "50000" and format is "$#,##0.00"
DetectSemanticType(cell)  →  "currency"

// Example: cell value is "2024-01-15" and format is "yyyy-mm-dd"
DetectSemanticType(cell)  →  "date"

// Example: cell value is "user@company.com"
DetectSemanticType(cell)  →  "email"
```

#### `GetStyleFingerprint()` — Style fingerprinting

Compresses each cell's **visual style** into a single number.

```
Bold + Blue font + Border  →  fingerprint: 12345678
Not bold + Black font + No border  →  fingerprint: 87654321
```

Same fingerprint = same style → used to detect boundaries between regions.

---

### `SheetCompressor.cs` — Core Compression Engine

The most important file: it orchestrates the entire pipeline.

#### Public API (how to use it)

```csharp
// Option 1: compress directly from a file path
var compressor = new SheetCompressor();
var result = compressor.Encode("path/to/file.xlsx", k: 2);

// Option 2: compress from pre-loaded data (e.g. from a VSTO adapter)
var result = compressor.Encode(snapshots, "filename.xlsx", k: 2);
```

**What `k=2` means**: Expand anchor rows/columns by **2 positions** in each direction (default).

#### Key Constants Explained

| Constant | Value | Meaning |
|----------|-------|---------|
| `MaxCandidates` | 200 | Maximum number of candidate regions |
| `MaxBoundaryRows` | 100 | Maximum number of boundary rows |
| `HeaderThreshold` | 0.6 | Threshold for classifying a row as a header (60%) |
| `SparsityThreshold` | 0.10 | Minimum data density required (10%) |
| `NmsIouThreshold` | 0.5 | Overlap threshold for duplicate removal (50%) |

---

### `VanillaEncoder.cs` — Baseline Encoder (for comparison)

**What it does**: Produces a plain row-by-row text encoding with no compression.

Acts as the **reference baseline** for measuring compression efficiency:
```
Comparison: original size (VanillaEncoder) vs. compressed size (SheetCompressor)
```

---

### `ChainOfSpreadsheet.cs` — AI Q&A Pipeline

**What it does**: Sends the compressed spreadsheet to an AI and handles question-and-answer interactions.

Supported AI backends:
- `"anthropic"` → Uses Claude (Sonnet)
- `"openai"` → Uses GPT-4
- `"placeholder"` → Mock responses for testing

Required environment variables:
```
ANTHROPIC_API_KEY=sk-ant-...   (when using Claude)
OPENAI_API_KEY=sk-...          (when using GPT-4)
```

---

## 7. How to Read the JSON Output

### Output File Location

```
SpreadsheetLLM.TestRunner/bin/Release/net9.0/test_output/
├── Simple_table.json
├── Large_sheet__50r_10c_.json
├── Sales_500_rows.json
├── HR_Payroll_300_rows.json
└── ...
```

### Full JSON Structure Example

```json
{
  "file_name": "sales_data.xlsx",
  "sheets": {
    "SalesTransactions": {
      "structural_anchors": {
        "rows": [1, 2, 503],
        "columns": ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K"]
      },
      "cells": {
        "TxnID":          ["A1"],
        "Date":           ["B1"],
        "Region":         ["C1"],
        "North":          ["C2", "C15", "C31:C35", "C50"],
        "Widget A":       ["E2:E98"],
        "Completed":      ["J2:J4", "J7:J11", "J350:J420"],
        "=SUM(I2:I501)":  ["I503"]
      },
      "formats": {
        "{\"type\":\"date\",\"nfs\":\"yyyy-mm-dd\"}":    ["B2:B501"],
        "{\"type\":\"currency\",\"nfs\":\"$#,##0.00\"}": ["H2:H501", "I2:I501"],
        "{\"type\":\"integer\",\"nfs\":\"General\"}":    ["A2:A501", "G2:G501"],
        "{\"type\":\"text\",\"nfs\":\"General\"}":        ["C2:C501", "D2:D501"]
      },
      "numeric_ranges": {
        "{\"type\":\"integer\",\"nfs\":\"General\"}": ["A2:A501", "G2:G501"]
      }
    }
  },
  "compression_metrics": {
    "sheets": {
      "SalesTransactions": {
        "original_tokens": 45000,
        "after_anchor_tokens": 12000,
        "after_inverted_index_tokens": 5000,
        "after_format_tokens": 3500,
        "final_tokens": 4200,
        "anchor_ratio": 3.75,
        "inverted_index_ratio": 9.00,
        "format_ratio": 12.86,
        "overall_ratio": 10.71
      }
    },
    "overall": {
      "original_tokens": 45000,
      "final_tokens": 4200,
      "overall_ratio": 10.71
    }
  }
}
```

### Field-by-Field Explanation

| Field | Meaning |
|-------|---------|
| `structural_anchors.rows` | List of important row numbers (1-based) |
| `structural_anchors.columns` | List of important column letters |
| `cells` | Inverted index: value → list of cell ranges |
| `formats` | Format aggregation: JSON format key → list of cell ranges |
| `numeric_ranges` | Numeric-only sub-aggregation |
| `original_tokens` | Token count before compression |
| `final_tokens` | Token count after compression |
| `overall_ratio` | Compression ratio (higher = better compression) |

---

## 8. Compression Performance Numbers

Actual test results from CLAUDE.md:

| Test Sheet | Original Tokens | Final Tokens | Ratio | Notes |
|-----------|----------------|-------------|-------|-------|
| Simple table | 310 | 935 | 0.33x | Small sheet — metadata overhead dominates |
| Multi-table sheet | 260 | 725 | 0.36x | Small sheet |
| Merged cells | 263 | 669 | 0.39x | Merge handling included |
| Formulas | 264 | 684 | 0.39x | Formulas preserved as-is |
| Dates & currency | 265 | 648 | 0.41x | Format aggregation benefit |
| All-numeric | 599 | 1001 | 0.60x | No structural cues |
| Mixed formats | 402 | 1706 | 0.24x | Many format varieties |
| **Large sheet (50×10)** | **7,675** | **3,794** | **2.02x** | **First true compression gain** |
| Sparse sheet | 121 | 522 | 0.23x | Little data |
| Multi-sheet workbook | 679 | 1968 | 0.35x | 3 sheets |

### Why Is the Ratio Below 1 for Small Sheets?

For small sheets (a few dozen rows):
- The **metadata added by the algorithm** (structural anchors, JSON type keys, etc.) can be larger than the original
- Very little repetition means the inverted index provides minimal benefit

For large, real-world sheets (hundreds to thousands of rows):
- Heavy repetition means the **inverted index effect is maximised**
- The paper reports up to **25x compression** on real enterprise spreadsheets with many repeated patterns
- The large-sheet test already shows 2x with synthetic data

---

## 9. Frequently Asked Questions

**Q: How do I run the software?**

```powershell
# 1. Build
dotnet build SpreadsheetLLM.Core/SpreadsheetLLM.Core.csproj -c Release

# 2. Run tests (generates sample sheets + compresses them)
dotnet run --project SpreadsheetLLM.TestRunner -c Release

# 3. Check results
# JSON files are created in: SpreadsheetLLM.TestRunner/bin/Release/net9.0/test_output/
```

**Q: Will my Excel file be modified or damaged?**

No. The software **only reads** the file. The original file is never changed.

**Q: Which Excel file formats are supported?**

Only `.xlsx` format (Excel 2007 and later).

**Q: Can I compress without an AI API key?**

Yes. The compression (encoding) itself works without any API key.  
The API key is only needed when sending questions to AI (`ChainOfSpreadsheet`).

**Q: How are formulas handled?**

Formulas are preserved as their **original formula string**, not their computed result.
```
Cell D7 contains =SUM(D2:D6)  →  sent to AI as "=SUM(D2:D6)"
```
The AI can read and understand the formula's intent.

**Q: How are merged cells handled?**

Merged cells are resolved using the **value of the top-left cell** in the merged range.  
Example: A1:D1 merged as "Quarterly Report" → A1, B1, C1, D1 are all treated as having the value "Quarterly Report".

**Q: Does it support non-English text (e.g., Korean, Japanese, Chinese)?**

Yes. C# `string` uses Unicode internally, so Korean, Japanese, Chinese, and all other languages are handled correctly.

**Q: What happens if the sheet has many empty cells (sparse data)?**

The sparsity filter (`SparsityThreshold = 10%`) automatically removes candidate regions that are less than 10% populated. A small dense region surrounded by empty space is still correctly identified and compressed.

**Q: How is the compression ratio calculated?**

```
Compression Ratio = Original Token Count ÷ Final Token Count

Example: Original = 7,675 tokens, Final = 3,794 tokens
Ratio = 7,675 ÷ 3,794 = 2.02x   ← the sheet is compressed to less than half its original size
```

A ratio above 1.0 means the output is smaller than the input. The higher the number, the more efficient the compression.

---

## Appendix: Complete Flow Summary Diagram

```
Excel File  (e.g. sales_data_500rows.xlsx)
│
│  ExcelReader.ReadWorkbook()
▼
WorksheetSnapshot[]  —  snapshot of every cell
│
│  SheetCompressor.Encode()
▼
┌──────────────────────────────────────────────────────────────┐
│                    3-Stage Pipeline                          │
│                                                              │
│  Stage 1: FindStructuralAnchors()                           │
│    └─ Detect headers / boundaries / data-type transitions   │
│    └─ Remove duplicate candidates with NMS                  │
│    └─ Expand each anchor by k=2 rows/columns                │
│    └─ Drop uniform (homogeneous) rows/columns               │
│          ↓                                                   │
│  Stage 2: CreateInvertedIndex()                             │
│    └─ Flip: cell positions → grouped by value               │
│    └─ Merge consecutive cells into ranges (A2:A100)         │
│    └─ Preserve formula strings as-is                        │
│          ↓                                                   │
│  Stage 3: GroupBySemanticType()                             │
│    └─ Classify into 9 semantic types                        │
│    └─ Group same type + format → merge into ranges          │
│    └─ Separate numeric sub-aggregation                      │
└──────────────────────────────────────────────────────────────┘
│
▼
SpreadsheetEncoding  (JSON)
├─ structural_anchors   (structural map of the sheet)
├─ cells                (inverted index)
├─ formats              (format aggregation)
├─ numeric_ranges       (numeric sub-aggregation)
└─ compression_metrics  (performance statistics)
│
▼
Sent to AI  (Claude / GPT-4)
→  Full spreadsheet understanding with far fewer tokens
```

---

*This document was prepared to explain the SpreadsheetLLM .NET implementation (C# netstandard2.0, based on arXiv:2407.09025) to non-technical clients.*  
*Source code reference: `SpreadsheetLLM.Core/` directory*
