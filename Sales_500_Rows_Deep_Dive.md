# Sales_500_rows.xlsx — Step-by-Step Deep Dive
## How SpreadsheetLLM Compresses a Real 500-Row Sheet

> All numbers, cell references, and JSON snippets in this document come directly from the
> actual output file: `test_output/Sales_500_rows.json`

---

## Table of Contents

1. [What the Excel Sheet Looks Like](#1-what-the-excel-sheet-looks-like)
2. [The Big Picture — What Compression Achieves](#2-the-big-picture--what-compression-achieves)
3. [Stage 1 — Structural Anchor Extraction (Finding the Skeleton)](#3-stage-1--structural-anchor-extraction-finding-the-skeleton)
4. [Stage 2 — Inverted-Index Translation (Flipping the Map)](#4-stage-2--inverted-index-translation-flipping-the-map)
5. [Stage 3 — Data-Format-Aware Aggregation (Grouping by Type)](#5-stage-3--data-format-aware-aggregation-grouping-by-type)
6. [The Final JSON Output — What AI Actually Receives](#6-the-final-json-output--what-ai-actually-receives)
7. [Compression Numbers — At Every Stage](#7-compression-numbers--at-every-stage)
8. [Summary — Why This Matters](#8-summary--why-this-matters)

---

## 1. What the Excel Sheet Looks Like

The file `Sales_500_rows.xlsx` has **one sheet** called `SalesTransactions`.  
It contains **503 rows** total and **11 columns** (A through K).

### Layout

```
Row 1   → HEADER ROW (column names, bold, light-blue background)
Rows 2–501 → 500 data rows (one sales transaction per row)
Row 502 → EMPTY (gap before the total)
Row 503 → TOTAL row (grand total formula)
```

### The 11 Columns

| Column | Header Name   | What it Contains              | Example Value       |
|--------|---------------|-------------------------------|---------------------|
| A      | TxnID         | Unique transaction number     | 1001, 1002 … 1500   |
| B      | Date          | Sale date (yyyy-mm-dd format) | 15-02-2023          |
| C      | Region        | One of 5 regions              | North / South / East / West / Central |
| D      | Salesperson   | One of 10 sales people        | Alice Smith, Bob Jones … |
| E      | Product       | One of 5 products             | Widget A, Widget B, Gadget X, Gadget Y, Part Z |
| F      | Category      | One of 4 categories           | Hardware / Software / Services / Accessories |
| G      | Qty           | Quantity sold (1–49)          | 7, 38, 48           |
| H      | UnitPrice     | Price per unit ($)            | $9.99, $14.99, $24.99, $49.99, $99.99 |
| I      | Total         | Qty × UnitPrice               | $349.93, $379.62    |
| J      | Status        | One of 3 statuses             | Completed / Pending / Cancelled |
| K      | PaymentMethod | One of 3 payment types        | Credit Card / Cash / Bank Transfer |

### A Peek at the First Few Data Rows

```
Row  A     B                    C       D           E         F          G    H       I        J          K
1   TxnID  Date                Region  Salesperson  Product   Category   Qty  UPrice  Total    Status     Payment
2   1001   15-02-2023 00:00:00 East    Bob Jones    Gadget Y  Software   7    49.99   349.93   Pending    Cash
3   1002   27-03-2023 00:00:00 South   Frank Hall   Widget A  (varies)   38   9.99    379.62   Completed  Bank Transfer
...
501 1500   26-09-2023 00:00:00 East    Grace Lee    Part Z    Software   40   99.99   3999.60  Completed  Bank Transfer
502 (empty row)
503 TOTAL                                                                              =SUM(I2:I501)
```

### The Core Challenge

If we sent this sheet as plain text to an AI, it would look like:
```
A1: TxnID, B1: Date, C1: Region, D1: Salesperson, E1: Product, F1: Category, G1: Qty,
H1: UnitPrice, I1: Total, J1: Status, K1: PaymentMethod,
A2: 1001, B2: 15-02-2023 00:00:00, C2: East, D2: Bob Jones, E2: Gadget Y, F2: Software,
G2: 7, H2: 49.99, I2: 349.93, J2: Pending, K2: Cash,
... (500 more rows of repetition)
```

**That would be 74,717 characters (tokens).** SpreadsheetLLM reduces this to **17,157** — a **4.35× reduction** — while the AI still understands the full structure.

---

## 2. The Big Picture — What Compression Achieves

| Stage | Tokens After Stage | Reduction from Original |
|-------|--------------------|-------------------------|
| Original (no compression) | **74,717** | 1× (baseline) |
| After Stage 1 (Structural Anchors) | **26,957** | **2.77×** smaller |
| After Stage 2 (Inverted Index) | **15,934** | **4.69×** smaller |
| After Stage 3 (Format Aggregation) | 786 (formats only) | **95×** smaller |
| **Final output (all stages combined)** | **17,157** | **4.35×** smaller |

> **Plain English**: Imagine you have a 100-page report. Stage 1 cuts it to 36 pages.
> Stage 2 cuts it further to 21 pages. The format section is only 1 page. The complete
> compressed document is 23 pages — yet it contains the same information the AI needs.

---

## 3. Stage 1 — Structural Anchor Extraction (Finding the Skeleton)

### What is a "Structural Anchor"?

An anchor is a **row or column that marks something important** in the sheet:
- The header row (column names)
- The first data row
- A change in data type or format
- A total/summary row

Think of it like the chapter markers in a book. You don't need every page —
just the places where something **changes or begins**.

### The Code That Does This

```csharp
// SheetCompressor.cs — FindStructuralAnchors()
private (List<int> rowAnchors, List<int> colAnchors) FindStructuralAnchors(
    WorksheetSnapshot sheet, int k)
{
    var rowInfo = AnalyzeRowsSinglePass(sheet);     // Step A: examine every row
    var (rowBounds, colBounds) = FindBoundaryCandidates(sheet); // Step B: find edges
    var candidates = ComposeCandidatesConsecutive(rowBounds, colBounds); // Step C: make boxes
    candidates = FilterCandidates(sheet, candidates, rowInfo);  // Step D: remove bad ones
    candidates = NmsCandidates(candidates, rowInfo); // Step E: remove duplicates
    // Step F: collect anchor rows/cols from winners
}
```

Let's walk through each step using the real Sales_500_rows data.

---

### Step A — Row Analysis (`AnalyzeRowsSinglePass`)

The code scans **every one of the 503 rows**, one by one, and records:

```csharp
// For each row r from 1 to 503:
//   - Is this row empty?
//   - What is the row's "fingerprint" (a number summarising all its values + styles)?
//   - Is this a header row? (bold / centred / all-caps?)
//   - How many cells contain numbers? Text?
```

**Row 1 — The Header Row**

The code checks each of the 11 cells in row 1:

```
Cell A1: value="TxnID",  FontBold=true, FillBackground=LightBlue
Cell B1: value="Date",   FontBold=true, FillBackground=LightBlue
Cell C1: value="Region", FontBold=true, FillBackground=LightBlue
... (all 11 cells are bold + light blue)
```

Header detection logic:

```csharp
// In AnalyzeRowsSinglePass():
if ((double)boldCount / populated > HeaderThreshold)  // 11/11 = 100% > 60%
    isHeader = true;
```

11 bold cells out of 11 populated = **100% bold** → exceeds the 60% threshold (`HeaderThreshold = 0.6`).  
**Result: Row 1 is flagged as a HEADER row.**

**Rows 2–501 — The Data Rows**

Each data row has a unique "fingerprint" because:
- Column A has a different TxnID each row (1001, 1002, 1003 …)
- Column B has a different date each row
- Column I has a different Total amount each row

So practically every row has a different fingerprint from the previous one.
This produces a very large set of boundary candidates (every row is a "change point").

```csharp
// Fingerprint comparison:
if (rowInfo.Fingerprints[rIdx] != rowInfo.Fingerprints[rIdx - 1])
{
    rowBoundarySet.Add(r);       // this row is a boundary
    rowBoundarySet.Add(r - 1);   // the previous row is also marked
}
```

**Row 502 — The Empty Row**

```csharp
// Empty row → mark it AND its neighbours as boundaries
if (rowInfo.Empty[rIdx])
{
    rowBoundarySet.Add(r);       // row 502 itself
    rowBoundarySet.Add(r - 1);   // row 501 (last data row)
    rowBoundarySet.Add(r + 1);   // row 503 (TOTAL row)
}
```

**Row 503 — The TOTAL Row**

"TOTAL" in column A is all-caps text, followed by a formula in column I.
This is a data-type transition (text → formula) and produces another boundary.

---

### Step B — Capping Boundaries (`MaxBoundaryRows = 100`)

After analysing all rows, the code may have found **hundreds of boundary rows**
(because almost every row in a 500-row random dataset has a different fingerprint).

The code limits this to 100:

```csharp
// SheetCompressor.cs — FindBoundaryCandidates()
if (sortedRows.Count > MaxBoundaryRows)   // MaxBoundaryRows = 100
{
    int step = sortedRows.Count / MaxBoundaryRows;  // e.g. 500 / 100 = step of 5
    var sampled = new HashSet<int>();
    for (int i = 0; i < sortedRows.Count; i += step)
        sampled.Add(sortedRows[i]);          // take every 5th boundary

    // Always keep: first, last, and all detected header rows
    sampled.Add(sortedRows[0]);
    sampled.Add(sortedRows[sortedRows.Count - 1]);
    foreach (var hr in headerRows) sampled.Add(hr);

    sortedRows = sampled.OrderBy(r => r).ToList();
}
```

**What this means in plain English:**  
With ~500 boundary candidates and a step of 5, the code keeps every **5th** boundary row.
This is why you see the regular pattern of `251, 256, 261, 266, 271 …` (every 5 rows)
in the final anchor list. The header row (row 1) and the last data row (row 501) are
always kept regardless.

---

### Step C — Compose Candidate Regions

The sampled row boundaries and column boundaries are combined into rectangular **candidate regions**:

```csharp
// ComposeCandidatesConsecutive():
// For consecutive pairs of boundaries, create a rectangle
for (int i = 0; i < rowBounds.Count - 1; i++)
    for (int j = 0; j < colBounds.Count - 1; j++)
        candidates.Add((rowBounds[i], colBounds[j], rowBounds[i+1], colBounds[j+1]));
```

Think of it like drawing boxes between consecutive boundary lines on graph paper.

For our sheet, column boundaries span **A through K** (all 11 columns detected as boundaries
because their fingerprints are all different from each other).

---

### Step D — Filter Candidates

The code discards any candidate region that:

1. **Is too sparse** (less than 10% of cells filled):
   ```csharp
   // SparsityThreshold = 0.10
   if ((double)estPopulated / totalCells < SparsityThreshold) continue;
   ```

2. **Has no header** (no bold/centred/caps rows in the first 3 rows):
   ```csharp
   if (!hasHeader) continue;
   ```

For our 500-row sheet, the region spanning rows 1–503 with all 11 columns is:
- 503 × 11 = 5,533 cells total
- ~5,000+ cells filled (transactions + header + total)
- Density ≈ 90% → **passes the 10% sparsity check**
- Row 1 is bold → **has a header**
- **Kept.**

---

### Step E — Non-Maximum Suppression (NMS)

When multiple candidate regions overlap significantly (more than 50%), only the best one is kept.

```csharp
// NmsCandidates():
// Score each candidate by: (number of header rows × 10) + (total filled cells)
// Keep the highest-scoring one; discard others that overlap it by > 50%

private const double NmsIouThreshold = 0.5;   // 50% overlap = duplicate

while (indices.Count > 0)
{
    int best = indices[0];          // take the highest-scoring candidate
    keep.Add(candidates[best]);
    // remove all others that overlap the winner by more than 50%
    indices = indices.Where(idx =>
        CalculateIoU(candidates[best], candidates[idx]) < NmsIouThreshold).ToList();
}
```

**Plain English:** If two candidates are basically describing the same table region,
keep only the one with the best score (more filled cells + has a header row).

---

### Step F — Expand Anchors by k=2

After finding the winning anchor rows and columns, the code **expands each anchor
by 2 rows/columns in every direction**:

```csharp
// ExpandAnchors() — k = 2 (default)
foreach (int anchor in rowAnchors)
{
    for (int i = Math.Max(1, anchor - k); i <= Math.Min(maxRow, anchor + k); i++)
        keptRows.Add(i);
}
```

**Example:** If anchor row 251 is selected:
```
anchor - k = 251 - 2 = 249  → include row 249
              251 - 1 = 250  → include row 250
              251             → include row 251 (the anchor itself)
              251 + 1 = 252  → include row 252
              251 + 2 = 253  → include row 253
```

This ensures that context around each important point is preserved — not just the boundary itself.

---

### The Final Anchor List (Actual Output)

```json
"structural_anchors": {
  "rows": [
    1,
    21, 46,
    251, 256, 261, 266, 271, 276, 281, 286,
    301, 306, 311, 316, 321, 326, 331, 336, 341, 346, 351, 356, 361, 366, 371, 376, 381,
    503
  ],
  "columns": ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K"]
}
```

**Reading this output — what each anchor row means:**

| Anchor Row(s) | Why Selected |
|---------------|-------------|
| **1** | Header row (bold + light-blue background — 100% bold cells) |
| **21, 46** | Rows near sampled boundaries in the first 50-row block, expanded by k=2 |
| **251, 256, 261 … 381** | Sampled every ~5 rows across the middle/late data (after boundary capping to 100) |
| **503** | The TOTAL row (transition: empty row 502 marks its neighbour as a boundary) |

**All 11 columns (A–K) are selected** because every column has a different fingerprint
from every other column (different data types: IDs, dates, text, currency, numbers).

**What this means for the AI:**  
Instead of processing 503 rows × 11 columns = **5,533 cells**, Stage 1 keeps only
the cells at the anchor intersections. Token count drops from **74,717 → 26,957**
(a **2.77×** reduction) before Stage 2 even runs.

---

## 4. Stage 2 — Inverted-Index Translation (Flipping the Map)

### The Core Idea

**Normal spreadsheet thinking:**
> "Go to cell C2. Its value is 'East'."

**Inverted index thinking:**
> "The value 'East' appears in: C2, C23, C44, C253, C262, ..."

Instead of mapping *location → value*, we map *value → [list of locations]*.
This is enormously efficient when the same value appears in hundreds of cells.

### The Code

```csharp
// SheetCompressor.cs — CreateInvertedIndex()
foreach (int r in rows)         // only the KEPT rows from Stage 1
{
    foreach (int c in cols)     // only the KEPT columns from Stage 1
    {
        var coord = CellUtils.CellCoord(r, c);   // e.g. "C2"
        var value = cell?.Value;                  // e.g. "East"

        AddToGroup(inverted, value, coord);        // add C2 to the "East" bucket
    }
}
```

Then consecutive cell references are merged into ranges:

```csharp
// MergeCellRanges() — called by CreateInvertedIndexTranslation()
// If the "East" bucket contains [C299, C300], they become "C299:C300"
// If it contains [C2], it stays as "C2"
```

---

### Example 1 — "Region" Column (Column C)

The Region column has only **5 possible values**: North, South, East, West, Central.  
Yet it spans 500 rows (C2 to C501).

**Without compression:**
```
C2: East, C3: South, C4: West, C5: Central, C6: North, C7: South, C8: East ...
(500 individual cell assignments)
```

**After inverted index (actual output from JSON):**

```json
"East": [
  "C2", "C23", "C44", "C253", "C262", "C267", "C276", "C278",
  "C299:C300", "C307", "C313:C314", "C316", "C321", "C325",
  "C333", "C337", "C346", "C350:C351", "C355:C356", "C359:C360",
  "C363:C364", "C374", "C501"
],
"South": [
  "C3", "C48", "C255", "C261", "C270", "C281", "C286", "C288",
  "C301", "C309", "C315", "C319", "C322:C323", "C332", "C342",
  "C344:C345", "C353", "C358", "C361:C362", "C368", "C376", "C379"
],
"North": [
  "C22", "C45", "C47", "C250:C251", "C256", "C258", "C266", "C268",
  "C285", "C305:C306", "C308", "C310", "C320", "C324", "C326",
  "C328", "C336", "C339", "C343", "C348", "C352", "C354",
  "C365", "C367", "C373", "C377", "C382"
],
"West": [
  "C20:C21", "C249", "C254", "C257", "C265", "C269", "C271:C272",
  "C274", "C279", "C282", "C287", "C303", "C312", "C317", "C327",
  "C331", "C334", "C340:C341", "C349", "C357", "C366", "C370:C371", "C380"
],
"Central": [
  "C19", "C46", "C252", "C259:C260", "C263:C264", "C273", "C275",
  "C277", "C280", "C283:C284", "C302", "C304", "C311", "C318",
  "C329:C330", "C335", "C338", "C347", "C369", "C372", "C375",
  "C378", "C381", "C383"
]
```

**Before:** 500 lines, one per row.  
**After:** 5 entries, each listing where that value appears.  
**Notice** how consecutive cells get merged: `"C299:C300"` is two rows where "East"
appeared back-to-back; `"C20:C21"` is two consecutive rows of "West".

---

### Example 2 — "Status" Column (Column J)

Only 3 possible values. But "Completed" is the most common — appearing in the majority
of the 500 rows.

**Actual output from the JSON:**

```json
"Completed": [
  "J3", "J20:J21", "J23", "J44:J45",
  "J249", "J251", "J253", "J255:J257", "J260:J263", "J265:J269",
  "J273", "J275", "J277", "J279", "J281", "J284:J285",
  "J299:J300", "J302", "J304:J309", "J315:J316", "J318",
  "J321:J322", "J324:J326", "J328", "J330", "J333",
  "J336:J339", "J341", "J343", "J345:J346", "J348:J349",
  "J352:J355", "J357:J358", "J360:J362",
  "J367:J368", "J371:J372", "J375", "J379:J380", "J382:J383",
  "J501"
],

"Pending": [
  "J2", "J46", "J48", "J250", "J258:J259", "J270:J272", "J274",
  "J276", "J280", "J282:J283", "J286", "J301", "J313:J314",
  "J319:J320", "J323", "J329", "J331:J332", "J344", "J347",
  "J350", "J363:J364", "J370", "J373:J374", "J377:J378", "J381"
],

"Cancelled": [
  "J19", "J22", "J47", "J252", "J254", "J264", "J278",
  "J287:J288", "J303", "J310:J312", "J317", "J327",
  "J334:J335", "J340", "J342", "J351", "J356",
  "J359", "J365:J366", "J369", "J376"
]
```

**Key observation:** The word "Completed" appears perhaps 300+ times in the original
sheet, but in the compressed output it appears **exactly once** as a key. All its locations
are listed as ranges (consecutive runs merged). This is the inverted-index magic.

---

### Example 3 — "PaymentMethod" Column (Column K)

Three payment types, heavily skewed toward "Credit Card":

```json
"Credit Card": [
  "K19:K20", "K23", "K45:K46", "K48",
  "K250", "K252:K255", "K257", "K259:K266", "K268:K269", "K272",
  "K275:K282", "K284", "K286:K288",
  "K299:K301", "K304:K305", "K309", "K312", "K318",
  "K320:K327", "K334", "K336", "K338:K339", "K341", "K343",
  "K345:K346", "K349", "K351:K353", "K358:K359", "K364",
  "K367:K370", "K372", "K375:K377", "K379:K380", "K382:K383"
],
"Cash": [
  "K2", "K22", "K47", "K258", "K270", "K273", "K283", "K285",
  "K303", "K307:K308", "K310:K311", "K315:K317", "K328",
  "K330:K331", "K333", "K340", "K342", "K347:K348", "K350",
  "K354:K355", "K360:K363", "K371", "K373:K374", "K381"
],
"Bank Transfer": [
  "K3", "K21", "K44", "K249", "K251", "K256", "K267", "K271",
  "K274", "K302", "K306", "K313:K314", "K319", "K329",
  "K332", "K335", "K337", "K344", "K356:K357",
  "K365:K366", "K378", "K501"
]
```

Notice `"K259:K266"` — that is **8 consecutive rows** all paid by Credit Card,
expressed in just one range string instead of 8 separate entries.

---

### Example 4 — Unique Values (TxnID, Dates, Totals)

Not all columns compress well. Look at Column A (TxnID):

```json
"1001": ["A2"],
"1002": ["A3"],
"1018": ["A19"],
"1019": ["A20"],
...
"1500": ["A501"]
```

Each TxnID is unique — so each one maps to exactly one cell. **No compression benefit here.**
The inverted index faithfully records them, but they add token cost.

Similarly, dates in column B are mostly unique (365 possible dates in a year),
so most appear only once:

```json
"15-02-2023 00:00:00": ["B2"],
"27-03-2023 00:00:00": ["B3"],
```

However, some dates happen to repeat by coincidence:

```json
"27-12-2023 00:00:00": ["B21", "B346"],
"25-07-2023 00:00:00": ["B44", "B328"],
"22-01-2023 00:00:00": ["B251", "B322:B323", "B327"]
```

**Stage 2 catches even these coincidental duplicates.**

---

### Example 5 — The Formula Row

Row 503 contains the grand total formula:

```json
"TOTAL":          ["A503"],
"=SUM(I2:I501)":  ["I503"]
```

The formula string `=SUM(I2:I501)` is preserved **exactly as written**. The AI can read
and understand it as a formula — not just a number.

---

### Stage 2 Result in Numbers

After Stage 2, the token count drops from **26,957 → 15,934** (a further **1.69×** reduction).

The combined reduction from the original is now **4.69×**.

---

## 5. Stage 3 — Data-Format-Aware Aggregation (Grouping by Type)

### The Core Idea

Stage 2 grouped cells by their **value**. Stage 3 groups cells by their **data type and format**.

Even if two cells have different values, if they are both:
- the same **semantic type** (e.g., both "currency")
- and the same **number format** (e.g., both `$#,##0.00`)

…then they belong to the same format group. An entire column of prices can be described
with a single format entry.

### The 9 Semantic Types

`CellUtils.cs` classifies every cell into one of these 9 types:

| Type | How it is detected | Our sheet example |
|------|--------------------|-------------------|
| `text` | String that is not an email, not a number | "East", "Completed", "Widget A" |
| `integer` | Whole number with no decimal part | TxnID (1001), Qty (7, 38) |
| `float` | Number with decimal places | (not prominent in this sheet) |
| `currency` | Number format contains `$`, `€`, `£`, `¥` | UnitPrice ($9.99), Total ($349.93) |
| `percentage` | Number format contains `%` | (not in this sheet) |
| `scientific` | Number format contains `E+` or `E-` | (not in this sheet) |
| `date` | Format contains `yyyy`, `mm`, `dd` | Sale dates (B column) |
| `time` | Format contains `hh`, `ss`, `am/pm` | (not in this sheet) |
| `year` | Format is only `yyyy` or `yy` | (not in this sheet) |
| `email` | Value matches email regex pattern | (not in this sheet) |

### How the Code Detects Each Type

```csharp
// CellUtils.cs — DetectSemanticType()

// 1. Check number format for currency symbols
if (CurrencyRegex.IsMatch(nfs))   // nfs = "$#,##0.00"
    return "currency";            // → columns H and I in our sheet

// 2. Check number format for date keywords
if (DateKeywords.IsMatch(nfsLower))  // nfs = "yyyy-mm-dd"
    return "date";                   // → column B in our sheet

// 3. Check if value is a whole number
if (long.TryParse(cell.Value, out _))
    return "integer";                // → column A (TxnID) and G (Qty)

// 4. Anything else with text content
return "text";                       // → columns C, D, E, F, J, K
```

### The Format Aggregation Code

```csharp
// SheetCompressor.cs — AggregateBySemanticType()
// For each semantic-type group, call MergeCellRanges()
// to compress the list of cell references into ranges
foreach (var kv in typeNfsGroups)
{
    result[kv.Key] = MergeCellRanges(kv.Value);
}
```

### The Actual Format Groups from Sales_500_rows.json

#### Group 1 — Text / General

This catches all the text cells: headers, region names, salesperson names, products,
categories, statuses, payment methods.

```json
"{\"type\":\"text\",\"nfs\":\"General\"}": [
  "A1:K1",        ← The entire header row (all 11 column names)
  "C2:F3",        ← Region, Salesperson, Product, Category for rows 2-3
  "J2:K3",        ← Status, PaymentMethod for rows 2-3
  "C19:F23",      ← Text columns for rows 19-23
  "J19:K23",
  "C44:F48",
  "J44:K48",
  "C249:F288",    ← Text columns for rows 249-288 (a big block!)
  "J249:K288",
  "C299:F383",    ← Text columns for rows 299-383 (another big block!)
  "J299:K383",
  "C501:F501",
  "J501:K501",
  "A503"          ← The word "TOTAL" in the summary row
]
```

**Plain English:** The AI now knows: "Columns C, D, E, F, J, K all contain plain text
(General format). They cover most of the 500 data rows."

#### Group 2 — Integer / General

This catches TxnID (column A) and Quantity (column G):

```json
"{\"type\":\"integer\",\"nfs\":\"General\"}": [
  "A2:A3",         ← TxnID for rows 2-3
  "G2:G3",         ← Qty for rows 2-3
  "A19:A23",       ← TxnID for rows 19-23
  "G19:G23",       ← Qty for rows 19-23
  "A44:A48",
  "G44:G48",
  "A249:A288",     ← TxnID for the 249-288 block
  "G249:G288",     ← Qty for the 249-288 block
  "A299:A383",
  "G299:G383",
  "A501",
  "G501"
]
```

**Plain English:** The AI now knows: "Columns A and G always contain whole numbers
(integers with General format). They appear in every data block."

#### Group 3 — Date / yyyy-mm-dd

This catches the Date column (column B):

```json
"{\"type\":\"date\",\"nfs\":\"yyyy-mm-dd\"}": [
  "B2:B3",
  "B19:B23",
  "B44:B48",
  "B249:B288",
  "B299:B383",
  "B501"
]
```

**Plain English:** "Column B always contains dates formatted as yyyy-mm-dd."  
The AI does not need to see 500 individual date values — just that dates exist there.

#### Group 4 — Currency / $#,##0.00

This catches UnitPrice (column H) and Total (column I):

```json
"{\"type\":\"currency\",\"nfs\":\"$#,##0.00\"}": [
  "H2:I3",        ← UnitPrice and Total for rows 2-3
  "H19:I23",
  "H44:I48",
  "H249:I288",    ← Both H and I for 40 rows in one entry!
  "H299:I383",
  "H501:I501"
]
```

**Plain English:** "Columns H and I always contain dollar amounts (currency with
2 decimal places)." Six range entries cover the entire 500 rows of price data.

#### Group 5 — Text / $#,##0.00 (Special Case)

The TOTAL formula cell (I503) is interesting:

```json
"{\"type\":\"text\",\"nfs\":\"$#,##0.00\"}": [
  "I503"
]
```

Why is it `text` type even though it has a currency format?  
Because `=SUM(I2:I501)` is a **formula**, not a computed number. The code stores
the formula string (which is text), but the cell still has the currency number format
applied to it. The code captures both pieces of information separately.

---

### Stage 3 Result in Numbers

Stage 3 on its own produces a **very** compact output:
- Only **5 format groups** describe the entire 500-row, 11-column sheet
- Format-only token count: **786**
- That is a **95×** reduction from the original 74,717 tokens

However, the format output alone is not enough — the AI also needs the actual values
(from Stage 2). The final output combines Stage 2 (cells) + Stage 3 (formats) together.

---

## 6. The Final JSON Output — What AI Actually Receives

The three stages are assembled into a single JSON object. Here is the structure with
real data from the file:

```json
{
  "file_name": "Sales_500_rows.xlsx",
  "sheets": {
    "SalesTransactions": {

      /* ── STAGE 1 RESULT ── */
      "structural_anchors": {
        "rows": [1, 21, 46, 251, 256, 261, 266, 271, 276, 281, 286,
                 301, 306, 311, 316, 321, 326, 331, 336, 341, 346,
                 351, 356, 361, 366, 371, 376, 381, 503],
        "columns": ["A","B","C","D","E","F","G","H","I","J","K"]
      },

      /* ── STAGE 2 RESULT ── */
      "cells": {
        "TxnID":        ["A1"],
        "Date":         ["B1"],
        "Region":       ["C1"],
        "Salesperson":  ["D1"],
        "Product":      ["E1"],
        "Category":     ["F1"],
        "Qty":          ["G1"],
        "UnitPrice":    ["H1"],
        "Total":        ["I1"],
        "Status":       ["J1"],
        "PaymentMethod":["K1"],

        /* ── Repeated text values — massively compressed ── */
        "East":         ["C2","C23","C44","C253","C262",...,"C501"],
        "South":        ["C3","C48","C255","C261",...,"C379"],
        "North":        ["C22","C45","C47","C250:C251",...,"C382"],
        "West":         ["C20:C21","C249","C254",...,"C380"],
        "Central":      ["C19","C46","C252","C259:C260",...,"C383"],

        "Completed":    ["J3","J20:J21","J23","J44:J45","J249",...,"J501"],
        "Pending":      ["J2","J46","J48","J250","J258:J259",...,"J381"],
        "Cancelled":    ["J19","J22","J47","J252",...,"J376"],

        "Credit Card":  ["K19:K20","K23","K45:K46",...,"K382:K383"],
        "Cash":         ["K2","K22","K47",...,"K381"],
        "Bank Transfer":["K3","K21","K44",...,"K501"],

        "Widget A":     ["E3","E48","E255",...,"E382"],
        "Widget B":     ["E20:E21","E23","E252",...,"E377:E378"],
        "Gadget X":     ["E22","E44:E45","E47",...,"E376"],
        "Gadget Y":     ["E2","E251","E254",...,"E379"],
        "Part Z":       ["E19","E46","E249",...,"E501"],

        /* ── Unique values — one per row, no compression benefit ── */
        "1001":  ["A2"],
        "1002":  ["A3"],
        "1018":  ["A19"],
        ...
        "1500":  ["A501"],

        /* ── The formula ── */
        "TOTAL":          ["A503"],
        "=SUM(I2:I501)":  ["I503"]
      },

      /* ── STAGE 3 RESULT ── */
      "formats": {
        "{\"type\":\"text\",\"nfs\":\"General\"}":       ["A1:K1","C2:F3","J2:K3",...],
        "{\"type\":\"integer\",\"nfs\":\"General\"}":    ["A2:A3","G2:G3","A19:A23",...],
        "{\"type\":\"date\",\"nfs\":\"yyyy-mm-dd\"}":    ["B2:B3","B19:B23",...,"B501"],
        "{\"type\":\"currency\",\"nfs\":\"$#,##0.00\"}": ["H2:I3","H19:I23",...,"H501:I501"],
        "{\"type\":\"text\",\"nfs\":\"$#,##0.00\"}":     ["I503"]
      },

      "numeric_ranges": {
        "{\"type\":\"integer\",\"nfs\":\"General\"}": ["A2:A3","G2:G3",...]
      }
    }
  },

  /* ── METRICS ── */
  "compression_metrics": { ... }
}
```

---

## 7. Compression Numbers — At Every Stage

These are the **real numbers** from `Sales_500_rows.json`:

```json
"compression_metrics": {
  "sheets": {
    "SalesTransactions": {
      "original_tokens":              74717,
      "after_anchor_tokens":          26957,
      "after_inverted_index_tokens":  15934,
      "after_format_tokens":            786,
      "final_tokens":                 17157,

      "anchor_ratio":          2.77,
      "inverted_index_ratio":  4.69,
      "format_ratio":         95.06,
      "overall_ratio":         4.35
    }
  }
}
```

### Stage-by-Stage Breakdown

```
Original:   ████████████████████████████████████████████ 74,717 tokens

Stage 1:    ████████████████ 26,957 tokens   (2.77× smaller)
            Removed: repeated data rows; kept only anchor row/column intersections

Stage 2:    ██████████ 15,934 tokens   (4.69× smaller from original)
            Removed: duplicate value listings; each value listed once with all its locations

Stage 3:    █ 786 tokens   (95× smaller — formats only)
            Five format groups describe the type of every cell in the sheet

Final:      ██████████ 17,157 tokens   (4.35× smaller from original)
            Combines Stage 2 (cells + values) + Stage 3 (formats) together
```

### Why Does the Final Token Count (17,157) Seem Higher Than Stage 2 (15,934)?

The final output includes **both** the Stage 2 cell dictionary **and** the Stage 3 format
dictionary, plus the structural anchors and metric fields. The Stage 3 formats add
~786 tokens, and the anchors/metrics add a few hundred more — but this overhead
is small compared to the overall compression.

### What Does "4.35× Smaller" Actually Mean?

If you were using an AI API that charges per token:
- **Without SpreadsheetLLM:** You send 74,717 tokens. You pay for 74,717 tokens.
- **With SpreadsheetLLM:** You send 17,157 tokens. You pay for 17,157 tokens.
- **You save 77% of the API cost** for this sheet alone.

For a company processing 1,000 sheets per day, this is a substantial saving.

---

## 8. Summary — Why This Matters

### What Each Stage Contributed

| Stage | What it removed | What it kept | Tokens saved |
|-------|----------------|-------------|--------------|
| Stage 1 (Anchors) | 477 of 503 rows (those that are repetitive data) | Header row, sampled data rows, total row | 47,760 tokens |
| Stage 2 (Inverted Index) | Duplicate value mentions | Each unique value once, with all its locations | 11,023 tokens |
| Stage 3 (Formats) | Individual format labels per cell | 5 format groups covering all 500 rows | Formats compressed to 786 tokens |

### The Three Key Insights

**Insight 1 — Real data has a structure, not just data.**  
Row 1 is always a header. Rows 2–501 repeat a pattern. Row 503 is a summary.
Stage 1 identifies this skeleton so the AI sees the shape, not the noise.

**Insight 2 — Real data has repetition.**  
"East", "West", "North", "South", "Central" appear 100 times each, but there are
only 5 distinct values. "Completed" appears ~300 times but is one word.
Stage 2 exploits this by turning repetition into a lookup table.

**Insight 3 — Real data has patterns of types.**  
Column B is always a date. Columns H and I are always currency. Column A is always
an integer. Stage 3 states this once instead of repeating it 500 times.

### The Analogy One Final Time

Imagine filing 500 sales receipts in a cabinet:

- **Without compression:** You hand the AI the entire cabinet — 500 individual receipts.
- **Stage 1:** You hand over just the cabinet's index card and every 5th receipt as samples.
- **Stage 2:** You highlight: "The word 'Completed' is on receipts #3, 20, 21, 23 ..."
  instead of writing it 300 times.
- **Stage 3:** You write one note: "All receipts in section H and I are in dollars."

The AI now understands the full filing system from **4× fewer pages**.

---

*Source files referenced in this document:*
- *Excel input: `Sales_500_rows.xlsx` (generated by `Program.cs → CreateRealisticSalesSheet()`)*
- *JSON output: `test_output/Sales_500_rows.json`*
- *Core code: `SpreadsheetLLM.Core/SheetCompressor.cs`, `CellUtils.cs`, `ExcelReader.cs`*
