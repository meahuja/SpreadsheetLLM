# Large Sheet (50 rows × 10 cols) — SpreadsheetLLM Deep Dive

## What This Document Is

This document walks you through exactly what happens when SpreadsheetLLM compresses the `Large_sheet__50r×10c_.xlsx` file. Every number, every cell reference, and every code snippet comes from the **real output** in `test_output/Large_sheet__50r×10c_.json`.

Think of it as a detective story: we start with a raw 50-employee spreadsheet, and we watch the system decide what to keep, what to group, and what to throw away — step by step.

---

## 1. What the Raw Sheet Looks Like

The sheet is named **"Large"** and contains a fictional company employee directory:

| Column | Letter | Data type | Example value |
|--------|--------|-----------|---------------|
| ID | A | Integer | 1, 2, 3 … 50 |
| Name | B | Text | Employee1, Employee2 … |
| Dept | C | Text | Engineering, Sales, HR, Finance, Marketing |
| Salary | D | Currency `$#,##0` | $92,275, $66,941 … |
| YearsExp | E | Integer | 13, 18, 15 … |
| Rating | F | Float | 3.1, 3.9, 4.1 … |
| Active | G | Text | "Yes" or "No" |
| StartDate | H | Date `yyyy-mm-dd` | 26-01-2021 |
| Email | I | Text (email) | emp1@company.com |
| Notes | J | Text | "Active" or "" (empty) |

**Size:** Row 1 = header. Rows 2–51 = 50 employee records. 10 columns (A–J). **510 cells total** (including many empty Notes cells).

**The goal:** Send this to an AI (LLM) using as few characters as possible, without losing the important structure.

---

## 2. The Big Picture Numbers

| Stage | What happened | Tokens after | Ratio vs original |
|-------|--------------|-------------|-------------------|
| Original (raw) | Vanilla row-by-row encoding | **5,761** | 1.00× (baseline) |
| After Stage 1 (Anchors) | Only important rows/cols kept | **2,797** | **2.06×** smaller |
| After Stage 2 (Inverted Index) | Same value → one entry | **2,615** | **2.20×** smaller |
| After Stage 3 (Format Groups) | Formats summarised by type | **761** | **7.57×** smaller |
| **Final JSON** (full encoding) | Everything assembled | **3,794** | **1.52×** smaller |

> **Why is the "final" larger than "after format"?** The final JSON also includes the structural anchors list, the inverted-index cells, and the format groups all together. The 761 is just the format dictionary alone; the complete assembled JSON totals 3,794 tokens. Even so, that is 1.52× smaller than the original 5,761 — already saving 33%.

**Key insight:** For this 50-row employee sheet, each row tends to have unique values (different name, different salary, different date). That means less repetition, so compression is more modest compared to the 500-row Sales sheet (which got 4.35×). The format aggregation (Stage 3) still delivers 7.57× on its own because all 50 salary cells share one format type.

---

## 3. Stage 1 — "Find the Important Rows and Columns"

### Plain English

Imagine you are summarising a 50-page employee binder for your boss. You would not copy every page — you would copy the cover page, the first few employees (to show the pattern), and the last few (to show how it ends). Stage 1 does exactly that, but using code.

### The 6 Sub-steps

#### Sub-step 1a: Row Fingerprinting

For every row, the code computes a single integer "fingerprint" — a hash of all values, styles, and merge status. If two adjacent rows have the **same fingerprint**, they look identical to the system.

```csharp
// SheetCompressor.cs — AnalyzeRowsSinglePass()
int fp = 0;
foreach (var part in fpParts)
{
    fp = CombineHash(CombineHash(fp, part.valHash),
                     CombineHash(part.isMerged ? 1 : 0, part.styleId));
}
```

`CombineHash` is a simple XOR-shift mix. Every different cell value or style produces a different hash.

**What the Large sheet looks like:**
- Row 1 hash is unique — it is all bold text headers.
- Rows 2–51 each have a different employee name, salary, and date → different fingerprints every row.

#### Sub-step 1b: Header Detection

The code examines each row to see if it "looks like a header":

```csharp
// SheetCompressor.cs — AnalyzeRowsSinglePass()
if ((double)boldCount / populated > HeaderThreshold)     isHeader = true;
else if ((double)centerCount / populated > HeaderThreshold) isHeader = true;
else if ((double)borderCount / populated > HeaderThreshold) isHeader = true;
else if (stringCount > 0 && (double)capsCount / stringCount > HeaderThreshold) isHeader = true;
```

`HeaderThreshold = 0.6` (60%). If more than 60% of populated cells in a row are bold, centred, bordered, or ALL CAPS, the row is flagged as a header.

**For the Large sheet:**
- Row 1 has 10 bold cells out of 10 populated → boldCount/populated = 10/10 = 1.00 > 0.6 → **header!**
- Rows 2–51 have no bold cells → not headers.

#### Sub-step 1c: Boundary Candidate Detection

Every time the fingerprint **changes** between two adjacent rows, both rows become "boundary candidates":

```csharp
// SheetCompressor.cs — FindBoundaryCandidates()
if (rIdx > 0 && rowInfo.Fingerprints[rIdx] != rowInfo.Fingerprints[rIdx - 1])
{
    rowBoundarySet.Add(r);
    if (r > 1) rowBoundarySet.Add(r - 1);
}
```

Header rows also trigger boundaries on both sides. In the 50-row employee sheet, since **every single row has a unique fingerprint** (different employee data), every row becomes a boundary candidate — potentially all 51 rows.

#### Sub-step 1d: Boundary Capping (The Sampling Step)

51 candidates is manageable, but real spreadsheets can have thousands. To stay within `MaxBoundaryRows = 100`, the code samples evenly:

```csharp
// SheetCompressor.cs — FindBoundaryCandidates()
if (sortedRows.Count > MaxBoundaryRows)   // MaxBoundaryRows = 100
{
    int step = sortedRows.Count / MaxBoundaryRows;
    var sampled = new HashSet<int>();
    for (int i = 0; i < sortedRows.Count; i += step)
        sampled.Add(sortedRows[i]);
    sampled.Add(sortedRows[0]);                          // always keep first
    sampled.Add(sortedRows[sortedRows.Count - 1]);       // always keep last
    foreach (var hr in headerRows) sampled.Add(hr);      // always keep headers
    sortedRows = sampled.OrderBy(r => r).ToList();
}
```

For the Large sheet: 51 candidates ≤ MaxBoundaryRows=100, so **no sampling needed**. All boundary rows are kept.

#### Sub-step 1e: NMS — Removing Overlapping Regions

`NMS (Non-Maximum Suppression)` removes candidate regions that overlap too much. Think of it as: if two candidate table regions are 90% the same, only keep the better-scoring one.

Score = header rows × 10 + filled cell count. For this sheet, the single contiguous region (rows 1–51, cols A–J) survives NMS.

#### Sub-step 1f: k-Expansion (Adding Context Rows)

Each selected anchor row gets a "buffer zone" of **k=2** rows in each direction:

```csharp
// SheetCompressor.cs — ExpandAnchors()
foreach (int anchor in rowAnchors)
    for (int i = Math.Max(1, anchor - k); i <= Math.Min(maxRow, anchor + k); i++)
        keptRows.Add(i);
```

**Actual anchor rows selected: `[1, 4, 10, 46, 47, 48, 51]`**

After k=2 expansion:
- Anchor 1 → keep rows max(1,1-2)=1 to min(51,1+2)=3 → **rows 1, 2, 3**
- Anchor 4 → keep rows 2 to 6 → **rows 2, 3, 4, 5, 6**
- Anchor 10 → keep rows 8 to 12 → **rows 8, 9, 10, 11, 12**
- Anchor 46 → keep rows 44 to 48 → **rows 44, 45, 46, 47, 48**
- Anchor 47 → keep rows 45 to 49 (overlaps above) → **rows 44–49**
- Anchor 48 → keep rows 46 to 50 (overlaps above) → **rows 44–50**
- Anchor 51 → keep rows 49 to 51 → **rows 49, 50, 51**

**Merged kept rows: rows 1–12 and rows 44–51**

**All 10 columns (A–J) are retained** because all column fingerprints differ (each column has a different pattern of values).

**Rows dropped: 13–43** — that's 31 rows of "middle" employees who are indistinguishable from the selected sample rows.

---

## 4. Stage 2 — "Flip the Map: Value → Cells"

### Plain English

In the raw spreadsheet, each cell has an address and a value: `G2 = "Yes"`, `G3 = "Yes"`, `G4 = "Yes"`. Stage 2 flips this around: group all addresses that share the same value into one entry: `"Yes" → ["G2:G5", "G9:G10", "G12", "G44:G51"]`.

This is the **inverted index** — like the index at the back of a book that lists every page number where a word appears.

### Real Examples from the Large Sheet

#### Example 1: "Yes" (Active column)

After Stage 1 filtering, these cells survived with value "Yes":
- G2, G3, G4, G5 (consecutive) → merged to `G2:G5`
- G9, G10 (consecutive) → merged to `G9:G10`
- G12 (alone) → stays `G12`
- G44, G45, G46, G47, G48, G49, G50, G51 (consecutive) → merged to `G44:G51`

**In JSON:**
```json
"Yes": ["G2:G5", "G9:G10", "G12", "G44:G51"]
```

Instead of listing 14 separate cells, 1 inverted index entry covers all of them.

#### Example 2: "Sales" (Dept column)

"Sales" appears at: C2, C11, C12, C44, C51

```json
"Sales": ["C2", "C11:C12", "C44", "C51"]
```

Note: C11 and C12 are consecutive, so they merge. C2, C44, C51 are isolated (far apart), so they stay separate.

#### Example 3: "Marketing" (Dept column)

"Marketing" appears at: C3, C4, C6

```json
"Marketing": ["C3:C4", "C6"]
```

C3 and C4 are consecutive → `C3:C4`. C6 is isolated.

#### Example 4: Unique values (most employee data)

Most cells contain unique values — every employee has a unique name, salary, and date:

```json
"Employee1": ["B2"],
"92275":     ["D2"],
"26-01-2021 00:00:00": ["H2"],
"emp1@company.com": ["I2"]
```

These cannot be merged because no other cell shares the same value. Unique values still get one entry each — they just point to a single cell.

#### Example 5: "3" as a shared value

The number 3 appears in multiple contexts in the filtered rows:
- A4 (employee ID 3)
- E9, E10 (YearsExp = 3)
- F44 (Rating = 3.0 stored as integer)
- F48 (Rating = 3.0 stored as integer)

```json
"3": ["A4", "E9:E10", "F44", "F48"]
```

This is an interesting case: the same raw string "3" appears in different columns for completely different purposes (ID vs years of experience vs rating). The inverted index groups them all together regardless of meaning.

### How Range Merging Works

The `MergeCellRanges` method sorts cell references and merges any that are consecutive **in the same column**:

```csharp
// SheetCompressor.cs — MergeCellRanges()
// Consecutive cells in same column merge: A1, A2, A3 → A1:A3
// Cells in different columns do NOT merge: A1, B1 remain separate
```

---

## 5. Stage 3 — "Group by Data Type"

### Plain English

Now the system looks at every remaining cell and asks: "What type of data is in this cell, and what display format does it use?" All cells that share the same type + format get grouped together into one entry.

Think of it like sorting your coins: instead of listing every coin individually, you say "I have 20 pennies, 5 nickels, 3 dimes."

### The 6 Format Groups (Actual Output)

The Large sheet produces exactly **6 format groups**:

#### Group 1: `{"type":"text","nfs":"General"}`
```json
"A1:J1",   ← entire header row
"B2:C6",   ← Name + Dept, rows 2-6
"G2:G6",   ← Active (Yes/No), rows 2-6
"J3",      ← Notes cell with "Active"
"J5",
"B8:C12",  ← Name + Dept, rows 8-12
"G8:G12",  ← Active (Yes/No), rows 8-12
"J8",
"J10:J11",
"B44:C51", ← Name + Dept, rows 44-51
"G44:G51", ← Active (Yes/No), rows 44-51
"J48"
```

All header text, department names, Active flags, and Notes are plain text with no special format code.

#### Group 2: `{"type":"integer","nfs":"General"}`
```json
"A2:A6",   ← Employee IDs 1-5
"E2:E6",   ← Years of experience, rows 2-6
"A8:A12",  ← Employee IDs 7-12
"E8:E12",  ← Years of experience, rows 8-12
"F11",     ← Rating=4 (happens to be a whole number)
"A44:A51", ← Employee IDs 43-50
"E44:F44", ← YearsExp + Rating for row 44 (both integers)
"E45:E51", ← Years of experience, rows 45-51
"F48"      ← Rating=3 (whole number)
```

Employee IDs and years of experience are whole numbers stored as General format.

#### Group 3: `{"type":"currency","nfs":"$#,##0"}`
```json
"D2:D6",   ← Salaries, rows 2-6
"D8:D12",  ← Salaries, rows 8-12
"D44:D51"  ← Salaries, rows 44-51
```

All 25 salary values in the kept rows collapse into just **3 range entries** — one per block of consecutive kept rows.

#### Group 4: `{"type":"float","nfs":"General"}`
```json
"F2:F6",     ← Ratings 3.1–4.9, rows 2-6
"F8:F10",    ← Ratings, rows 8-10
"F12",       ← Rating row 12
"F45:F47",   ← Ratings, rows 45-47
"F49:F51"    ← Ratings, rows 49-51
```

Decimal ratings (like 3.9, 4.1, 4.5) are stored as General format floats.

#### Group 5: `{"type":"date","nfs":"yyyy-mm-dd"}`
```json
"H2:H6",   ← Dates, rows 2-6
"H8:H12",  ← Dates, rows 8-12
"H44:H51"  ← Dates, rows 44-51
```

All 25 date cells in the kept rows reduce to **3 range entries**. The AI now knows that every cell in these ranges is a date in `yyyy-mm-dd` format — it does not need the individual values listed.

#### Group 6: `{"type":"email","nfs":"General"}`
```json
"I2:I6",   ← Emails, rows 2-6
"I8:I12",  ← Emails, rows 8-12
"I44:I51"  ← Emails, rows 44-51
```

The system detects emails using a regex pattern for `@` and `.com/.org` etc. All email cells group together.

### Why Stage 3 Alone Gets 7.57× Compression

The format dictionary alone went from 2,615 tokens → 761 tokens (7.57×). This is because:

- Instead of listing 25 individual date values → 3 range strings
- Instead of listing 25 individual salary amounts → 3 range strings
- Instead of listing 25 individual email addresses → 3 range strings

The AI still knows "columns H rows 2–51 are dates in yyyy-mm-dd format" without needing to see every single date.

---

## 6. The Final JSON Structure

Here is what the complete output looks like, with real values from the actual JSON:

```json
{
  "file_name": "Large_sheet__50r×10c_.xlsx",
  "sheets": {
    "Large": {
      "structural_anchors": {
        "rows": [1, 4, 10, 46, 47, 48, 51],
        "columns": ["A","B","C","D","E","F","G","H","I","J"]
      },
      "cells": {
        "ID": ["A1"],
        "Name": ["B1"],
        "Yes": ["G2:G5", "G9:G10", "G12", "G44:G51"],
        "Sales": ["C2", "C11:C12", "C44", "C51"],
        "Marketing": ["C3:C4", "C6"],
        "Engineering": ["C5", "C9", "C45", "C50"],
        "...": "... (130+ unique entries total) ..."
      },
      "formats": {
        "{\"type\":\"text\",\"nfs\":\"General\"}":     ["A1:J1", "B2:C6", ...],
        "{\"type\":\"integer\",\"nfs\":\"General\"}":  ["A2:A6", "E2:E6", ...],
        "{\"type\":\"currency\",\"nfs\":\"$#,##0\"}":  ["D2:D6", "D8:D12", "D44:D51"],
        "{\"type\":\"float\",\"nfs\":\"General\"}":    ["F2:F6", "F8:F10", ...],
        "{\"type\":\"date\",\"nfs\":\"yyyy-mm-dd\"}":  ["H2:H6", "H8:H12", "H44:H51"],
        "{\"type\":\"email\",\"nfs\":\"General\"}":    ["I2:I6", "I8:I12", "I44:I51"]
      },
      "numeric_ranges": {
        "{\"type\":\"integer\",\"nfs\":\"General\"}": ["A2:A6", "E2:E6", ...],
        "{\"type\":\"float\",\"nfs\":\"General\"}":   ["F2:F6", "F8:F10", ...]
      }
    }
  },
  "compression_metrics": {
    "sheets": {
      "Large": {
        "original_tokens": 5761,
        "after_anchor_tokens": 2797,
        "after_inverted_index_tokens": 2615,
        "after_format_tokens": 761,
        "final_tokens": 3794,
        "anchor_ratio": 2.06,
        "inverted_index_ratio": 2.20,
        "format_ratio": 7.57,
        "overall_ratio": 1.52
      }
    }
  }
}
```

---

## 7. Token Count Visualisation

```
Original (5,761 tokens)
████████████████████████████████████████████████████████ 5761
                                                          ← 50 employees, all unique data

After Stage 1 – Anchor Extraction (2,797 tokens)
█████████████████████████████ 2797  (2.06× smaller)
                                                          ← Only rows 1-12 + 44-51 kept
                                                          ← Rows 13-43 dropped (31 rows)

After Stage 2 – Inverted Index (2,615 tokens)
███████████████████████████ 2615  (2.20× smaller)
                                                          ← Repeated values grouped:
                                                          ← "Yes" at 14 cells → 4 entries
                                                          ← dept names consolidated

After Stage 3 – Format Groups (761 tokens — format dict alone)
████████ 761  (7.57× smaller — format dict only)
                                                          ← 25 dates → 3 range strings
                                                          ← 25 salaries → 3 range strings
                                                          ← 25 emails → 3 range strings

Final JSON (3,794 tokens — full assembled output)
███████████████████████████████████ 3794  (1.52× smaller)
                                                          ← Includes anchors + cells + formats
```

---

## 8. Why Compression is Modest vs the Sales Sheet

The Large sheet gets 1.52× compression overall. The Sales 500-row sheet gets 4.35×. Why the difference?

| Factor | Large Sheet (50 rows) | Sales Sheet (500 rows) |
|--------|----------------------|------------------------|
| Repeated column values | Low — each employee has unique salary, date, email | Very high — only 5 regions, 5 products, 3 statuses |
| Rows dropped by Stage 1 | 31 out of 50 (62%) | Most of 500 (only ~29 kept) |
| Inverted index benefit | Small — most values unique | Large — "Completed" appears 375+ times |
| Format groups | 6 groups, all useful | 5 groups, dates+currency cover most rows |

**The key rule:** SpreadsheetLLM gives the biggest savings on **large, repetitive** spreadsheets. A sheet where every row has a different structure and unique values is harder to compress than one with 500 rows of the same 5 products being sold repeatedly.

Even so, the format aggregation alone (7.57×) means that an AI reading this sheet does not need to parse 50 individual date strings, 50 salary amounts, or 50 email addresses — it just knows the pattern of which cells contain what type, and can look up values in the cells dictionary for specific employees of interest.

---

## 9. Summary — The Filing Cabinet Analogy

Imagine the 50-employee sheet is a filing cabinet with 50 folders.

**Without SpreadsheetLLM** (raw): You hand the AI every single page from every folder — 510 individual cell values to read.

**Stage 1 (Anchors)**: You keep only the **table of contents** (row 1) and **samples** — the first few employees and the last few. The AI sees the pattern without needing all 50 entries (rows 13–43 are dropped).

**Stage 2 (Inverted Index)**: Instead of repeating "Engineering" five times across five files, you create a single sticky note that says: *"Engineering employees are at rows 2, 9, 45, 50."* The AI finds anyone by looking up their value.

**Stage 3 (Format Groups)**: Instead of listing all 25 salary numbers, you write a single note: *"Column D rows 2–51 are all salary amounts in dollar format."* The AI already knows the pattern.

**Result**: The AI gets a 1.52× smaller file that still tells it everything it needs to know: where data lives, what shape the table has, and where specific values appear — without reading every cell individually.
