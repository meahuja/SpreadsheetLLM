---
name: spreadsheet-encoding
description: SpreadsheetLLM encoding patterns and SheetCompressor pipeline. Activate when working on spreadsheet compression, anchor detection, inverted indexing, or format aggregation.
origin: project
---

# SpreadsheetLLM Encoding Patterns

> Paper: arXiv:2407.09025 — SpreadsheetLLM: Encoding Spreadsheets for Large Language Models

## When to Activate

- Implementing or modifying SheetCompressor stages
- Working on boundary detection or anchor extraction
- Implementing inverted index or range merging
- Working on semantic type detection or format aggregation
- Implementing Chain-of-Spreadsheet QA pipeline

## 3-Stage SheetCompressor Pipeline

### Stage 1: Structural-Anchor-Based Extraction
```
Input: Full sheet → Profile rows/cols → Find boundaries → Compose candidates
→ Filter (size, sparsity, header) → NMS → Expand by k → Compress homogeneous
→ Output: kept_rows, kept_cols
```

### Stage 2: Inverted-Index Translation
```
Input: kept cells → Group by value → Merge adjacent refs into ranges
→ Output: {value: [ranges]} (lossless)
```

### Stage 3: Data-Format-Aware Aggregation
```
Input: kept cells → Classify semantic type → Group by (type, NFS)
→ Find contiguous rectangles → Output: {type_key: [ranges]}
```

## 9 Semantic Types (Paper Section 3.4)
1. Year — NFS contains yyyy/yy only (no month/day)
2. Integer — int value or NFS "0" / "#,##0"
3. Float — float value or NFS "0.00" / "#,##0.00"
4. Percentage — NFS contains %
5. Scientific — NFS contains E+/E-
6. Date — NFS contains y/m/d patterns
7. Time — NFS contains h/m/s patterns
8. Currency — NFS contains $€£¥
9. Email — value matches email regex

## CoS QA Prompt Templates (Appendix L.3)

### Stage 1 — Table Identification
```
Given compressed spreadsheet encoding and a question,
identify which table contains the answer.
Return the range like ['range': 'A1:F9'].
```

### Stage 2 — Answer Generation
```
Given uncompressed table encoding and a question,
find the cell address containing the answer.
Return like [B3] or [SUM(A2:A10)].
```

## Key Implementation Patterns

### Range Merging Algorithm
```python
# Greedy maximal rectangle from a set of (row, col) coordinates:
# 1. Sort coordinates
# 2. For each unprocessed cell, expand right while in set
# 3. Then expand down while full row width is in set
# 4. Mark all cells in rectangle as processed
```

### Boundary Detection Signals
- Adjacent row profiles differ (value, merge status, style)
- Data type transition (numeric → text or vice versa)
- Column population width changes
- Empty row/col (always a boundary)
