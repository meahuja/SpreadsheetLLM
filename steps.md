All 16 tests pass (0 failures). Here is a full summary of everything delivered:

Task 1 — Large_Sheet_50r_10c_Deep_Dive.md
Created Large_Sheet_50r_10c_Deep_Dive.md — same style as the Sales deep-dive but for the 50-row employee sheet. Key findings from the real JSON:

Stage	Tokens	Ratio
Original	5,761	baseline
After Stage 1 (anchors)	2,797	2.06×
After Stage 2 (inverted index)	2,615	2.20×
After Stage 3 (format groups)	761 (dict alone)	7.57×
Final JSON	3,794	1.52×
Anchor rows selected: [1, 4, 10, 46, 47, 48, 51] → keeps rows 1–12 and 44–51, drops rows 13–43.

Task 2 — Messy 2000 rows
Added CreateMessy2000 to Program.cs. Features:

8 sections × 250 rows = 2,000 data rows + 7 bold italic sub-header rows (tests mid-table boundary detection)
~20% missing dates, ~33% missing notes, ~25% missing flags
~5% of Amount cells are "N/A" or "TBD" (text in a numeric column — messy!)
Result: 7.11× compression (194,413 → 27,333 tokens) — high repetition in Category/Status columns
Task 3 — plain_adjacent.xlsx in .NET code
Added CreatePlainAdjacent to Program.cs. Recreates the exact same structure as the original file:

Table 1 (rows 1–6): Fruit/Price/Stock/Total Value with =Bn*Cn formulas and AVERAGE/SUM summaries
Table 2 (rows 7–12): Student/Math/Science/Average with AVERAGE formulas
No gap row between tables
Result: 0.36× ratio — identical to the original plain_adjacent.xlsx output ✓
Task 4 — Large Adjacent Report (3 tables, no gaps)
Added CreateLargeAdjacentReport + RunAdjacentReportTest to Program.cs. Produces two JSON files:

File	Description
LargeAdjacentReport_Full.json	All 3 stages: anchors + inverted index + format aggregation
LargeAdjacentReport_InvertedIndexOnly.json	Stage 2 only: ALL cells as inverted index, no anchor filtering, no format grouping
The sheet has 3 tables back-to-back (no gap rows):

Rows 1–151: Product Catalog (ProductID, Name, Category, UnitPrice, StockQty, ReorderPoint)
Rows 152–251: Sales Rep Quotas (RepID, Name, Territory, AnnualQuota, Q1–Q4 Targets)
Rows 252–501: Customer Orders (OrderID, Date, Customer, Product, Qty, Price, Total, Status)
The Full.json shows what the anchor algorithm considers the "interesting" rows across all 3 tables; InvertedIndexOnly.json shows ALL cells grouped by value — useful to see what data the anchor step discards.

Also added 3 new public methods to SheetCompressor.cs:

FindAnchorsOnly(filePath, k) — Stage 1 only
EncodeStage1And2(filePath, k) — Stages 1+2, no format aggregation
EncodeInvertedIndexOnly(filePath) — Stage 2 on ALL cells, no anchor filtering, no format aggregation
Task 5 — Running the Three Techniques One by One
With the new methods added to SheetCompressor, here are the exact steps:

Option A — From the test runner command line

# Run everything (all 3 stages) on any file:
& "C:\Program Files\dotnet\dotnet.exe" run --project SpreadsheetLLM.TestRunner -c Release -- "C:\path\to\your.xlsx"
# JSON saved to test_output\your.json
Option B — From your own C# code, stage by stage

var compressor = new SheetCompressor();
var filePath   = @"C:\path\to\your.xlsx";

// ── Stage 1 only ──────────────────────────────────────────────
// Returns which rows/columns the algorithm considers "important"
var anchors = compressor.FindAnchorsOnly(filePath, k: 2);
foreach (var (sheetName, a) in anchors)
    Console.WriteLine($"{sheetName}: anchor rows={string.Join(",", a.Rows)}");

// ── Stages 1 + 2 (no format aggregation) ─────────────────────
// Shows inverted index AFTER anchor filtering but WITHOUT format grouping
var stage1and2 = compressor.EncodeStage1And2(filePath, k: 2);
// → stage1and2.Sheets["Sheet1"].Cells  has the inverted index
// → stage1and2.Sheets["Sheet1"].Formats is empty

// ── Stage 2 only on ALL cells (no anchor filtering) ──────────
// Shows inverted index of every cell — useful as a comparison baseline
var invIdxOnly = compressor.EncodeInvertedIndexOnly(filePath);
// → invIdxOnly.Sheets["Sheet1"].Cells  has all cells grouped by value
// → invIdxOnly.Sheets["Sheet1"].StructuralAnchors is empty

// ── All 3 stages (full pipeline) ─────────────────────────────
var full = compressor.Encode(filePath, k: 2);
// → full.Sheets["Sheet1"] has anchors + cells + formats + numeric_ranges
What each output tells you
Method	Anchors?	Inverted Index?	Format Groups?	Use when…
FindAnchorsOnly	✓	✗	✗	Debugging which rows Stage 1 picks
EncodeStage1And2	✓	✓ (filtered)	✗	Comparing anchor-filtered vs all-cell index
EncodeInvertedIndexOnly	✗	✓ (all cells)	✗	Seeing full inverted index without any filtering
Encode	✓	✓	✓	Production use — all 3 stages