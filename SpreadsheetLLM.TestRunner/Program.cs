using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using ClosedXML.Excel;
using SpreadsheetLLM.Core;
using SpreadsheetLLM.Core.Models;

/// <summary>
/// Test runner for SpreadsheetLLM.Core — generates sample Excel files and encodes them.
/// Covers: simple table, multi-table, merged cells, formulas, dates/currency,
///         empty sheet, all-numeric, mixed formats, large sheet.
/// </summary>
static class Program
{
    static readonly string OutDir = Path.Combine(
        AppContext.BaseDirectory, "test_output");

    static readonly JsonSerializerOptions JsonOpts = new JsonSerializerOptions
    {
        WriteIndented = true
    };

    static int _pass, _fail;

    // -------------------------------------------------------------------------
    // Entry point
    // -------------------------------------------------------------------------
    static void Main(string[] args)
    {
        Directory.CreateDirectory(OutDir);

        // If a file path is provided as argument, encode just that file and exit.
        if (args.Length > 0)
        {
            var path = args[0];
            if (!File.Exists(path))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"File not found: {path}");
                Console.ResetColor();
                Environment.Exit(1);
            }
            Console.WriteLine($"=== Encoding: {path} ===\n");
            RunFileTest(Path.GetFileName(path), path);
            Console.WriteLine($"\nJSON saved to: {OutDir}");
            Environment.Exit(_fail > 0 ? 1 : 0);
        }

        Console.WriteLine("=== SpreadsheetLLM.Core Test Runner ===\n");

        RunTest("Simple table",         CreateSimpleTable);
        RunTest("Multi-table sheet",    CreateMultiTable);
        RunTest("Merged cells",         CreateMergedCells);
        RunTest("Formulas",             CreateFormulas);
        RunTest("Dates & currency",     CreateDatesCurrency);
        RunTest("All-numeric",          CreateAllNumeric);
        RunTest("Mixed formats",        CreateMixedFormats);
        RunTest("Large sheet (50r×10c)", CreateLargeSheet);
        RunTest("Empty cells sparse",    CreateSparseCells);
        RunTest("Multi-sheet workbook",  CreateMultiSheet);
        RunTest("Sales 500 rows",        CreateRealisticSalesSheet);
        RunTest("HR Payroll 300 rows",   CreateRealisticHRSheet);

        // Task 3 — plain_adjacent recreated in .NET
        RunTest("Plain adjacent (recreated)", CreatePlainAdjacent);

        // Task 2 — messy 2000-row sheet
        RunTest("Messy 2000 rows",       CreateMessy2000);

        // Task 4 — Large adjacent report: generates TWO json files for comparison
        RunAdjacentReportTest();

        // Also encode the existing plain_adjacent.xlsx sample if present
        var samplePath = Path.Combine(
            AppContext.BaseDirectory, "..", "..", "..", "..", "plain_adjacent.xlsx");
        samplePath = Path.GetFullPath(samplePath);
        if (File.Exists(samplePath))
            RunFileTest("plain_adjacent.xlsx", samplePath);

        Console.WriteLine($"\n=== Results: {_pass} passed, {_fail} failed ===");
        Environment.Exit(_fail > 0 ? 1 : 0);
    }

    // -------------------------------------------------------------------------
    // Test helpers
    // -------------------------------------------------------------------------
    static void RunTest(string name, Action<XLWorkbook> populate)
    {
        Console.Write($"  [{name}] ... ");
        var wb = new XLWorkbook();
        try
        {
            populate(wb);
            var path = Path.Combine(OutDir, Sanitize(name) + ".xlsx");
            wb.SaveAs(path);
            RunFileTest(name, path);
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"FAIL (create): {ex.Message}");
            Console.ResetColor();
            _fail++;
        }
        finally { wb.Dispose(); }
    }

    static void RunFileTest(string name, string path)
    {
        Console.Write($"  [{name}] ... ");
        try
        {
            var compressor = new SheetCompressor();
            var encoding = compressor.Encode(path);

            // Assertions
            Assert(encoding != null,                            "encoding is null");
            Assert(encoding!.Sheets.Count > 0,                 "no sheets in output");
            Assert(encoding.CompressionMetrics != null,        "metrics is null");
            Assert(encoding.CompressionMetrics!.Overall != null,"overall metrics is null");

            foreach (var (sheetName, sheet) in encoding.Sheets)
            {
                Assert(sheet.StructuralAnchors != null,
                    $"[{sheetName}] StructuralAnchors is null");
                Assert(sheet.Cells != null,
                    $"[{sheetName}] Cells dict is null");
                Assert(sheet.Formats != null,
                    $"[{sheetName}] Formats dict is null");

                // Every cell-range string must be parseable
                foreach (var (val, ranges) in sheet.Cells)
                    foreach (var r in ranges)
                        AssertValidRange(r, sheetName, "cells");

                foreach (var (key, ranges) in sheet.Formats)
                    foreach (var r in ranges)
                        AssertValidRange(r, sheetName, "formats");
            }

            // Save JSON output
            var json = JsonSerializer.Serialize(encoding, JsonOpts);
            var jsonPath = Path.Combine(OutDir, Sanitize(name) + ".json");
            File.WriteAllText(jsonPath, json);

            var m = encoding.CompressionMetrics.Overall;
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(
                $"PASS  (orig={m.OriginalTokens} final={m.FinalTokens} " +
                $"ratio={m.OverallRatio:F2}x  sheets={encoding.Sheets.Count})");
            Console.ResetColor();
            _pass++;
        }
        catch (AssertionException ae)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"FAIL (assert): {ae.Message}");
            Console.ResetColor();
            _fail++;
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"FAIL (exception): {ex.GetType().Name}: {ex.Message}");
            Console.ResetColor();
            _fail++;
        }
    }

    static void Assert(bool condition, string msg)
    {
        if (!condition) throw new AssertionException(msg);
    }

    static void AssertValidRange(string r, string sheet, string section)
    {
        // Valid formats: "A1" or "A1:B3"
        try
        {
            if (r.Contains(':'))
            {
                var parts = r.Split(':');
                Assert(parts.Length == 2, $"bad range '{r}' in [{sheet}].{section}");
                CellUtils.SplitCellRef(parts[0]);
                CellUtils.SplitCellRef(parts[1]);
            }
            else
            {
                CellUtils.SplitCellRef(r);
            }
        }
        catch
        {
            throw new AssertionException($"unparseable range '{r}' in [{sheet}].{section}");
        }
    }

    // -------------------------------------------------------------------------
    // Task 4 — Two-file output for the large adjacent report
    // File 1: LargeAdjacentReport_Full.json         — all 3 stages applied
    // File 2: LargeAdjacentReport_InvertedIndexOnly.json — Stage 2 only (all cells, no anchors, no formats)
    // -------------------------------------------------------------------------
    static void RunAdjacentReportTest()
    {
        const string name = "Large adjacent report (3 tables, no gaps)";
        Console.Write($"  [{name}] ... ");
        var wb = new XLWorkbook();
        try
        {
            CreateLargeAdjacentReport(wb);
            var xlsxPath = Path.Combine(OutDir, "LargeAdjacentReport.xlsx");
            wb.SaveAs(xlsxPath);

            var compressor = new SheetCompressor();

            // --- File 1: Full pipeline (all 3 stages) ---
            var fullEncoding = compressor.Encode(xlsxPath);
            var fullJson     = JsonSerializer.Serialize(fullEncoding, JsonOpts);
            var fullPath     = Path.Combine(OutDir, "LargeAdjacentReport_Full.json");
            File.WriteAllText(fullPath, fullJson);

            // --- File 2: Inverted index only (Stage 2, all cells, no anchor filter, no formats) ---
            var idxOnlyEncoding = compressor.EncodeInvertedIndexOnly(xlsxPath);
            var idxOnlyJson     = JsonSerializer.Serialize(idxOnlyEncoding, JsonOpts);
            var idxOnlyPath     = Path.Combine(OutDir, "LargeAdjacentReport_InvertedIndexOnly.json");
            File.WriteAllText(idxOnlyPath, idxOnlyJson);

            // Basic assertions on the full encoding
            Assert(fullEncoding != null,                    "full encoding is null");
            Assert(fullEncoding!.Sheets.Count > 0,          "no sheets in full encoding");
            Assert(fullEncoding.CompressionMetrics != null, "metrics is null");

            // Basic assertions on the index-only encoding
            Assert(idxOnlyEncoding != null,              "index-only encoding is null");
            Assert(idxOnlyEncoding!.Sheets.Count > 0,    "no sheets in index-only encoding");
            foreach (var (_, sheet) in idxOnlyEncoding.Sheets)
            {
                Assert(sheet.Cells != null,                         "index-only Cells is null");
                Assert(sheet.Formats.Count == 0,                    "index-only Formats should be empty");
                Assert(sheet.StructuralAnchors.Rows.Count == 0,     "index-only Anchors.Rows should be empty");
            }

            var m = fullEncoding.CompressionMetrics!.Overall;
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(
                $"PASS  (full: orig={m.OriginalTokens} final={m.FinalTokens} ratio={m.OverallRatio:F2}x)");
            Console.ResetColor();
            Console.WriteLine($"        Full encoding   → {fullPath}");
            Console.WriteLine($"        InvIndex only   → {idxOnlyPath}");
            _pass++;
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"FAIL: {ex.GetType().Name}: {ex.Message}");
            Console.ResetColor();
            _fail++;
        }
        finally { wb.Dispose(); }
    }

    static string Sanitize(string s) => s.Replace(' ', '_').Replace('/', '_').Replace('(', '_').Replace(')', '_');

    class AssertionException : Exception { public AssertionException(string msg) : base(msg) { } }

    // -------------------------------------------------------------------------
    // Sample workbook factories
    // -------------------------------------------------------------------------

    /// Simple 5-row sales table with a header row
    static void CreateSimpleTable(XLWorkbook wb)
    {
        var ws = wb.Worksheets.Add("Sales");
        string[] headers = { "Product", "Quantity", "Unit Price", "Total" };
        for (int c = 0; c < headers.Length; c++)
        {
            var cell = ws.Cell(1, c + 1);
            cell.Value = headers[c];
            cell.Style.Font.Bold = true;
        }
        var data = new object?[][] {
            new object?[] { "Apple",  100, 0.50, null },
            new object?[] { "Banana", 200, 0.30, null },
            new object?[] { "Cherry", 150, 1.20, null },
            new object?[] { "Date",    80, 2.00, null },
            new object?[] { "Elderberry", 60, 3.50, null },
        };
        for (int r = 0; r < data.Length; r++)
        {
            for (int c = 0; c < data[r].Length; c++)
            {
                if (data[r][c] != null)
                    ws.Cell(r + 2, c + 1).Value = (dynamic)data[r][c]!;
            }
            // Formula in column D
            ws.Cell(r + 2, 4).FormulaA1 = $"B{r + 2}*C{r + 2}";
        }
    }

    /// Two adjacent tables on the same sheet (multi-table detection)
    static void CreateMultiTable(XLWorkbook wb)
    {
        var ws = wb.Worksheets.Add("MultiTable");

        // Table 1 (rows 1-5, cols A-C)
        ws.Cell(1, 1).Value = "Name";  ws.Cell(1, 1).Style.Font.Bold = true;
        ws.Cell(1, 2).Value = "Score"; ws.Cell(1, 2).Style.Font.Bold = true;
        ws.Cell(1, 3).Value = "Grade"; ws.Cell(1, 3).Style.Font.Bold = true;
        var t1 = new[] { ("Alice",95,"A"), ("Bob",72,"B"), ("Carol",88,"A"), ("Dave",61,"C") };
        for (int i = 0; i < t1.Length; i++)
        {
            ws.Cell(i + 2, 1).Value = t1[i].Item1;
            ws.Cell(i + 2, 2).Value = t1[i].Item2;
            ws.Cell(i + 2, 3).Value = t1[i].Item3;
        }

        // Gap row 6 — empty

        // Table 2 (rows 7-10, cols A-B)
        ws.Cell(7, 1).Value = "Month"; ws.Cell(7, 1).Style.Font.Bold = true;
        ws.Cell(7, 2).Value = "Sales"; ws.Cell(7, 2).Style.Font.Bold = true;
        var t2 = new[] { ("Jan",1200), ("Feb",1450), ("Mar",980) };
        for (int i = 0; i < t2.Length; i++)
        {
            ws.Cell(i + 8, 1).Value = t2[i].Item1;
            ws.Cell(i + 8, 2).Value = t2[i].Item2;
        }
    }

    /// Sheet with merged header cells
    static void CreateMergedCells(XLWorkbook wb)
    {
        var ws = wb.Worksheets.Add("Merged");

        // Merged title row
        ws.Range("A1:D1").Merge().Value = "Quarterly Report";
        ws.Cell(1, 1).Style.Font.Bold = true;
        ws.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        ws.Cell(2, 1).Value = "Region"; ws.Cell(2, 1).Style.Font.Bold = true;
        ws.Cell(2, 2).Value = "Q1";     ws.Cell(2, 2).Style.Font.Bold = true;
        ws.Cell(2, 3).Value = "Q2";     ws.Cell(2, 3).Style.Font.Bold = true;
        ws.Cell(2, 4).Value = "Q3";     ws.Cell(2, 4).Style.Font.Bold = true;

        var data = new[] {
            ("North", 3200, 4100, 3800),
            ("South", 2800, 3100, 2600),
            ("East",  4100, 3800, 4500),
            ("West",  2300, 2700, 2900),
        };
        for (int i = 0; i < data.Length; i++)
        {
            ws.Cell(i + 3, 1).Value = data[i].Item1;
            ws.Cell(i + 3, 2).Value = data[i].Item2;
            ws.Cell(i + 3, 3).Value = data[i].Item3;
            ws.Cell(i + 3, 4).Value = data[i].Item4;
        }
    }

    /// Sheet with SUM/AVERAGE formulas
    static void CreateFormulas(XLWorkbook wb)
    {
        var ws = wb.Worksheets.Add("Formulas");

        ws.Cell(1, 1).Value = "Item";   ws.Cell(1, 1).Style.Font.Bold = true;
        ws.Cell(1, 2).Value = "Value";  ws.Cell(1, 2).Style.Font.Bold = true;

        for (int r = 2; r <= 7; r++)
        {
            ws.Cell(r, 1).Value = $"Item{r - 1}";
            ws.Cell(r, 2).Value = r * 10;
        }

        ws.Cell(8, 1).Value = "Sum";
        ws.Cell(8, 2).FormulaA1 = "SUM(B2:B7)";

        ws.Cell(9, 1).Value = "Average";
        ws.Cell(9, 2).FormulaA1 = "AVERAGE(B2:B7)";

        ws.Cell(10, 1).Value = "Max";
        ws.Cell(10, 2).FormulaA1 = "MAX(B2:B7)";
    }

    /// Sheet with date and currency formatted cells
    static void CreateDatesCurrency(XLWorkbook wb)
    {
        var ws = wb.Worksheets.Add("DatesCurrency");

        ws.Cell(1, 1).Value = "Date";     ws.Cell(1, 1).Style.Font.Bold = true;
        ws.Cell(1, 2).Value = "Amount";   ws.Cell(1, 2).Style.Font.Bold = true;
        ws.Cell(1, 3).Value = "Category"; ws.Cell(1, 3).Style.Font.Bold = true;

        var rows = new (DateTime date, double amount, string cat)[] {
            (new DateTime(2024, 1, 15), 1250.50, "Revenue"),
            (new DateTime(2024, 2, 28), 875.00,  "Expense"),
            (new DateTime(2024, 3, 10), 3200.75, "Revenue"),
            (new DateTime(2024, 4,  5), 450.25,  "Expense"),
        };
        for (int i = 0; i < rows.Length; i++)
        {
            var dateCell = ws.Cell(i + 2, 1);
            dateCell.Value = rows[i].date;
            dateCell.Style.NumberFormat.Format = "yyyy-mm-dd";

            var amtCell = ws.Cell(i + 2, 2);
            amtCell.Value = rows[i].amount;
            amtCell.Style.NumberFormat.Format = "$#,##0.00";

            ws.Cell(i + 2, 3).Value = rows[i].cat;
        }
    }

    /// Sheet of purely numeric data (no headers, triggers edge case)
    static void CreateAllNumeric(XLWorkbook wb)
    {
        var ws = wb.Worksheets.Add("AllNumeric");
        var rng = new Random(42);
        for (int r = 1; r <= 10; r++)
            for (int c = 1; c <= 5; c++)
                ws.Cell(r, c).Value = rng.Next(100, 9999);
    }

    /// Sheet with many different number/text format combos
    static void CreateMixedFormats(XLWorkbook wb)
    {
        var ws = wb.Worksheets.Add("MixedFormats");

        ws.Cell(1, 1).Value = "Label";      ws.Cell(1, 1).Style.Font.Bold = true;
        ws.Cell(1, 2).Value = "Value";      ws.Cell(1, 2).Style.Font.Bold = true;
        ws.Cell(1, 3).Value = "Formatted";  ws.Cell(1, 3).Style.Font.Bold = true;

        ws.Cell(2, 1).Value = "Integer";
        ws.Cell(2, 2).Value = 42;
        ws.Cell(2, 3).Value = 42; ws.Cell(2, 3).Style.NumberFormat.Format = "#,##0";

        ws.Cell(3, 1).Value = "Float";
        ws.Cell(3, 2).Value = 3.14159;
        ws.Cell(3, 3).Value = 3.14159; ws.Cell(3, 3).Style.NumberFormat.Format = "0.00";

        ws.Cell(4, 1).Value = "Percentage";
        ws.Cell(4, 2).Value = 0.75;
        ws.Cell(4, 3).Value = 0.75; ws.Cell(4, 3).Style.NumberFormat.Format = "0.00%";

        ws.Cell(5, 1).Value = "Currency";
        ws.Cell(5, 2).Value = 1500.0;
        ws.Cell(5, 3).Value = 1500.0; ws.Cell(5, 3).Style.NumberFormat.Format = "$#,##0.00";

        ws.Cell(6, 1).Value = "Date";
        ws.Cell(6, 2).Value = new DateTime(2024, 6, 15);
        ws.Cell(6, 3).Value = new DateTime(2024, 6, 15);
        ws.Cell(6, 3).Style.NumberFormat.Format = "dd/MM/yyyy";

        ws.Cell(7, 1).Value = "Scientific";
        ws.Cell(7, 2).Value = 0.000012345;
        ws.Cell(7, 3).Value = 0.000012345; ws.Cell(7, 3).Style.NumberFormat.Format = "0.000E+00";

        ws.Cell(8, 1).Value = "Text";
        ws.Cell(8, 2).Value = "hello";
        ws.Cell(8, 3).Value = "world";

        ws.Cell(9, 1).Value = "Boolean";
        ws.Cell(9, 2).Value = true;
        ws.Cell(9, 3).Value = false;
    }

    /// Large sheet to stress-test the pipeline
    static void CreateLargeSheet(XLWorkbook wb)
    {
        var ws = wb.Worksheets.Add("Large");

        string[] cols = { "ID", "Name", "Dept", "Salary", "YearsExp", "Rating", "Active", "StartDate", "Email", "Notes" };
        for (int c = 0; c < cols.Length; c++)
        {
            ws.Cell(1, c + 1).Value = cols[c];
            ws.Cell(1, c + 1).Style.Font.Bold = true;
        }

        var depts = new[] { "Engineering", "Sales", "HR", "Finance", "Marketing" };
        var rng = new Random(7);
        for (int r = 2; r <= 51; r++)
        {
            ws.Cell(r, 1).Value = r - 1;
            ws.Cell(r, 2).Value = $"Employee{r - 1}";
            ws.Cell(r, 3).Value = depts[rng.Next(depts.Length)];
            ws.Cell(r, 4).Value = 40000 + rng.Next(60000);
            ws.Cell(r, 4).Style.NumberFormat.Format = "$#,##0";
            ws.Cell(r, 5).Value = rng.Next(1, 20);
            ws.Cell(r, 6).Value = Math.Round(3.0 + rng.NextDouble() * 2, 1);
            ws.Cell(r, 7).Value = rng.Next(2) == 0 ? "Yes" : "No";
            ws.Cell(r, 8).Value = new DateTime(2015 + rng.Next(9), rng.Next(1, 13), rng.Next(1, 28));
            ws.Cell(r, 8).Style.NumberFormat.Format = "yyyy-mm-dd";
            ws.Cell(r, 9).Value = $"emp{r - 1}@company.com";
            ws.Cell(r, 10).Value = rng.Next(2) == 0 ? "Active" : "";
        }
    }

    /// Sparse sheet — many empty cells, tests sparsity filter
    static void CreateSparseCells(XLWorkbook wb)
    {
        var ws = wb.Worksheets.Add("Sparse");

        // A small dense region
        ws.Cell(1, 1).Value = "Key";   ws.Cell(1, 1).Style.Font.Bold = true;
        ws.Cell(1, 2).Value = "Value"; ws.Cell(1, 2).Style.Font.Bold = true;
        ws.Cell(2, 1).Value = "Alpha"; ws.Cell(2, 2).Value = 100;
        ws.Cell(3, 1).Value = "Beta";  ws.Cell(3, 2).Value = 200;
        ws.Cell(4, 1).Value = "Gamma"; ws.Cell(4, 2).Value = 300;

        // Scattered isolated cells far away
        ws.Cell(10, 8).Value = "Note";
        ws.Cell(15, 12).Value = 999;
    }

    /// Multiple worksheets in one workbook
    static void CreateMultiSheet(XLWorkbook wb)
    {
        // Sheet 1 — same as simple table
        CreateSimpleTable(wb);

        // Sheet 2 — date/currency
        CreateDatesCurrency(wb);

        // Sheet 3 — summary sheet
        var ws = wb.Worksheets.Add("Summary");
        ws.Cell(1, 1).Value = "Summary"; ws.Cell(1, 1).Style.Font.Bold = true;
        ws.Cell(2, 1).Value = "Total Sales";
        ws.Cell(2, 2).FormulaA1 = "SUM(Sales!D2:D6)";
        ws.Cell(3, 1).Value = "Generated";
        ws.Cell(3, 2).Value = DateTime.Today;
        ws.Cell(3, 2).Style.NumberFormat.Format = "yyyy-mm-dd";
    }

    /// 500-row sales transaction log — high repetition in Region, Product, Status columns.
    /// This is the realistic scenario SpreadsheetLLM is designed for.
    static void CreateRealisticSalesSheet(XLWorkbook wb)
    {
        var ws = wb.Worksheets.Add("SalesTransactions");
        var rng = new Random(42);

        // Header
        string[] headers = { "TxnID", "Date", "Region", "Salesperson", "Product", "Category", "Qty", "UnitPrice", "Total", "Status", "PaymentMethod" };
        for (int c = 0; c < headers.Length; c++)
        {
            ws.Cell(1, c + 1).Value = headers[c];
            ws.Cell(1, c + 1).Style.Font.Bold = true;
            ws.Cell(1, c + 1).Style.Fill.BackgroundColor = XLColor.LightBlue;
        }

        // Small pools of values → lots of repetition in inverted index
        var regions      = new[] { "North", "South", "East", "West", "Central" };
        var products     = new[] { "Widget A", "Widget B", "Gadget X", "Gadget Y", "Part Z" };
        var categories   = new[] { "Hardware", "Software", "Services", "Accessories" };
        var statuses     = new[] { "Completed", "Completed", "Completed", "Pending", "Cancelled" }; // weighted
        var payments     = new[] { "Credit Card", "Bank Transfer", "Cash", "Credit Card", "Credit Card" }; // weighted
        var salespeople  = new[] { "Alice Smith", "Bob Jones", "Carol White", "David Brown", "Eva Green",
                                   "Frank Hall",  "Grace Lee",  "Henry King",  "Iris Chen",  "Jack Ford" };
        var unitPrices   = new[] { 9.99, 14.99, 24.99, 49.99, 99.99 };

        var startDate = new DateTime(2023, 1, 1);

        for (int r = 0; r < 500; r++)
        {
            int row = r + 2;
            var product    = products[rng.Next(products.Length)];
            var unitPrice  = unitPrices[Array.IndexOf(products, product)];
            var qty        = rng.Next(1, 50);
            var total      = Math.Round(qty * unitPrice, 2);
            var txDate     = startDate.AddDays(rng.Next(365));

            ws.Cell(row, 1).Value  = r + 1001;                                        // TxnID
            ws.Cell(row, 2).Value  = txDate;
            ws.Cell(row, 2).Style.NumberFormat.Format = "yyyy-mm-dd";
            ws.Cell(row, 3).Value  = regions[rng.Next(regions.Length)];               // Region
            ws.Cell(row, 4).Value  = salespeople[rng.Next(salespeople.Length)];       // Salesperson
            ws.Cell(row, 5).Value  = product;                                         // Product
            ws.Cell(row, 6).Value  = categories[rng.Next(categories.Length)];         // Category
            ws.Cell(row, 7).Value  = qty;                                             // Qty
            ws.Cell(row, 8).Value  = unitPrice;
            ws.Cell(row, 8).Style.NumberFormat.Format = "$#,##0.00";
            ws.Cell(row, 9).Value  = total;
            ws.Cell(row, 9).Style.NumberFormat.Format = "$#,##0.00";
            ws.Cell(row, 10).Value = statuses[rng.Next(statuses.Length)];             // Status
            ws.Cell(row, 11).Value = payments[rng.Next(payments.Length)];             // PaymentMethod
        }

        // Summary row
        int sumRow = 503;
        ws.Cell(sumRow, 1).Value = "TOTAL";
        ws.Cell(sumRow, 1).Style.Font.Bold = true;
        ws.Cell(sumRow, 9).FormulaA1 = "SUM(I2:I501)";
        ws.Cell(sumRow, 9).Style.NumberFormat.Format = "$#,##0.00";
        ws.Cell(sumRow, 9).Style.Font.Bold = true;
    }

    /// 300-row HR payroll sheet — heavily repeated dept, job title, pay grade, location.
    static void CreateRealisticHRSheet(XLWorkbook wb)
    {
        var ws = wb.Worksheets.Add("Payroll");
        var rng = new Random(99);

        string[] headers = { "EmpID", "Name", "Department", "JobTitle", "PayGrade", "Location", "StartDate", "BaseSalary", "Bonus", "TotalComp", "Status", "Manager" };
        for (int c = 0; c < headers.Length; c++)
        {
            ws.Cell(1, c + 1).Value = headers[c];
            ws.Cell(1, c + 1).Style.Font.Bold = true;
            ws.Cell(1, c + 1).Style.Fill.BackgroundColor = XLColor.LightGreen;
        }

        var depts      = new[] { "Engineering", "Engineering", "Engineering", "Sales", "Sales",
                                  "Finance", "HR", "Marketing", "Operations", "Legal" };
        var titles     = new[] { "Software Engineer", "Senior Engineer", "Engineering Manager",
                                  "Sales Rep", "Account Executive",
                                  "Financial Analyst", "HR Specialist", "Marketing Manager",
                                  "Operations Lead", "Legal Counsel" };
        var grades     = new[] { "L3", "L4", "L5", "L3", "L4", "L4", "L3", "L5", "L4", "L5" };
        var locations  = new[] { "New York", "New York", "San Francisco", "Chicago", "Boston",
                                  "New York", "Chicago", "San Francisco", "Austin", "New York" };
        var managers   = new[] { "John Adams", "Sarah Lee", "Mike Brown", "Lisa Chan", "Tom White" };
        var statuses   = new[] { "Active", "Active", "Active", "Active", "On Leave" };
        var baseSals   = new[] { 85000, 115000, 145000, 75000, 95000, 90000, 72000, 105000, 88000, 130000 };

        var startBase  = new DateTime(2015, 1, 1);

        for (int r = 0; r < 300; r++)
        {
            int row   = r + 2;
            int idx   = rng.Next(depts.Length);
            var sal   = baseSals[idx] + rng.Next(-5000, 15000);
            var bonus = (int)(sal * (0.05 + rng.NextDouble() * 0.15));
            var start = startBase.AddDays(rng.Next(3000));

            ws.Cell(row, 1).Value  = 2000 + r;
            ws.Cell(row, 2).Value  = $"Employee {2000 + r}";
            ws.Cell(row, 3).Value  = depts[idx];
            ws.Cell(row, 4).Value  = titles[idx];
            ws.Cell(row, 5).Value  = grades[idx];
            ws.Cell(row, 6).Value  = locations[idx];
            ws.Cell(row, 7).Value  = start;
            ws.Cell(row, 7).Style.NumberFormat.Format = "yyyy-mm-dd";
            ws.Cell(row, 8).Value  = sal;
            ws.Cell(row, 8).Style.NumberFormat.Format = "$#,##0";
            ws.Cell(row, 9).Value  = bonus;
            ws.Cell(row, 9).Style.NumberFormat.Format = "$#,##0";
            ws.Cell(row, 10).FormulaA1 = $"H{row}+I{row}";
            ws.Cell(row, 10).Style.NumberFormat.Format = "$#,##0";
            ws.Cell(row, 11).Value = statuses[rng.Next(statuses.Length)];
            ws.Cell(row, 12).Value = managers[rng.Next(managers.Length)];
        }
    }

    // -------------------------------------------------------------------------
    // Task 3 — plain_adjacent.xlsx recreated in .NET
    // Two tables placed directly adjacent (no gap row) — the same structure as
    // the plain_adjacent.xlsx sample file that ships with the project.
    // -------------------------------------------------------------------------

    /// Recreates plain_adjacent.xlsx: two tables back-to-back with no gap row.
    /// Table 1 (rows 1-6): Fruit price/stock inventory with formulas.
    /// Table 2 (rows 7-12): Student scores with averages.
    static void CreatePlainAdjacent(XLWorkbook wb)
    {
        var ws = wb.Worksheets.Add("Sheet1");

        // ---- Table 1: Fruit inventory (rows 1-6) ----
        ws.Cell(1, 1).Value = "Fruit";       ws.Cell(1, 1).Style.Font.Bold = true;
        ws.Cell(1, 2).Value = "Price";       ws.Cell(1, 2).Style.Font.Bold = true;
        ws.Cell(1, 3).Value = "Stock";       ws.Cell(1, 3).Style.Font.Bold = true;
        ws.Cell(1, 4).Value = "Total Value"; ws.Cell(1, 4).Style.Font.Bold = true;

        var fruits = new (string name, double price, int stock)[]
        {
            ("Apple",  1.50, 200),
            ("Banana", 0.75, 350),
            ("Cherry", 3.00, 100),
            ("Mango",  2.25, 150),
        };
        for (int i = 0; i < fruits.Length; i++)
        {
            int row = i + 2;
            ws.Cell(row, 1).Value = fruits[i].name;
            ws.Cell(row, 2).Value = fruits[i].price;
            ws.Cell(row, 3).Value = fruits[i].stock;
            ws.Cell(row, 4).FormulaA1 = $"B{row}*C{row}";
        }
        // Summary row (row 6) — no gap, immediately follows data
        ws.Cell(6, 1).Value = "Total";
        ws.Cell(6, 2).FormulaA1 = "AVERAGE(B2:B5)";
        ws.Cell(6, 3).FormulaA1 = "SUM(C2:C5)";
        ws.Cell(6, 4).FormulaA1 = "SUM(D2:D5)";

        // ---- Table 2: Student scores (rows 7-12) — no gap row ----
        ws.Cell(7, 1).Value = "Student"; ws.Cell(7, 1).Style.Font.Bold = true;
        ws.Cell(7, 2).Value = "Math";    ws.Cell(7, 2).Style.Font.Bold = true;
        ws.Cell(7, 3).Value = "Science"; ws.Cell(7, 3).Style.Font.Bold = true;
        ws.Cell(7, 4).Value = "Average"; ws.Cell(7, 4).Style.Font.Bold = true;

        var students = new (string name, int math, int science)[]
        {
            ("John", 85, 92),
            ("Sara", 78, 88),
            ("Mike", 92, 76),
            ("Lisa", 95, 98),
        };
        for (int i = 0; i < students.Length; i++)
        {
            int row = i + 8;
            ws.Cell(row, 1).Value = students[i].name;
            ws.Cell(row, 2).Value = students[i].math;
            ws.Cell(row, 3).Value = students[i].science;
            ws.Cell(row, 4).FormulaA1 = $"AVERAGE(B{row}:C{row})";
        }
        // Summary row (row 12)
        ws.Cell(12, 1).Value = "Class Avg";
        ws.Cell(12, 2).FormulaA1 = "AVERAGE(B8:B11)";
        ws.Cell(12, 3).FormulaA1 = "AVERAGE(C8:C11)";
        ws.Cell(12, 4).FormulaA1 = "AVERAGE(D8:D11)";
    }

    // -------------------------------------------------------------------------
    // Task 2 — Messy 2000-row sheet
    // Tests how the pipeline handles: large data, missing cells, sub-headers
    // inserted mid-table, mixed numeric/text in the same column, and sparse areas.
    // -------------------------------------------------------------------------

    /// 2000-row "messy" data-entry sheet with irregular structure.
    /// Columns: RecordID, Date, Category, SubCategory, Amount, Notes, Status, Flag
    /// Every 250 rows a sub-header summary row is inserted (tests anchor detection).
    /// ~20% of cells are intentionally empty. Amount column has mixed types.
    static void CreateMessy2000(XLWorkbook wb)
    {
        var ws = wb.Worksheets.Add("Messy2000");
        var rng = new Random(13);

        string[] headers = { "RecordID", "Date", "Category", "SubCategory", "Amount", "Notes", "Status", "Flag" };
        for (int c = 0; c < headers.Length; c++)
        {
            ws.Cell(1, c + 1).Value = headers[c];
            ws.Cell(1, c + 1).Style.Font.Bold = true;
            ws.Cell(1, c + 1).Style.Fill.BackgroundColor = XLColor.LightYellow;
        }

        var categories    = new[] { "Revenue", "Expense", "Transfer", "Adjustment", "Refund" };
        var subCategories = new[] { "Direct", "Indirect", "Internal", "External", "Capital" };
        var statuses      = new[] { "Approved", "Approved", "Approved", "Pending", "Rejected", "On Hold" };
        var flags         = new[] { "Y", "N", "Y", "Y", "" }; // weighted, some empty

        var startDate = new DateTime(2020, 1, 1);
        int dataRow   = 2;    // tracks actual Excel row number
        int recordId  = 1;

        for (int batch = 0; batch < 8; batch++)
        {
            // Insert a section sub-header every 250 records (tests mid-table boundary detection)
            if (batch > 0)
            {
                ws.Cell(dataRow, 1).Value = $"--- Section {batch + 1} ---";
                ws.Cell(dataRow, 1).Style.Font.Bold = true;
                ws.Cell(dataRow, 1).Style.Font.Italic = true;
                ws.Cell(dataRow, 1).Style.Fill.BackgroundColor = XLColor.LightCyan;
                dataRow++;
            }

            for (int i = 0; i < 250; i++)
            {
                int catIdx = rng.Next(categories.Length);
                bool skipDate   = rng.Next(5) == 0;   // ~20% missing
                bool skipSub    = rng.Next(5) == 0;
                bool skipNotes  = rng.Next(3) == 0;   // ~33% missing notes
                bool skipFlag   = rng.Next(4) == 0;

                ws.Cell(dataRow, 1).Value = recordId++;

                if (!skipDate)
                {
                    ws.Cell(dataRow, 2).Value = startDate.AddDays(rng.Next(1460));
                    ws.Cell(dataRow, 2).Style.NumberFormat.Format = "yyyy-mm-dd";
                }

                ws.Cell(dataRow, 3).Value = categories[catIdx];

                if (!skipSub)
                    ws.Cell(dataRow, 4).Value = subCategories[rng.Next(subCategories.Length)];

                // Amount: mostly numeric, but ~5% are text like "N/A" or "TBD" (messy!)
                if (rng.Next(20) == 0)
                    ws.Cell(dataRow, 5).Value = rng.Next(2) == 0 ? "N/A" : "TBD";
                else
                {
                    ws.Cell(dataRow, 5).Value = Math.Round(rng.NextDouble() * 50000, 2);
                    ws.Cell(dataRow, 5).Style.NumberFormat.Format = "$#,##0.00";
                }

                if (!skipNotes)
                    ws.Cell(dataRow, 6).Value = $"Note-{rng.Next(1, 500)}";

                ws.Cell(dataRow, 7).Value = statuses[rng.Next(statuses.Length)];

                if (!skipFlag)
                    ws.Cell(dataRow, 8).Value = flags[rng.Next(flags.Length)];

                dataRow++;
            }
        }

        // Footer summary row
        ws.Cell(dataRow, 1).Value = "TOTAL RECORDS";
        ws.Cell(dataRow, 1).Style.Font.Bold = true;
        ws.Cell(dataRow, 2).Value = recordId - 1;
        ws.Cell(dataRow, 2).Style.Font.Bold = true;
    }

    // -------------------------------------------------------------------------
    // Task 4 — Large adjacent report (3 tables, ~500 rows, NO gap rows)
    // The three tables sit back-to-back in the same sheet with no empty row
    // between them. This is a realistic layout for management reports.
    //
    // Table 1 (rows   1-151): Product Catalog  — 1 header + 150 products
    // Table 2 (rows 152-251): Sales Rep Quotas — 1 header + 100 reps
    // Table 3 (rows 252-501): Customer Orders  — 1 header + 250 orders
    // Total: 503 rows (≈ 500)
    // -------------------------------------------------------------------------

    /// Large sheet with 3 back-to-back tables and no gap rows between them.
    static void CreateLargeAdjacentReport(XLWorkbook wb)
    {
        var ws  = wb.Worksheets.Add("AdjacentReport");
        var rng = new Random(55);

        // ---- Table 1: Product Catalog (rows 1-151) ----
        string[] prodHeaders = { "ProductID", "ProductName", "Category", "UnitPrice", "StockQty", "ReorderPoint" };
        for (int c = 0; c < prodHeaders.Length; c++)
        {
            ws.Cell(1, c + 1).Value = prodHeaders[c];
            ws.Cell(1, c + 1).Style.Font.Bold = true;
            ws.Cell(1, c + 1).Style.Fill.BackgroundColor = XLColor.LightBlue;
        }

        var prodCategories = new[] { "Electronics", "Apparel", "Home", "Sports", "Books" };
        var prodNames      = new[] { "Widget", "Gadget", "Gizmo", "Doohickey", "Thingamajig",
                                     "Contraption", "Device", "Module", "Component", "Unit" };

        for (int i = 0; i < 150; i++)
        {
            int row = i + 2;
            int cat = rng.Next(prodCategories.Length);
            ws.Cell(row, 1).Value = 1000 + i;
            ws.Cell(row, 2).Value = $"{prodNames[rng.Next(prodNames.Length)]}-{rng.Next(100, 999)}";
            ws.Cell(row, 3).Value = prodCategories[cat];
            ws.Cell(row, 4).Value = Math.Round(5.0 + rng.NextDouble() * 495, 2);
            ws.Cell(row, 4).Style.NumberFormat.Format = "$#,##0.00";
            ws.Cell(row, 5).Value = rng.Next(0, 500);
            ws.Cell(row, 6).Value = rng.Next(10, 100);
        }

        // ---- Table 2: Sales Rep Quotas (rows 152-251 — no gap) ----
        string[] repHeaders = { "RepID", "RepName", "Territory", "AnnualQuota", "Q1Target", "Q2Target", "Q3Target", "Q4Target" };
        for (int c = 0; c < repHeaders.Length; c++)
        {
            ws.Cell(152, c + 1).Value = repHeaders[c];
            ws.Cell(152, c + 1).Style.Font.Bold = true;
            ws.Cell(152, c + 1).Style.Fill.BackgroundColor = XLColor.LightGreen;
        }

        var territories = new[] { "North", "South", "East", "West", "Central", "Northeast", "Southwest" };
        var repFirstNames = new[] { "Alice", "Bob", "Carol", "David", "Eva", "Frank", "Grace", "Henry" };
        var repLastNames  = new[] { "Smith", "Jones", "White", "Brown", "Green", "Hall", "Lee", "King" };

        for (int i = 0; i < 100; i++)
        {
            int row = i + 153;
            int quota = (rng.Next(5, 20)) * 10000;
            ws.Cell(row, 1).Value = 2000 + i;
            ws.Cell(row, 2).Value = $"{repFirstNames[rng.Next(repFirstNames.Length)]} {repLastNames[rng.Next(repLastNames.Length)]}";
            ws.Cell(row, 3).Value = territories[rng.Next(territories.Length)];
            ws.Cell(row, 4).Value = quota;
            ws.Cell(row, 4).Style.NumberFormat.Format = "$#,##0";
            ws.Cell(row, 5).Value = (int)(quota * 0.25);
            ws.Cell(row, 5).Style.NumberFormat.Format = "$#,##0";
            ws.Cell(row, 6).Value = (int)(quota * 0.25);
            ws.Cell(row, 6).Style.NumberFormat.Format = "$#,##0";
            ws.Cell(row, 7).Value = (int)(quota * 0.25);
            ws.Cell(row, 7).Style.NumberFormat.Format = "$#,##0";
            ws.Cell(row, 8).Value = (int)(quota * 0.25);
            ws.Cell(row, 8).Style.NumberFormat.Format = "$#,##0";
        }

        // ---- Table 3: Customer Orders (rows 252-501 — no gap) ----
        string[] orderHeaders = { "OrderID", "OrderDate", "CustomerName", "Product", "Qty", "UnitPrice", "Total", "Status" };
        for (int c = 0; c < orderHeaders.Length; c++)
        {
            ws.Cell(252, c + 1).Value = orderHeaders[c];
            ws.Cell(252, c + 1).Style.Font.Bold = true;
            ws.Cell(252, c + 1).Style.Fill.BackgroundColor = XLColor.LightSalmon;
        }

        var customers     = new[] { "Acme Corp", "Globex", "Initech", "Umbrella Ltd", "Stark Industries",
                                    "Wayne Enterprises", "Cyberdyne", "Oscorp", "Massive Dynamic", "Hooli" };
        var orderStatuses = new[] { "Shipped", "Shipped", "Shipped", "Processing", "Pending", "Cancelled" };
        var orderDate     = new DateTime(2024, 1, 1);

        for (int i = 0; i < 250; i++)
        {
            int row       = i + 253;
            int prodIdx   = rng.Next(150);          // reference a product from Table 1
            double price  = Math.Round(5.0 + rng.NextDouble() * 495, 2);
            int qty       = rng.Next(1, 100);
            double total  = Math.Round(price * qty, 2);

            ws.Cell(row, 1).Value = 5000 + i;
            ws.Cell(row, 2).Value = orderDate.AddDays(rng.Next(365));
            ws.Cell(row, 2).Style.NumberFormat.Format = "yyyy-mm-dd";
            ws.Cell(row, 3).Value = customers[rng.Next(customers.Length)];
            ws.Cell(row, 4).Value = $"Product-{1000 + prodIdx}";
            ws.Cell(row, 5).Value = qty;
            ws.Cell(row, 6).Value = price;
            ws.Cell(row, 6).Style.NumberFormat.Format = "$#,##0.00";
            ws.Cell(row, 7).Value = total;
            ws.Cell(row, 7).Style.NumberFormat.Format = "$#,##0.00";
            ws.Cell(row, 8).Value = orderStatuses[rng.Next(orderStatuses.Length)];
        }
    }
}
