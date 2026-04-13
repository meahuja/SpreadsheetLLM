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

        // Also encode the existing sample if present
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
}
