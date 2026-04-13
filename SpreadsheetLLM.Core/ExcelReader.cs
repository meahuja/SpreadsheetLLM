using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using SpreadsheetLLM.Core.Models;

namespace SpreadsheetLLM.Core
{
    /// <summary>
    /// Reads an Excel .xlsx file using ClosedXML and produces WorksheetSnapshot objects.
    /// Preserves formulas (FormulaA1) rather than computed values — mirrors openpyxl data_only=False.
    /// </summary>
    public static class ExcelReader
    {
        /// <summary>
        /// Load all worksheets from an xlsx file into WorksheetSnapshot arrays.
        /// </summary>
        public static WorksheetSnapshot[] ReadWorkbook(string path)
        {
            using var workbook = new XLWorkbook(path);
            var snapshots = new List<WorksheetSnapshot>(workbook.Worksheets.Count);

            foreach (var ws in workbook.Worksheets)
            {
                var snapshot = ReadWorksheet(ws);
                if (snapshot != null)
                    snapshots.Add(snapshot);
            }

            return snapshots.ToArray();
        }

        /// <summary>
        /// Convert a ClosedXML worksheet into a WorksheetSnapshot.
        /// </summary>
        private static WorksheetSnapshot? ReadWorksheet(IXLWorksheet ws)
        {
            if (ws.IsEmpty())
                return null;

            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
            var lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 0;

            if (lastRow == 0 || lastCol == 0)
                return null;

            // Build merged cell map: coord → start_coord
            var mergedCellMap = BuildMergedCellMap(ws);

            // Allocate cell grid
            var cells = new CellData?[lastRow][];
            for (int r = 0; r < lastRow; r++)
                cells[r] = new CellData?[lastCol];

            // Populate cells
            foreach (var row in ws.RowsUsed())
            {
                int r = row.RowNumber();
                if (r > lastRow) break;

                foreach (var cell in row.CellsUsed())
                {
                    int c = cell.Address.ColumnNumber;
                    if (c > lastCol) continue;

                    var coord = CellUtils.CellCoord(r, c);
                    string? mergeStart = mergedCellMap.TryGetValue(coord, out var ms) ? ms : null;

                    cells[r - 1][c - 1] = BuildCellData(cell, r, c, mergeStart);
                }
            }

            return new WorksheetSnapshot
            {
                Name = ws.Name,
                MaxRow = lastRow,
                MaxColumn = lastCol,
                Cells = cells,
                MergedCellMap = mergedCellMap,
            };
        }

        private static Dictionary<string, string> BuildMergedCellMap(IXLWorksheet ws)
        {
            var map = new Dictionary<string, string>();

            foreach (var mergedRange in ws.MergedRanges)
            {
                var startCoord = mergedRange.FirstCell().Address.ToString();

                foreach (var cell in mergedRange.Cells())
                {
                    var coord = cell.Address.ToString();
                    map[coord] = startCoord;
                }
            }

            return map;
        }

        private static CellData BuildCellData(IXLCell cell, int row, int col,
            string? mergeStartCoord)
        {
            string? value;
            string? computedValue = null;
            bool isFormula = cell.HasFormula;

            if (isFormula)
            {
                // Preserve the formula string (prepend "=" to match Python openpyxl behaviour)
                value = "=" + cell.FormulaA1;
                // Computed value is not reliably available from ClosedXML without calculation engine
                computedValue = null;
            }
            else
            {
                var rawValue = cell.Value;
                value = rawValue.IsBlank ? null : rawValue.ToString();
            }

            // Style extraction
            bool fontBold = false;
            bool fontItalic = false;
            string? fontColor = null;
            string? borderBottom = null;
            string? fillPattern = null;
            string? fillFg = null;
            string? alignH = null;

            try
            {
                var style = cell.Style;
                fontBold = style.Font.Bold;
                fontItalic = style.Font.Italic;
                // XLColor has no IsEmpty; just capture the string and let the
                // catch block handle any access failure.
                fontColor = style.Font.FontColor.ToString();
                borderBottom = style.Border.BottomBorder.ToString();
                fillPattern = style.Fill.PatternType.ToString();
                fillFg = style.Fill.PatternColor.ToString();
                alignH = style.Alignment.Horizontal.ToString();
            }
            catch
            {
                // Style access failure — leave defaults
            }

            return new CellData
            {
                Row = row,
                Col = col,
                Value = value,
                ComputedValue = computedValue,
                IsFormula = isFormula,
                IsMerged = mergeStartCoord != null,
                MergeStartCoord = mergeStartCoord ?? "",
                NumberFormat = cell.Style.NumberFormat.Format ?? "General",
                FontBold = fontBold,
                FontItalic = fontItalic,
                FontColor = fontColor,
                BorderBottomStyle = borderBottom,
                FillPatternType = fillPattern,
                FillFgColor = fillFg,
                AlignmentHorizontal = alignH,
            };
        }
    }
}
