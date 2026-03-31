using System.Collections.Generic;

namespace SpreadsheetLLM.Core.Models
{
    /// <summary>
    /// In-memory 2D grid of CellData for a single worksheet.
    /// Row/column indices are 1-based throughout the codebase (matching Excel conventions).
    /// </summary>
    public sealed class WorksheetSnapshot
    {
        public string Name { get; init; } = "";
        public int MaxRow { get; init; }
        public int MaxColumn { get; init; }

        /// <summary>
        /// Cells[row-1][col-1] → CellData or null (empty cell).
        /// Access via GetCell(row, col) for safe bounds-checked access.
        /// </summary>
        public CellData?[][] Cells { get; init; } = System.Array.Empty<CellData?[]>();

        /// <summary>
        /// Merged cell map: coordinate string → start coordinate of the merge range.
        /// E.g. "B2" → "A1" means B2 is part of a merge starting at A1.
        /// Built once by ExcelReader for O(1) lookups.
        /// </summary>
        public Dictionary<string, string> MergedCellMap { get; init; } = new Dictionary<string, string>();

        /// <summary>Returns the CellData at (row, col), or null if out of bounds or empty.</summary>
        public CellData? GetCell(int row, int col)
        {
            if (row < 1 || row > MaxRow || col < 1 || col > MaxColumn)
                return null;
            return Cells[row - 1][col - 1];
        }
    }
}
