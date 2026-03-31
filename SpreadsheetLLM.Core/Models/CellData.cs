namespace SpreadsheetLLM.Core.Models
{
    /// <summary>
    /// Immutable snapshot of a single cell's value, formula, and style.
    /// Populated by ExcelReader (ClosedXML) — decouples the algorithm from the reader source.
    /// </summary>
    public sealed class CellData
    {
        public int Row { get; init; }
        public int Col { get; init; }

        /// <summary>
        /// Formula string (e.g. "=SUM(B2:B5)") if IsFormula is true, otherwise the raw value string.
        /// Null when the cell is empty.
        /// </summary>
        public string? Value { get; init; }

        /// <summary>Cached computed value (may be null if unavailable or cell is empty).</summary>
        public string? ComputedValue { get; init; }

        public bool IsFormula { get; init; }
        public bool IsMerged { get; init; }

        /// <summary>Coordinate of the top-left cell of the merge range (e.g. "A1"). Empty when not merged.</summary>
        public string MergeStartCoord { get; init; } = "";

        public string NumberFormat { get; init; } = "General";

        // Style attributes — used for boundary detection heuristics
        public bool FontBold { get; init; }
        public bool FontItalic { get; init; }
        public string? FontColor { get; init; }
        public string? BorderBottomStyle { get; init; }
        public string? FillPatternType { get; init; }
        public string? FillFgColor { get; init; }
        public string? AlignmentHorizontal { get; init; }
    }
}
