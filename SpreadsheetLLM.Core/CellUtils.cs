using System;
using System.Text.RegularExpressions;
using SpreadsheetLLM.Core.Models;

namespace SpreadsheetLLM.Core
{
    /// <summary>
    /// Cell type detection, format parsing, and semantic type classification.
    /// Port of spreadsheet_llm/cell_utils.py.
    ///
    /// Implements the 9 semantic types from SpreadsheetLLM paper Section 3.4:
    /// Year, Integer, Float, Percentage, Scientific, Date, Time, Currency, Email.
    /// </summary>
    public static class CellUtils
    {
        // Pre-compiled regex patterns (static for performance — module-level equivalent)
        private static readonly Regex EmailRegex =
            new Regex(@"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$",
                RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly Regex CurrencyRegex =
            new Regex(@"[$€£¥]", RegexOptions.Compiled);

        private static readonly Regex ScientificRegex =
            new Regex(@"E[+-]", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly Regex DateKeywords =
            new Regex(@"(yyyy|yy|mmmm|mmm|dd|ddd|dddd)", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly Regex TimeKeywords =
            new Regex(@"(hh?|ss?|am/pm|a/p)", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly Regex YearOnlyRegex =
            new Regex(@"^y{2,4}$", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly Regex NumericNfsRegex =
            new Regex(@"^[#0,]+(\.[#0]+)?$", RegexOptions.Compiled);

        /// <summary>
        /// Infer the basic data type of a cell.
        /// Returns: "empty" | "text" | "numeric" | "boolean" | "datetime" | "email" | "error" | "formula"
        /// </summary>
        public static string InferCellDataType(CellData cell)
        {
            if (cell.Value == null)
                return "empty";

            // Check email before falling through to text
            if (!cell.IsFormula && EmailRegex.IsMatch(cell.Value))
                return "email";

            if (cell.IsFormula)
            {
                // For formula cells, infer from computed value
                if (cell.ComputedValue == null)
                    return "formula";
                if (double.TryParse(cell.ComputedValue, out _))
                    return "numeric";
                if (bool.TryParse(cell.ComputedValue, out _))
                    return "boolean";
                return "text";
            }

            // Try to infer from value string
            if (cell.Value == "TRUE" || cell.Value == "FALSE")
                return "boolean";
            if (double.TryParse(cell.Value, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out _))
                return "numeric";

            // Check datetime via number format
            var nfs = cell.NumberFormat.ToLowerInvariant();
            if (DateKeywords.IsMatch(nfs) || TimeKeywords.IsMatch(nfs))
                return "datetime";

            return "text";
        }

        /// <summary>Returns the raw Number Format String for a cell.</summary>
        public static string GetNumberFormatString(CellData cell)
        {
            var nfs = cell.NumberFormat;
            if (string.IsNullOrEmpty(nfs))
                return "General";
            return nfs;
        }

        /// <summary>
        /// Categorize a number format string into a broad category.
        /// Returns: general | currency | percentage | scientific | fraction |
        ///          date_custom | time_custom | datetime_custom | integer | float |
        ///          other_numeric | not_applicable | text_format
        /// </summary>
        public static string CategorizeNumberFormat(string nfs, CellData cell)
        {
            var cellType = InferCellDataType(cell);
            if (cellType != "numeric" && cellType != "datetime")
                return "not_applicable";

            if (string.IsNullOrEmpty(nfs) || nfs.Equals("General", StringComparison.OrdinalIgnoreCase))
                return cellType == "datetime" ? "datetime_general" : "general";

            if (nfs == "@" || nfs.Equals("text", StringComparison.OrdinalIgnoreCase))
                return "text_format";

            if (CurrencyRegex.IsMatch(nfs))
                return "currency";
            if (nfs.Contains("%"))
                return "percentage";
            if (ScientificRegex.IsMatch(nfs))
                return "scientific";
            if (nfs.Contains("#") && nfs.Contains("/") && nfs.Contains("?"))
                return "fraction";

            var nfsLower = nfs.ToLowerInvariant();
            bool isDate = DateKeywords.IsMatch(nfsLower);
            bool isTime = TimeKeywords.IsMatch(nfsLower);

            // Bare "m" disambiguation: month vs minute
            if (nfsLower.Contains("m") && !isDate && !isTime)
            {
                if (nfsLower.Contains("h") || nfsLower.Contains("s"))
                    isTime = true;
                else if (Regex.IsMatch(nfsLower, @"^m{1,5}$"))
                    isDate = true;
            }

            if (isDate && isTime) return "datetime_custom";
            if (isDate) return "date_custom";
            if (isTime) return "time_custom";

            if (cellType == "numeric")
            {
                if (nfs == "0" || nfs == "#,##0") return "integer";
                if (nfs == "0.00" || nfs == "#,##0.00" || nfs == "0.0" || nfs == "#,##0.0") return "float";
                if (NumericNfsRegex.IsMatch(nfs)) return "other_numeric";
            }

            if (cellType == "datetime")
                return "other_date";

            return "not_applicable";
        }

        /// <summary>
        /// Detect the semantic type of a cell per the paper's 9 categories.
        /// Returns: year | integer | float | percentage | scientific | date | time | currency | email | text | empty | boolean | error
        /// </summary>
        public static string DetectSemanticType(CellData cell)
        {
            var dataType = InferCellDataType(cell);

            if (dataType == "empty") return "empty";
            if (dataType == "email") return "email";
            if (dataType == "boolean") return "boolean";
            if (dataType == "error") return "error";

            var nfs = GetNumberFormatString(cell);
            var category = CategorizeNumberFormat(nfs, cell);
            var nfsLower = nfs.ToLowerInvariant();

            if (category == "percentage") return "percentage";
            if (category == "currency") return "currency";
            if (category == "scientific") return "scientific";

            if (category == "date_custom" || category == "datetime_custom" ||
                category == "datetime_general" || category == "other_date")
            {
                if (YearOnlyRegex.IsMatch(nfsLower))
                    return "year";
                return "date";
            }

            if (category == "time_custom") return "time";

            if (dataType == "numeric")
            {
                // Check if the value string represents an integer
                if (cell.Value != null && !cell.IsFormula)
                {
                    if (long.TryParse(cell.Value, out _))
                        return "integer";
                    if (category == "integer")
                        return "integer";
                    return "float";
                }
                if (category == "integer") return "integer";
                if (category == "float" || category == "other_numeric") return "float";
                return "integer"; // numeric fallback
            }

            return "text";
        }

        /// <summary>
        /// Returns a stable integer fingerprint for a cell's style.
        /// Used for boundary detection — two cells with the same style get the same fingerprint.
        /// Uses HashCode.Combine on actual attribute values (not object identity).
        /// </summary>
        public static int GetStyleFingerprint(CellData cell)
        {
            // Combine all style-relevant fields into a single hash
            int fontKey = CombineHash(
                cell.FontBold.GetHashCode(),
                cell.FontItalic.GetHashCode(),
                (cell.FontColor ?? "").GetHashCode()
            );
            int borderKey = (cell.BorderBottomStyle ?? "").GetHashCode();
            int fillKey = CombineHash(
                (cell.FillPatternType ?? "").GetHashCode(),
                (cell.FillFgColor ?? "").GetHashCode()
            );
            int alignKey = (cell.AlignmentHorizontal ?? "").GetHashCode();

            return CombineHash(fontKey, borderKey, fillKey, alignKey);
        }

        private static int CombineHash(int h1, int h2) =>
            unchecked(((h1 << 5) + h1) ^ h2);

        private static int CombineHash(int h1, int h2, int h3) =>
            CombineHash(CombineHash(h1, h2), h3);

        private static int CombineHash(int h1, int h2, int h3, int h4) =>
            CombineHash(CombineHash(h1, h2), CombineHash(h3, h4));

        /// <summary>Converts (row, col) 1-based integers to an Excel cell reference like "A1".</summary>
        public static string CellCoord(int row, int col)
        {
            return ColumnLetter(col) + row.ToString();
        }

        /// <summary>Converts a 1-based column number to a column letter string (e.g., 1 → "A", 27 → "AA").</summary>
        public static string ColumnLetter(int col)
        {
            var result = "";
            while (col > 0)
            {
                col--;
                result = (char)('A' + col % 26) + result;
                col /= 26;
            }
            return result;
        }

        /// <summary>Converts a column letter string to a 1-based column number (e.g., "A" → 1, "AA" → 27).</summary>
        public static int ColumnNumber(string colLetter)
        {
            int result = 0;
            foreach (var ch in colLetter.ToUpperInvariant())
            {
                result = result * 26 + (ch - 'A' + 1);
            }
            return result;
        }

        /// <summary>Splits a cell reference like "AB12" into ("AB", 12).</summary>
        public static (string colLetter, int row) SplitCellRef(string cellRef)
        {
            int i = 0;
            while (i < cellRef.Length && char.IsLetter(cellRef[i]))
                i++;
            var col = cellRef.Substring(0, i).ToUpperInvariant();
            var row = int.Parse(cellRef.Substring(i));
            return (col, row);
        }
    }
}
