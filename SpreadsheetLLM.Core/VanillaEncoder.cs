using System.Collections.Generic;
using System.Text;
using SpreadsheetLLM.Core.Models;

namespace SpreadsheetLLM.Core
{
    /// <summary>
    /// Vanilla baseline encoding: row-major, pipe-delimited, includes all non-empty cells.
    /// Port of spreadsheet_llm/vanilla.py (Paper Section 3.1).
    ///
    /// Format per row: "A1,value|B1,value|C1,value"
    /// Rows separated by newlines.
    /// </summary>
    public static class VanillaEncoder
    {
        /// <summary>Encode all worksheets. Returns sheet_name → encoded string.</summary>
        public static Dictionary<string, string> Encode(WorksheetSnapshot[] sheets)
        {
            var result = new Dictionary<string, string>();
            foreach (var sheet in sheets)
                result[sheet.Name] = EncodeSheet(sheet);
            return result;
        }

        private static string EncodeSheet(WorksheetSnapshot sheet)
        {
            var sb = new StringBuilder();
            bool firstRow = true;

            for (int r = 1; r <= sheet.MaxRow; r++)
            {
                var rowParts = new List<string>();

                for (int c = 1; c <= sheet.MaxColumn; c++)
                {
                    var cell = sheet.GetCell(r, c);
                    if (cell?.Value != null)
                    {
                        var coord = CellUtils.CellCoord(r, c);
                        rowParts.Add($"{coord},{cell.Value}");
                    }
                }

                if (rowParts.Count > 0)
                {
                    if (!firstRow) sb.Append('\n');
                    sb.Append(string.Join("|", rowParts));
                    firstRow = false;
                }
            }

            return sb.ToString();
        }
    }
}
