using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace SpreadsheetLLM.Core.Models
{
    /// <summary>Top-level encoding result for an entire workbook.</summary>
    public sealed class SpreadsheetEncoding
    {
        [JsonPropertyName("file_name")]
        public string FileName { get; set; } = "";

        [JsonPropertyName("sheets")]
        public Dictionary<string, SheetEncoding> Sheets { get; set; } = new Dictionary<string, SheetEncoding>();

        [JsonPropertyName("compression_metrics")]
        public CompressionMetricsContainer CompressionMetrics { get; set; } = new CompressionMetricsContainer();
    }

    /// <summary>Encoding for a single worksheet.</summary>
    public sealed class SheetEncoding
    {
        [JsonPropertyName("structural_anchors")]
        public StructuralAnchors StructuralAnchors { get; set; } = new StructuralAnchors();

        /// <summary>Inverted index: cell value (or formula) → list of cell ranges.</summary>
        [JsonPropertyName("cells")]
        public Dictionary<string, List<string>> Cells { get; set; } = new Dictionary<string, List<string>>();

        /// <summary>Format aggregation: JSON type key → list of cell ranges.</summary>
        [JsonPropertyName("formats")]
        public Dictionary<string, List<string>> Formats { get; set; } = new Dictionary<string, List<string>>();

        /// <summary>Numeric-only sub-aggregation from formats.</summary>
        [JsonPropertyName("numeric_ranges")]
        public Dictionary<string, List<string>> NumericRanges { get; set; } = new Dictionary<string, List<string>>();
    }

    public sealed class StructuralAnchors
    {
        [JsonPropertyName("rows")]
        public List<int> Rows { get; set; } = new List<int>();

        [JsonPropertyName("columns")]
        public List<string> Columns { get; set; } = new List<string>();
    }
}
