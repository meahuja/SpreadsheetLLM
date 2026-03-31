using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace SpreadsheetLLM.Core.Models
{
    public sealed class SheetMetrics
    {
        [JsonPropertyName("original_tokens")]
        public int OriginalTokens { get; set; }

        [JsonPropertyName("after_anchor_tokens")]
        public int AfterAnchorTokens { get; set; }

        [JsonPropertyName("after_inverted_index_tokens")]
        public int AfterInvertedIndexTokens { get; set; }

        [JsonPropertyName("after_format_tokens")]
        public int AfterFormatTokens { get; set; }

        [JsonPropertyName("final_tokens")]
        public int FinalTokens { get; set; }

        [JsonPropertyName("anchor_ratio")]
        public double AnchorRatio { get; set; }

        [JsonPropertyName("inverted_index_ratio")]
        public double InvertedIndexRatio { get; set; }

        [JsonPropertyName("format_ratio")]
        public double FormatRatio { get; set; }

        [JsonPropertyName("overall_ratio")]
        public double OverallRatio { get; set; }
    }

    public sealed class CompressionMetricsContainer
    {
        [JsonPropertyName("sheets")]
        public Dictionary<string, SheetMetrics> Sheets { get; set; } = new Dictionary<string, SheetMetrics>();

        [JsonPropertyName("overall")]
        public SheetMetrics Overall { get; set; } = new SheetMetrics();
    }
}
