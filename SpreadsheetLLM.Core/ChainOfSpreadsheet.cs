using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using SpreadsheetLLM.Core.Models;

namespace SpreadsheetLLM.Core
{
    /// <summary>
    /// Chain of Spreadsheet (CoS) — QA pipeline for SpreadsheetLLM.
    /// Port of spreadsheet_llm/cos.py (Paper Section 4).
    ///
    /// Stage 1: Table identification from compressed encoding (Section 4.1)
    /// Stage 2: Response generation from uncompressed table (Section 4.2)
    /// Table splitting for large tables (Appendix M.2, Algorithm 2)
    ///
    /// LLM backend: Anthropic (Claude) or OpenAI (GPT-4) via REST, or placeholder for demo.
    /// </summary>
    public sealed class ChainOfSpreadsheet
    {
        // Prompt templates from Appendix L.3
        private const string QaStage1Prompt = @"INSTRUCTION:
Given an input that is a string denoting data of cells in a table. The input table contains many tuples, describing the cells with content in the spreadsheet. Each tuple consists of two elements separated by a '|': the cell content and the cell address/region, like (Year|A1), ( |A1) or (IntNum|A1:B3). The content in some cells such as '#,##0'/'d-mmm-yy'/'H:mm:ss', etc., represents the CELL DATA FORMATS of Excel. The content in some cells such as 'IntNum'/'DateData'/'EmailData', etc., represents a category of data with the same format and similar semantics. For example, 'IntNum' represents integer type data, and 'ScientificNum' represents scientific notation type data. 'A1:B3' represents a region in a spreadsheet, from the first row to the third row and from column A to column B. Some cells with empty content in the spreadsheet are not entered. How many tables are there in the spreadsheet? Below is a question about one certain table in this spreadsheet. I need you to determine in which table the answer to the following question can be found, and return the RANGE of the ONE table you choose, LIKE ['range': 'A1:F9']. DON'T ADD OTHER WORDS OR EXPLANATION.

INPUT:
{encoded_sheet}

QUESTION:
{question}";

        private const string QaStage2Prompt = @"INSTRUCTION:
Given an input that is a string denoting data of cells in a table and a question about this table. The answer to the question can be found in the table. The input table includes many pairs, and each pair consists of a cell address and the text in that cell with a ',' in between, like 'A1,Year'. Cells are separated by '|' like 'A1,Year|A2,Profit'. The text can be empty so the cell data is like 'A1, |A2,Profit'. The cells are organized in row-major order. The answer to the input question is contained in the input table and can be represented by cell address. I need you to find the cell address of the answer in the given table based on the given question description, and return the cell ADDRESS of the answer like '[B3]' or '[SUM(A2:A10)]'. DON'T ADD ANY OTHER WORDS.

INPUT:
{encoded_table}

QUESTION:
{question}";

        private readonly HttpClient _httpClient;

        public ChainOfSpreadsheet(HttpClient? httpClient = null)
        {
            _httpClient = httpClient ?? new HttpClient();
        }

        // =====================================================================
        // Public API
        // =====================================================================

        /// <summary>
        /// CoS Stage 1: Identify the relevant table range for a query.
        /// Returns a range string like "A1:F9", or null.
        /// </summary>
        public async Task<string?> IdentifyTableAsync(
            SpreadsheetEncoding encoding, string query, string? provider = null)
        {
            provider ??= DetectProvider();

            var sheetName = FindRelevantSheet(encoding, query);
            if (sheetName == null) return null;

            var sheetData = encoding.Sheets[sheetName];
            var encoded = JsonSerializer.Serialize(sheetData);
            var prompt = QaStage1Prompt
                .Replace("{encoded_sheet}", encoded)
                .Replace("{question}", query);

            var response = await CallLlmAsync(prompt, provider);

            var match = Regex.Match(response, @"'([A-Z]+\d+:[A-Z]+\d+)'");
            return match.Success ? match.Groups[1].Value : null;
        }

        /// <summary>
        /// CoS Stage 2: Generate an answer from the identified table.
        /// Returns a cell address like "[B3]".
        /// </summary>
        public async Task<string> GenerateResponseAsync(
            SheetEncoding tableData, string query, string? provider = null)
        {
            provider ??= DetectProvider();

            var encoded = JsonSerializer.Serialize(tableData);
            var prompt = QaStage2Prompt
                .Replace("{encoded_table}", encoded)
                .Replace("{question}", query);

            return await CallLlmAsync(prompt, provider);
        }

        /// <summary>
        /// Handle QA for large tables by splitting into chunks (Algorithm 2, Appendix M.2).
        /// </summary>
        public async Task<string> TableSplitQaAsync(
            SheetEncoding sheetData, string tableRange, string query,
            int tokenLimit = 4096, string? provider = null)
        {
            provider ??= DetectProvider();

            int tableTokens = JsonSerializer.Serialize(sheetData).Length;
            if (tableTokens <= tokenLimit)
                return await GenerateResponseAsync(sheetData, query, provider);

            var allItems = sheetData.Cells.ToList();
            if (allItems.Count == 0)
                return await GenerateResponseAsync(sheetData, query, provider);

            int headerSize = Math.Max(1, allItems.Count / 10);
            var headerItems = allItems.Take(headerSize).ToList();
            var bodyItems = allItems.Skip(headerSize).ToList();

            var headerData = new SheetEncoding();
            foreach (var kv in headerItems) headerData.Cells[kv.Key] = kv.Value;
            int headerTokens = JsonSerializer.Serialize(headerData).Length;
            int chunkBudget = tokenLimit - headerTokens;

            if (chunkBudget <= 0)
                return await GenerateResponseAsync(sheetData, query, provider);

            var answers = new List<string>();
            var chunk = new List<KeyValuePair<string, List<string>>>();
            int chunkTokens = 0;

            foreach (var item in bodyItems)
            {
                int itemTokens = JsonSerializer.Serialize(new Dictionary<string, List<string>> { [item.Key] = item.Value }).Length;

                if (chunkTokens + itemTokens > chunkBudget && chunk.Count > 0)
                {
                    var chunkData = BuildChunkEncoding(sheetData, headerItems, chunk);
                    answers.Add(await GenerateResponseAsync(chunkData, query, provider));
                    chunk.Clear();
                    chunkTokens = 0;
                }

                chunk.Add(item);
                chunkTokens += itemTokens;
            }

            if (chunk.Count > 0)
            {
                var chunkData = BuildChunkEncoding(sheetData, headerItems, chunk);
                answers.Add(await GenerateResponseAsync(chunkData, query, provider));
            }

            if (answers.Count == 1) return answers[0];
            return $"Aggregated from {answers.Count} chunks: " + string.Join(" | ", answers);
        }

        // =====================================================================
        // LLM Backends
        // =====================================================================

        private async Task<string> CallLlmAsync(string prompt, string provider)
        {
            return provider switch
            {
                "anthropic" => await CallAnthropicAsync(prompt),
                "openai" => await CallOpenAiAsync(prompt),
                _ => CallPlaceholder(prompt),
            };
        }

        private async Task<string> CallAnthropicAsync(string prompt)
        {
            var apiKey = Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY");
            if (string.IsNullOrEmpty(apiKey)) return "";

            var body = JsonSerializer.Serialize(new
            {
                model = "claude-sonnet-4-20250514",
                max_tokens = 512,
                messages = new[] { new { role = "user", content = prompt } }
            });

            using var request = new HttpRequestMessage(HttpMethod.Post,
                "https://api.anthropic.com/v1/messages");
            request.Headers.Add("x-api-key", apiKey);
            request.Headers.Add("anthropic-version", "2023-06-01");
            request.Content = new StringContent(body, Encoding.UTF8, "application/json");

            var response = await _httpClient.SendAsync(request);
            var json = await response.Content.ReadAsStringAsync();

            using var doc = JsonDocument.Parse(json);
            return doc.RootElement
                .GetProperty("content")[0]
                .GetProperty("text")
                .GetString() ?? "";
        }

        private async Task<string> CallOpenAiAsync(string prompt)
        {
            var apiKey = Environment.GetEnvironmentVariable("OPENAI_API_KEY");
            if (string.IsNullOrEmpty(apiKey)) return "";

            var body = JsonSerializer.Serialize(new
            {
                model = "gpt-4",
                max_tokens = 512,
                messages = new[] { new { role = "user", content = prompt } }
            });

            using var request = new HttpRequestMessage(HttpMethod.Post,
                "https://api.openai.com/v1/chat/completions");
            request.Headers.Add("Authorization", $"Bearer {apiKey}");
            request.Content = new StringContent(body, Encoding.UTF8, "application/json");

            var response = await _httpClient.SendAsync(request);
            var json = await response.Content.ReadAsStringAsync();

            using var doc = JsonDocument.Parse(json);
            return doc.RootElement
                .GetProperty("choices")[0]
                .GetProperty("message")
                .GetProperty("content")
                .GetString() ?? "";
        }

        private static string CallPlaceholder(string prompt)
        {
            if (prompt.Contains("determine in which table"))
                return "['range': 'A1:F9']";
            if (prompt.Contains("find the cell address"))
                return "[B3]";
            return "[placeholder response]";
        }

        private static string DetectProvider()
        {
            if (!string.IsNullOrEmpty(Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY")))
                return "anthropic";
            if (!string.IsNullOrEmpty(Environment.GetEnvironmentVariable("OPENAI_API_KEY")))
                return "openai";
            return "placeholder";
        }

        // =====================================================================
        // Helpers
        // =====================================================================

        private static string? FindRelevantSheet(SpreadsheetEncoding encoding, string query)
        {
            var queryTokens = query.ToLowerInvariant().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            int bestScore = 0;
            string? bestSheet = null;

            foreach (var kv in encoding.Sheets)
            {
                int score = 0;
                foreach (var cellValue in kv.Value.Cells.Keys)
                {
                    var lower = cellValue.ToLowerInvariant();
                    if (queryTokens.Any(t => lower.Contains(t)))
                        score++;
                }
                if (score > bestScore)
                {
                    bestScore = score;
                    bestSheet = kv.Key;
                }
            }

            // Fallback: first sheet
            return bestSheet ?? encoding.Sheets.Keys.FirstOrDefault();
        }

        private static SheetEncoding BuildChunkEncoding(
            SheetEncoding original,
            List<KeyValuePair<string, List<string>>> header,
            List<KeyValuePair<string, List<string>>> body)
        {
            var chunk = new SheetEncoding
            {
                StructuralAnchors = original.StructuralAnchors,
                Formats = original.Formats,
                NumericRanges = original.NumericRanges,
            };
            foreach (var kv in header) chunk.Cells[kv.Key] = kv.Value;
            foreach (var kv in body) chunk.Cells[kv.Key] = kv.Value;
            return chunk;
        }
    }
}
