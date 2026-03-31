using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using SpreadsheetLLM.Core.Models;

namespace SpreadsheetLLM.Core
{
    /// <summary>
    /// SheetCompressor — 3-stage spreadsheet compression pipeline.
    ///
    /// Port of spreadsheet_llm/encoder.py.
    ///
    /// Stage 1: Structural-Anchor-Based Extraction (Paper Section 3.2)
    /// Stage 2: Inverted-Index Translation (Paper Section 3.3)
    /// Stage 3: Data-Format-Aware Aggregation (Paper Section 3.4)
    /// </summary>
    public sealed class SheetCompressor
    {
        private const int MaxCandidates = 200;
        private const int MaxBoundaryRows = 100;
        private const double HeaderThreshold = 0.6;
        private const double SparsityThreshold = 0.10;
        private const double NmsIouThreshold = 0.5;
        private const int HeaderScoreWeight = 10;

        // =====================================================================
        // Public API
        // =====================================================================

        /// <summary>Encode an Excel file by path. Reads via ClosedXML, runs 3-stage pipeline.</summary>
        public SpreadsheetEncoding Encode(string filePath, int k = 2)
        {
            var sheets = ExcelReader.ReadWorkbook(filePath);
            return Encode(sheets, Path.GetFileName(filePath), k);
        }

        /// <summary>Encode pre-loaded WorksheetSnapshot array (e.g. from VSTO adapter).</summary>
        public SpreadsheetEncoding Encode(WorksheetSnapshot[] sheets, string fileName, int k = 2)
        {
            var result = new SpreadsheetEncoding { FileName = fileName };
            var metricsContainer = new CompressionMetricsContainer();
            int totalOriginal = 0, totalAnchor = 0, totalIndex = 0, totalFormat = 0, totalFinal = 0;

            foreach (var sheet in sheets)
            {
                if (IsEmptySheet(sheet)) continue;

                // Original token count
                var originalCells = CollectAllCells(sheet);
                int originalTokens = TokenCount(originalCells);

                // Stage 1: Structural-Anchor-Based Extraction
                var (rowAnchors, colAnchors) = FindStructuralAnchors(sheet, k);
                var (keptRows, keptCols) = ExpandAnchors(rowAnchors, colAnchors, sheet.MaxRow, sheet.MaxColumn, k);
                (keptRows, keptCols) = CompressHomogeneousRegions(sheet, keptRows, keptCols);

                var anchorCells = CollectKeptCells(sheet, keptRows, keptCols);
                int anchorTokens = TokenCount(anchorCells);

                // Stage 2: Inverted-Index Translation
                var (invertedIndex, formatMap) = CreateInvertedIndex(sheet, keptRows, keptCols);
                var mergedIndex = CreateInvertedIndexTranslation(invertedIndex);
                int indexTokens = TokenCount(mergedIndex);

                // Stage 3: Data-Format-Aware Aggregation
                var typeNfsGroups = GroupBySemanticType(sheet, formatMap);
                var aggregatedFormats = AggregateBySemanticType(typeNfsGroups);
                int formatTokens = TokenCount(aggregatedFormats);

                // Numeric sub-aggregation
                var numericGroups = new Dictionary<string, List<string>>();
                foreach (var kv in typeNfsGroups)
                {
                    var parsed = JsonDocument.Parse(kv.Key);
                    var typeVal = parsed.RootElement.TryGetProperty("type", out var tp) ? tp.GetString() : null;
                    if (typeVal == "integer" || typeVal == "float" || typeVal == "numeric")
                        numericGroups[kv.Key] = kv.Value;
                }
                var numericRanges = AggregateBySemanticType(numericGroups);

                // Assemble sheet encoding
                var sheetEncoding = new SheetEncoding
                {
                    StructuralAnchors = new StructuralAnchors
                    {
                        Rows = rowAnchors.OrderBy(r => r).ToList(),
                        Columns = colAnchors.OrderBy(c => c).Select(c => CellUtils.ColumnLetter(c)).ToList(),
                    },
                    Cells = mergedIndex,
                    Formats = aggregatedFormats,
                    NumericRanges = numericRanges,
                };
                int finalTokens = TokenCount(sheetEncoding);

                result.Sheets[sheet.Name] = sheetEncoding;

                var sheetMetrics = ComputeMetrics(originalTokens, anchorTokens, indexTokens, formatTokens, finalTokens);
                metricsContainer.Sheets[sheet.Name] = sheetMetrics;

                totalOriginal += originalTokens;
                totalAnchor += anchorTokens;
                totalIndex += indexTokens;
                totalFormat += formatTokens;
                totalFinal += finalTokens;
            }

            metricsContainer.Overall = ComputeMetrics(totalOriginal, totalAnchor, totalIndex, totalFormat, totalFinal);
            result.CompressionMetrics = metricsContainer;
            return result;
        }

        // =====================================================================
        // Stage 1: Structural-Anchor-Based Extraction
        // =====================================================================

        private sealed class RowAnalysis
        {
            public List<int> Fingerprints { get; } = new List<int>();
            public List<bool> Empty { get; } = new List<bool>();
            public List<(int min, int max)> Widths { get; } = new List<(int, int)>();
            public List<bool> IsHeader { get; } = new List<bool>();
            public List<int> NumericCount { get; } = new List<int>();
            public List<int> TextCount { get; } = new List<int>();
            public List<int> PopulatedCount { get; } = new List<int>();
        }

        private RowAnalysis AnalyzeRowsSinglePass(WorksheetSnapshot sheet)
        {
            var analysis = new RowAnalysis();

            for (int r = 1; r <= sheet.MaxRow; r++)
            {
                var fpParts = new List<(int valHash, bool isMerged, int styleId)>();
                bool isEmpty = true;
                int minC = sheet.MaxColumn + 1, maxC = 0;
                int populated = 0, boldCount = 0, centerCount = 0;
                int borderCount = 0, stringCount = 0, capsCount = 0;
                int numeric = 0, text = 0;

                for (int c = 1; c <= sheet.MaxColumn; c++)
                {
                    var cell = sheet.GetCell(r, c);
                    var coord = CellUtils.CellCoord(r, c);
                    bool isMerged = sheet.MergedCellMap.ContainsKey(coord);
                    var val = cell?.Value;
                    int valHash = val?.GetHashCode() ?? 0;
                    int styleId = cell != null ? CellUtils.GetStyleFingerprint(cell) : 0;
                    fpParts.Add((valHash, isMerged, styleId));

                    if (!string.IsNullOrWhiteSpace(val))
                    {
                        isEmpty = false;
                        populated++;
                        if (c < minC) minC = c;
                        if (c > maxC) maxC = c;

                        // Header heuristics
                        if (cell!.FontBold) boldCount++;
                        if (cell.AlignmentHorizontal == "Center" || cell.AlignmentHorizontal == "center")
                            centerCount++;
                        if (!string.IsNullOrEmpty(cell.BorderBottomStyle) &&
                            cell.BorderBottomStyle != "None" && cell.BorderBottomStyle != "none" &&
                            cell.BorderBottomStyle != "NoBoader") // ClosedXML uses "NoBoader" for none
                            borderCount++;

                        if (!cell.IsFormula && val != null)
                        {
                            stringCount++;
                            if (val.ToUpperInvariant() == val && val.Length > 1)
                                capsCount++;
                        }

                        // Data type
                        var dt = CellUtils.InferCellDataType(cell);
                        if (dt == "numeric") numeric++;
                        else if (dt == "text") text++;
                    }
                }

                // Compute row fingerprint
                int fp = 0;
                foreach (var part in fpParts)
                {
                    fp = CombineHash(CombineHash(fp, part.valHash), CombineHash(part.isMerged ? 1 : 0, part.styleId));
                }

                analysis.Fingerprints.Add(fp);
                analysis.Empty.Add(isEmpty);
                analysis.Widths.Add(isEmpty ? (0, 0) : (minC, maxC));
                analysis.PopulatedCount.Add(populated);
                analysis.NumericCount.Add(numeric);
                analysis.TextCount.Add(text);

                bool isHeader = false;
                if (populated > 0)
                {
                    if ((double)boldCount / populated > HeaderThreshold) isHeader = true;
                    else if ((double)centerCount / populated > HeaderThreshold) isHeader = true;
                    else if ((double)borderCount / populated > HeaderThreshold) isHeader = true;
                    else if (stringCount > 0 && (double)capsCount / stringCount > HeaderThreshold) isHeader = true;
                }
                analysis.IsHeader.Add(isHeader);
            }

            return analysis;
        }

        private int[] AnalyzeColsSinglePass(WorksheetSnapshot sheet)
        {
            var colParts = new List<int>[sheet.MaxColumn];
            for (int c = 0; c < sheet.MaxColumn; c++)
                colParts[c] = new List<int>();

            for (int r = 1; r <= sheet.MaxRow; r++)
            {
                for (int c = 1; c <= sheet.MaxColumn; c++)
                {
                    var cell = sheet.GetCell(r, c);
                    var coord = CellUtils.CellCoord(r, c);
                    bool isMerged = sheet.MergedCellMap.ContainsKey(coord);
                    int valHash = cell?.Value?.GetHashCode() ?? 0;
                    int styleId = cell != null ? CellUtils.GetStyleFingerprint(cell) : 0;
                    colParts[c - 1].Add(CombineHash(CombineHash(valHash, isMerged ? 1 : 0), styleId));
                }
            }

            var fingerprints = new int[sheet.MaxColumn];
            for (int c = 0; c < sheet.MaxColumn; c++)
            {
                int fp = 0;
                foreach (var h in colParts[c])
                    fp = CombineHash(fp, h);
                fingerprints[c] = fp;
            }
            return fingerprints;
        }

        private (List<int> rowBounds, List<int> colBounds) FindBoundaryCandidates(WorksheetSnapshot sheet)
        {
            var rowInfo = AnalyzeRowsSinglePass(sheet);
            var colFps = AnalyzeColsSinglePass(sheet);

            var rowBoundarySet = new HashSet<int>();

            for (int rIdx = 0; rIdx < sheet.MaxRow; rIdx++)
            {
                int r = rIdx + 1;

                // Fingerprint difference with previous row
                if (rIdx > 0 && rowInfo.Fingerprints[rIdx] != rowInfo.Fingerprints[rIdx - 1])
                {
                    rowBoundarySet.Add(r);
                    if (r > 1) rowBoundarySet.Add(r - 1);
                }

                // Empty row → boundary on both sides
                if (rowInfo.Empty[rIdx])
                {
                    rowBoundarySet.Add(r);
                    if (r > 1) rowBoundarySet.Add(r - 1);
                    if (r < sheet.MaxRow) rowBoundarySet.Add(r + 1);
                }

                // Column width transition
                if (rIdx > 0 && rowInfo.Widths[rIdx] != rowInfo.Widths[rIdx - 1])
                    rowBoundarySet.Add(r);

                // Header row → boundary
                if (rowInfo.IsHeader[rIdx])
                {
                    rowBoundarySet.Add(r);
                    if (r > 1) rowBoundarySet.Add(r - 1);
                }

                // Data type transition (numeric-dominant ↔ text-dominant)
                if (rIdx > 0)
                {
                    int prevN = rowInfo.NumericCount[rIdx - 1], prevT = rowInfo.TextCount[rIdx - 1];
                    int currN = rowInfo.NumericCount[rIdx], currT = rowInfo.TextCount[rIdx];
                    if (prevN + prevT >= 2 && currN + currT >= 2)
                    {
                        bool prevDomNum = prevN > prevT;
                        bool currDomNum = currN > currT;
                        if (prevDomNum != currDomNum)
                            rowBoundarySet.Add(r);
                    }
                }
            }

            // Column boundaries from fingerprint differences
            var colBoundarySet = new HashSet<int>();
            for (int cIdx = 1; cIdx < sheet.MaxColumn; cIdx++)
            {
                if (colFps[cIdx] != colFps[cIdx - 1])
                {
                    colBoundarySet.Add(cIdx + 1);
                    colBoundarySet.Add(cIdx);
                }
            }

            // Clamp
            var sortedRows = rowBoundarySet
                .Where(r => r >= 1 && r <= sheet.MaxRow)
                .OrderBy(r => r)
                .ToList();

            // Cap row boundaries
            if (sortedRows.Count > MaxBoundaryRows)
            {
                var headerRows = new HashSet<int>(
                    sortedRows.Where(r => r - 1 < rowInfo.IsHeader.Count && rowInfo.IsHeader[r - 1]));
                int step = sortedRows.Count / MaxBoundaryRows;
                var sampled = new HashSet<int>();
                for (int i = 0; i < sortedRows.Count; i += step)
                    sampled.Add(sortedRows[i]);
                sampled.Add(sortedRows[0]);
                sampled.Add(sortedRows[sortedRows.Count - 1]);
                foreach (var hr in headerRows) sampled.Add(hr);
                sortedRows = sampled.OrderBy(r => r).ToList();
            }

            var sortedCols = colBoundarySet
                .Where(c => c >= 1 && c <= sheet.MaxColumn)
                .OrderBy(c => c)
                .ToList();

            return (sortedRows, sortedCols);
        }

        private List<(int r1, int c1, int r2, int c2)> ComposeCandidatesConsecutive(
            List<int> rowBounds, List<int> colBounds)
        {
            var candidates = new HashSet<(int, int, int, int)>();

            if (rowBounds.Count < 2 || colBounds.Count < 2)
            {
                if (rowBounds.Count > 0 && colBounds.Count > 0)
                    candidates.Add((rowBounds[0], colBounds[0], rowBounds[rowBounds.Count - 1], colBounds[colBounds.Count - 1]));
                return candidates.ToList();
            }

            // Consecutive row × col pairs
            for (int i = 0; i < rowBounds.Count - 1; i++)
            {
                for (int j = 0; j < colBounds.Count - 1; j++)
                {
                    candidates.Add((rowBounds[i], colBounds[j], rowBounds[i + 1], colBounds[j + 1]));
                }
            }

            // Spanning candidates from first boundary
            int rFirst = rowBounds[0], cFirst = colBounds[0], cLast = colBounds[colBounds.Count - 1];
            for (int i = 2; i < Math.Min(rowBounds.Count, 10); i++)
                candidates.Add((rFirst, cFirst, rowBounds[i], cLast));

            // Full-span candidate
            candidates.Add((rowBounds[0], colBounds[0], rowBounds[rowBounds.Count - 1], colBounds[colBounds.Count - 1]));

            var list = candidates.ToList();
            if (list.Count > MaxCandidates)
            {
                list.Sort((a, b) => ((b.Item3 - b.Item1) * (b.Item4 - b.Item2))
                    .CompareTo((a.Item3 - a.Item1) * (a.Item4 - a.Item2)));
                list = list.Take(MaxCandidates).ToList();
            }
            return list;
        }

        private List<(int r1, int c1, int r2, int c2)> FilterCandidates(
            WorksheetSnapshot sheet,
            List<(int r1, int c1, int r2, int c2)> candidates,
            RowAnalysis rowInfo)
        {
            var filtered = new List<(int, int, int, int)>();

            foreach (var (r1, c1, r2, c2) in candidates)
            {
                if (r2 - r1 < 1 || c2 - c1 < 1) continue;

                int totalCells = (r2 - r1 + 1) * (c2 - c1 + 1);
                double colFraction = sheet.MaxColumn > 0 ? (double)(c2 - c1 + 1) / sheet.MaxColumn : 1.0;

                int estPopulated = 0;
                for (int r = r1; r <= r2; r++)
                {
                    if (r - 1 < rowInfo.PopulatedCount.Count)
                        estPopulated += rowInfo.PopulatedCount[r - 1];
                }
                estPopulated = (int)(estPopulated * colFraction);

                if (totalCells > 0 && (double)estPopulated / totalCells < SparsityThreshold) continue;

                bool hasHeader = false;
                for (int r = r1; r <= Math.Min(r1 + 2, r2); r++)
                {
                    if (r - 1 < rowInfo.IsHeader.Count && rowInfo.IsHeader[r - 1])
                    {
                        hasHeader = true;
                        break;
                    }
                }

                if (!hasHeader && totalCells > 0)
                {
                    // Fallback: >50% populated and text-dominant first row
                    if ((double)estPopulated / totalCells > 0.5)
                    {
                        if (r1 - 1 < rowInfo.TextCount.Count &&
                            rowInfo.TextCount[r1 - 1] > rowInfo.NumericCount[r1 - 1])
                            hasHeader = true;
                    }
                }

                if (!hasHeader) continue;
                filtered.Add((r1, c1, r2, c2));
            }

            return filtered;
        }

        private static double CalculateIoU(
            (int r1, int c1, int r2, int c2) box1,
            (int r1, int c1, int r2, int c2) box2)
        {
            int interR1 = Math.Max(box1.r1, box2.r1);
            int interC1 = Math.Max(box1.c1, box2.c1);
            int interR2 = Math.Min(box1.r2, box2.r2);
            int interC2 = Math.Min(box1.c2, box2.c2);

            int interArea = Math.Max(0, interR2 - interR1 + 1) * Math.Max(0, interC2 - interC1 + 1);
            int area1 = (box1.r2 - box1.r1 + 1) * (box1.c2 - box1.c1 + 1);
            int area2 = (box2.r2 - box2.r1 + 1) * (box2.c2 - box2.c1 + 1);
            int unionArea = area1 + area2 - interArea;

            return unionArea > 0 ? (double)interArea / unionArea : 0.0;
        }

        private List<(int r1, int c1, int r2, int c2)> NmsCandidates(
            List<(int r1, int c1, int r2, int c2)> candidates,
            RowAnalysis rowInfo)
        {
            if (candidates.Count == 0) return candidates;

            var scores = new double[candidates.Count];
            for (int i = 0; i < candidates.Count; i++)
            {
                var (r1, _, r2, _) = candidates[i];
                double score = 0;
                for (int r = r1; r <= Math.Min(r1 + 2, r2); r++)
                {
                    if (r - 1 < rowInfo.IsHeader.Count && rowInfo.IsHeader[r - 1])
                        score += HeaderScoreWeight;
                }
                for (int r = r1; r <= r2; r++)
                {
                    if (r - 1 < rowInfo.PopulatedCount.Count)
                        score += rowInfo.PopulatedCount[r - 1];
                }
                scores[i] = score;
            }

            var indices = Enumerable.Range(0, candidates.Count)
                .OrderByDescending(i => scores[i])
                .ToList();

            var keep = new List<(int, int, int, int)>();
            while (indices.Count > 0)
            {
                int best = indices[0];
                indices.RemoveAt(0);
                keep.Add(candidates[best]);

                indices = indices
                    .Where(idx => CalculateIoU(candidates[best], candidates[idx]) < NmsIouThreshold)
                    .ToList();
            }

            return keep;
        }

        private (List<int> rowAnchors, List<int> colAnchors) FindStructuralAnchors(
            WorksheetSnapshot sheet, int k)
        {
            var rowInfo = AnalyzeRowsSinglePass(sheet);
            var (rowBounds, colBounds) = FindBoundaryCandidates(sheet);

            if (rowBounds.Count == 0 || colBounds.Count == 0)
                return (Enumerable.Range(1, sheet.MaxRow).ToList(), Enumerable.Range(1, sheet.MaxColumn).ToList());

            var candidates = ComposeCandidatesConsecutive(rowBounds, colBounds);
            candidates = FilterCandidates(sheet, candidates, rowInfo);
            candidates = NmsCandidates(candidates, rowInfo);

            if (candidates.Count == 0)
                return (Enumerable.Range(1, sheet.MaxRow).ToList(), Enumerable.Range(1, sheet.MaxColumn).ToList());

            var rowAnchorSet = new HashSet<int>();
            var colAnchorSet = new HashSet<int>();

            foreach (var (r1, c1, r2, c2) in candidates)
            {
                rowAnchorSet.Add(r1); rowAnchorSet.Add(r2);
                colAnchorSet.Add(c1); colAnchorSet.Add(c2);
            }

            // Add header rows
            for (int i = 0; i < rowInfo.IsHeader.Count; i++)
            {
                if (rowInfo.IsHeader[i])
                    rowAnchorSet.Add(i + 1);
            }

            return (rowAnchorSet.OrderBy(r => r).ToList(), colAnchorSet.OrderBy(c => c).ToList());
        }

        private (List<int> rows, List<int> cols) ExpandAnchors(
            List<int> rowAnchors, List<int> colAnchors, int maxRow, int maxCol, int k)
        {
            var keptRows = new HashSet<int>();
            foreach (int anchor in rowAnchors)
            {
                for (int i = Math.Max(1, anchor - k); i <= Math.Min(maxRow, anchor + k); i++)
                    keptRows.Add(i);
            }

            var keptCols = new HashSet<int>();
            foreach (int anchor in colAnchors)
            {
                for (int i = Math.Max(1, anchor - k); i <= Math.Min(maxCol, anchor + k); i++)
                    keptCols.Add(i);
            }

            return (keptRows.OrderBy(r => r).ToList(), keptCols.OrderBy(c => c).ToList());
        }

        private (List<int> rows, List<int> cols) CompressHomogeneousRegions(
            WorksheetSnapshot sheet, List<int> rows, List<int> cols)
        {
            bool RowIsHomogeneous(int r)
            {
                string? firstVal = null;
                string? firstFmt = null;
                bool first = true;
                foreach (int c in cols)
                {
                    var cell = sheet.GetCell(r, c);
                    var val = cell?.Value;
                    var fmt = cell?.NumberFormat ?? "General";
                    if (first) { firstVal = val; firstFmt = fmt; first = false; }
                    else if (val != firstVal || fmt != firstFmt) return false;
                }
                return true;
            }

            bool ColIsHomogeneous(int c)
            {
                string? firstVal = null;
                string? firstFmt = null;
                bool first = true;
                foreach (int r in rows)
                {
                    var cell = sheet.GetCell(r, c);
                    var val = cell?.Value;
                    var fmt = cell?.NumberFormat ?? "General";
                    if (first) { firstVal = val; firstFmt = fmt; first = false; }
                    else if (val != firstVal || fmt != firstFmt) return false;
                }
                return true;
            }

            var keptRows = rows.Where(r => !RowIsHomogeneous(r)).ToList();
            var keptCols = cols.Where(c => !ColIsHomogeneous(c)).ToList();

            if (keptRows.Count == 0) keptRows = rows;
            if (keptCols.Count == 0) keptCols = cols;

            return (keptRows, keptCols);
        }

        // =====================================================================
        // Stage 2: Inverted-Index Translation
        // =====================================================================

        private (Dictionary<string, List<string>> index, Dictionary<string, List<string>> formatMap)
            CreateInvertedIndex(WorksheetSnapshot sheet, List<int> rows, List<int> cols)
        {
            var inverted = new Dictionary<string, List<string>>();
            var formatGroups = new Dictionary<string, List<string>>();

            foreach (int r in rows)
            {
                foreach (int c in cols)
                {
                    var coord = CellUtils.CellCoord(r, c);
                    var cell = sheet.GetCell(r, c);

                    // Resolve merged cell value
                    string? value;
                    if (sheet.MergedCellMap.TryGetValue(coord, out var startCoord))
                    {
                        var (sc, sr) = CellUtils.SplitCellRef(startCoord);
                        var startCell = sheet.GetCell(sr, CellUtils.ColumnNumber(sc));
                        value = startCell?.Value;
                    }
                    else
                    {
                        value = cell?.Value;
                    }

                    if (value == null) continue;

                    AddToGroup(inverted, value, coord);

                    var semType = CellUtils.DetectSemanticType(cell ?? new CellData());
                    var nfs = CellUtils.GetNumberFormatString(cell ?? new CellData());
                    var fmtKey = JsonSerializer.Serialize(new { type = semType, nfs }, new JsonSerializerOptions { WriteIndented = false });
                    AddToGroup(formatGroups, fmtKey, coord);
                }
            }

            return (inverted, formatGroups);
        }

        public List<string> MergeCellRanges(List<string> refs)
        {
            if (refs.Count == 0) return new List<string>();

            var coords = new HashSet<(int row, int col)>();
            foreach (var refStr in refs)
            {
                try
                {
                    var (colLetter, row) = CellUtils.SplitCellRef(refStr);
                    coords.Add((row, CellUtils.ColumnNumber(colLetter)));
                }
                catch { }
            }

            if (coords.Count == 0) return new List<string>();

            var processed = new HashSet<(int, int)>();
            var ranges = new List<string>();

            foreach (var (row, col) in coords.OrderBy(x => x.row).ThenBy(x => x.col))
            {
                if (processed.Contains((row, col))) continue;

                // Expand right
                int width = 1;
                while (coords.Contains((row, col + width)) && !processed.Contains((row, col + width)))
                    width++;

                // Expand down (full-width rows only)
                int height = 1;
                bool expanding = true;
                while (expanding)
                {
                    int nextRow = row + height;
                    for (int w = 0; w < width; w++)
                    {
                        if (!coords.Contains((nextRow, col + w)) || processed.Contains((nextRow, col + w)))
                        {
                            expanding = false;
                            break;
                        }
                    }
                    if (expanding) height++;
                }

                int endRow = row + height - 1;
                int endCol = col + width - 1;
                var startRef = CellUtils.CellCoord(row, col);
                var endRef = CellUtils.CellCoord(endRow, endCol);

                ranges.Add(width == 1 && height == 1 ? startRef : $"{startRef}:{endRef}");

                for (int dr = 0; dr < height; dr++)
                    for (int dc = 0; dc < width; dc++)
                        processed.Add((row + dr, col + dc));
            }

            return ranges;
        }

        private Dictionary<string, List<string>> CreateInvertedIndexTranslation(
            Dictionary<string, List<string>> invertedIndex)
        {
            var result = new Dictionary<string, List<string>>();
            foreach (var kv in invertedIndex)
            {
                if (string.IsNullOrWhiteSpace(kv.Key)) continue;
                result[kv.Key] = MergeCellRanges(kv.Value);
            }
            return result;
        }

        // =====================================================================
        // Stage 3: Data-Format-Aware Aggregation
        // =====================================================================

        private Dictionary<string, List<string>> GroupBySemanticType(
            WorksheetSnapshot sheet, Dictionary<string, List<string>> formatMap)
        {
            var groups = new Dictionary<string, List<string>>();
            foreach (var kv in formatMap)
            {
                foreach (var refStr in kv.Value)
                {
                    try
                    {
                        var (colLetter, row) = CellUtils.SplitCellRef(refStr);
                        var cell = sheet.GetCell(row, CellUtils.ColumnNumber(colLetter));
                        if (cell == null) continue;

                        var semType = CellUtils.DetectSemanticType(cell);
                        var nfs = CellUtils.GetNumberFormatString(cell);
                        var key = JsonSerializer.Serialize(new { type = semType, nfs },
                            new JsonSerializerOptions { WriteIndented = false });
                        AddToGroup(groups, key, refStr);
                    }
                    catch { }
                }
            }
            return groups;
        }

        private Dictionary<string, List<string>> AggregateBySemanticType(
            Dictionary<string, List<string>> typeNfsGroups)
        {
            var result = new Dictionary<string, List<string>>();
            foreach (var kv in typeNfsGroups)
            {
                if (kv.Value.Count > 0)
                    result[kv.Key] = MergeCellRanges(kv.Value);
            }
            return result;
        }

        // =====================================================================
        // Helpers
        // =====================================================================

        private static bool IsEmptySheet(WorksheetSnapshot sheet)
        {
            return sheet.MaxRow == 0 || sheet.MaxColumn == 0;
        }

        private static Dictionary<string, string> CollectAllCells(WorksheetSnapshot sheet)
        {
            var cells = new Dictionary<string, string>();
            for (int r = 1; r <= sheet.MaxRow; r++)
            {
                for (int c = 1; c <= sheet.MaxColumn; c++)
                {
                    var val = sheet.GetCell(r, c)?.Value;
                    if (val != null)
                        cells[CellUtils.CellCoord(r, c)] = val;
                }
            }
            return cells;
        }

        private static Dictionary<string, string> CollectKeptCells(
            WorksheetSnapshot sheet, List<int> rows, List<int> cols)
        {
            var cells = new Dictionary<string, string>();
            foreach (int r in rows)
            {
                foreach (int c in cols)
                {
                    var val = sheet.GetCell(r, c)?.Value;
                    if (val != null)
                        cells[CellUtils.CellCoord(r, c)] = val;
                }
            }
            return cells;
        }

        private static int TokenCount(object data)
        {
            return JsonSerializer.Serialize(data).Length;
        }

        private static double CompressionRatio(int original, int compressed)
        {
            if (compressed == 0) return 0.0;
            return (double)original / compressed;
        }

        private static SheetMetrics ComputeMetrics(
            int original, int anchor, int index, int fmt, int final)
        {
            return new SheetMetrics
            {
                OriginalTokens = original,
                AfterAnchorTokens = anchor,
                AfterInvertedIndexTokens = index,
                AfterFormatTokens = fmt,
                FinalTokens = final,
                AnchorRatio = CompressionRatio(original, anchor),
                InvertedIndexRatio = CompressionRatio(original, index),
                FormatRatio = CompressionRatio(original, fmt),
                OverallRatio = CompressionRatio(original, final),
            };
        }

        private static void AddToGroup<TKey, TValue>(
            Dictionary<TKey, List<TValue>> dict, TKey key, TValue value)
            where TKey : notnull
        {
            if (!dict.TryGetValue(key, out var list))
            {
                list = new List<TValue>();
                dict[key] = list;
            }
            list.Add(value);
        }

        private static int CombineHash(int h1, int h2) =>
            unchecked(((h1 << 5) + h1) ^ h2);
    }
}
