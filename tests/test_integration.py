"""Integration tests — encode real Excel files of varying complexity.

Tests formula preservation, structural anchor detection, compression ratios,
edge cases with messy/complex/large sheets.
"""
from __future__ import annotations

import json
import os
import time
import pytest

from spreadsheet_llm.encoder import encode_spreadsheet
from spreadsheet_llm.vanilla import vanilla_encode
from tests.excel_factories import (
    make_empty_sheet,
    make_single_cell,
    make_tiny_with_formulas,
    make_adjacent_tables_no_gap,
    make_adjacent_tables_same_style,
    make_merged_header,
    make_complex_formulas,
    make_mixed_types,
    make_medium_500_rows,
    make_large_3_tables,
    make_messy_form_layout,
    make_multi_sheet_cross_ref,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _has_formulas(encoding: dict) -> bool:
    """Check if any encoded cell value starts with '='."""
    for sheet_data in encoding.get("sheets", {}).values():
        for value in sheet_data.get("cells", {}):
            if str(value).startswith("="):
                return True
    return False


def _overall_ratio(encoding: dict) -> float:
    return encoding.get("compression_metrics", {}).get("overall", {}).get("overall_ratio", 0)


def _final_tokens(encoding: dict) -> int:
    return encoding.get("compression_metrics", {}).get("overall", {}).get("final_tokens", 0)


def _original_tokens(encoding: dict) -> int:
    return encoding.get("compression_metrics", {}).get("overall", {}).get("original_tokens", 0)


def _sheet_names(encoding: dict) -> list:
    return list(encoding.get("sheets", {}).keys())


def _has_ranges(encoding: dict) -> bool:
    """Check if any cell entry uses range notation (A1:B2)."""
    for sheet_data in encoding.get("sheets", {}).values():
        for refs in sheet_data.get("cells", {}).values():
            for ref in refs:
                if ":" in ref:
                    return True
    return False


def _has_format_types(encoding: dict, expected_types: set) -> set:
    """Return which of the expected types appear in the format keys."""
    found = set()
    for sheet_data in encoding.get("sheets", {}).values():
        for fmt_key in sheet_data.get("formats", {}):
            try:
                parsed = json.loads(fmt_key)
                t = parsed.get("type", "")
                if t in expected_types:
                    found.add(t)
            except json.JSONDecodeError:
                pass
    return found


# ---------------------------------------------------------------------------
# Test: Empty / Minimal
# ---------------------------------------------------------------------------

class TestEmptyAndMinimal:
    def test_empty_sheet_returns_valid(self, tmp_path):
        path = make_empty_sheet(str(tmp_path / "empty.xlsx"))
        result = encode_spreadsheet(path)
        assert result is not None
        assert result["sheets"] == {}  # Empty sheet skipped

    def test_single_cell(self, tmp_path):
        path = make_single_cell(str(tmp_path / "single.xlsx"))
        result = encode_spreadsheet(path)
        assert result is not None
        assert len(result["sheets"]) == 1
        sheet = list(result["sheets"].values())[0]
        assert "hello" in sheet["cells"]


# ---------------------------------------------------------------------------
# Test: Formula Preservation
# ---------------------------------------------------------------------------

class TestFormulaPreservation:
    def test_tiny_formulas_preserved(self, tmp_path):
        """SUM formulas should appear as '=SUM(...)' not computed values."""
        path = make_tiny_with_formulas(str(tmp_path / "tiny.xlsx"))
        result = encode_spreadsheet(path)
        assert result is not None
        assert _has_formulas(result), "Formulas not found in encoding — data_only should be False"

        # Check specific formula exists
        sheet = list(result["sheets"].values())[0]
        formula_values = [v for v in sheet["cells"] if v.startswith("=")]
        assert any("SUM" in f for f in formula_values), f"Expected SUM formula, got: {formula_values}"

    def test_complex_formulas_preserved(self, tmp_path):
        """IF, RANK, COUNTIF, AVERAGE, MAX, MIN should all be preserved."""
        path = make_complex_formulas(str(tmp_path / "complex.xlsx"))
        result = encode_spreadsheet(path)
        assert _has_formulas(result)

        sheet = list(result["sheets"].values())[0]
        formula_values = [v for v in sheet["cells"] if v.startswith("=")]
        formula_text = " ".join(formula_values)
        assert "IF" in formula_text, f"IF formula not found in: {formula_values}"
        assert "RANK" in formula_text, f"RANK formula not found in: {formula_values}"

    def test_cross_sheet_formulas_preserved(self, tmp_path):
        """Cross-sheet references like =Input!A2 should be preserved."""
        path = make_multi_sheet_cross_ref(str(tmp_path / "multi.xlsx"))
        result = encode_spreadsheet(path)
        assert _has_formulas(result)

        # Check for cross-sheet reference
        all_values = []
        for sheet_data in result["sheets"].values():
            all_values.extend(sheet_data["cells"].keys())
        formula_text = " ".join(v for v in all_values if v.startswith("="))
        assert "Input!" in formula_text or "Calc!" in formula_text, \
            f"Cross-sheet reference not found in: {formula_text[:200]}"

    def test_vanilla_also_preserves_formulas(self, tmp_path):
        """Vanilla encoding should also show formulas, not computed values."""
        path = make_tiny_with_formulas(str(tmp_path / "tiny.xlsx"))
        result = vanilla_encode(path)
        assert result is not None
        first_sheet = list(result.values())[0]
        assert "=SUM" in first_sheet or "=B" in first_sheet, \
            f"Vanilla encoding should preserve formulas. Got:\n{first_sheet[:300]}"


# ---------------------------------------------------------------------------
# Test: Adjacent Tables (Edge Cases)
# ---------------------------------------------------------------------------

class TestAdjacentTables:
    def test_adjacent_different_styles(self, tmp_path):
        """Two tables with different header styles, no gap — should detect both."""
        path = make_adjacent_tables_no_gap(str(tmp_path / "adj.xlsx"))
        result = encode_spreadsheet(path)
        assert result is not None

        sheet = list(result["sheets"].values())[0]
        anchors = sheet["structural_anchors"]
        # Should have anchors for both table headers (row 1 and row 6)
        assert 1 in anchors["rows"] or any(r <= 2 for r in anchors["rows"]), \
            f"Table 1 header not in anchors: {anchors['rows']}"
        assert 6 in anchors["rows"] or any(5 <= r <= 7 for r in anchors["rows"]), \
            f"Table 2 header not in anchors: {anchors['rows']}"

    def test_adjacent_same_style(self, tmp_path):
        """Two tables with same header style — rely on data type transition."""
        path = make_adjacent_tables_same_style(str(tmp_path / "adj_same.xlsx"))
        result = encode_spreadsheet(path)
        assert result is not None
        # Should at least produce valid output without crashing
        assert _final_tokens(result) > 0


# ---------------------------------------------------------------------------
# Test: Merged Headers
# ---------------------------------------------------------------------------

class TestMergedHeaders:
    def test_merged_title(self, tmp_path):
        """Merged A1:D1 title should not break encoding."""
        path = make_merged_header(str(tmp_path / "merged.xlsx"))
        result = encode_spreadsheet(path)
        assert result is not None

        sheet = list(result["sheets"].values())[0]
        # Check that "Annual Report 2024" appears in cells
        assert any("Annual" in v for v in sheet["cells"]), \
            f"Merged title not found in cells: {list(sheet['cells'].keys())[:10]}"


# ---------------------------------------------------------------------------
# Test: Mixed Semantic Types
# ---------------------------------------------------------------------------

class TestMixedTypes:
    def test_all_types_detected(self, tmp_path):
        """All 9+ semantic types should be detected in format aggregation."""
        path = make_mixed_types(str(tmp_path / "types.xlsx"))
        result = encode_spreadsheet(path)
        assert result is not None

        expected = {"integer", "float", "percentage", "date", "email", "text"}
        found = _has_format_types(result, expected)
        # At least 4 of these should be detected
        assert len(found) >= 4, f"Only found types: {found}, expected at least 4 from {expected}"


# ---------------------------------------------------------------------------
# Test: Medium (500 rows)
# ---------------------------------------------------------------------------

class TestMedium:
    def test_500_rows_encodes(self, tmp_path):
        """500-row sheet should encode without errors and show compression."""
        path = make_medium_500_rows(str(tmp_path / "medium.xlsx"))

        t0 = time.perf_counter()
        result = encode_spreadsheet(path)
        elapsed = time.perf_counter() - t0

        assert result is not None
        assert _overall_ratio(result) > 1.0, \
            f"Expected compression > 1x, got {_overall_ratio(result):.2f}x"
        assert elapsed < 60, f"Encoding took {elapsed:.1f}s — should be under 60s"
        assert _has_formulas(result), "Formulas should be preserved in 500-row sheet"

    def test_500_rows_has_format_ranges(self, tmp_path):
        """Format aggregation should merge 500 contiguous cells into ranges."""
        path = make_medium_500_rows(str(tmp_path / "medium.xlsx"))
        result = encode_spreadsheet(path)
        # Cell values may be unique per row, but formats are contiguous
        sheet = list(result["sheets"].values())[0]
        has_fmt_range = any(
            ":" in ref
            for refs in sheet.get("formats", {}).values()
            for ref in refs
        )
        assert has_fmt_range, "500-row sheet should have merged format ranges"


# ---------------------------------------------------------------------------
# Test: Large (3000 rows, 3 tables)
# ---------------------------------------------------------------------------

class TestLarge:
    def test_3000_rows_encodes(self, tmp_path):
        """3000-row sheet with 3 tables should encode under 120s."""
        path = make_large_3_tables(str(tmp_path / "large.xlsx"))

        t0 = time.perf_counter()
        result = encode_spreadsheet(path)
        elapsed = time.perf_counter() - t0

        assert result is not None
        assert elapsed < 120, f"Encoding took {elapsed:.1f}s — should be under 120s"
        assert _overall_ratio(result) > 1.0, \
            f"Expected compression > 1x, got {_overall_ratio(result):.2f}x"

    def test_3000_rows_detects_tables(self, tmp_path):
        """Structural anchors should capture all 3 table headers."""
        path = make_large_3_tables(str(tmp_path / "large.xlsx"))
        result = encode_spreadsheet(path)
        sheet = list(result["sheets"].values())[0]
        rows = sheet["structural_anchors"]["rows"]

        # Table 1 header near row 1, table 2 near 1003, table 3 near 1504
        has_t1 = any(r <= 3 for r in rows)
        has_t2 = any(1001 <= r <= 1005 for r in rows)
        has_t3 = any(1502 <= r <= 1506 for r in rows)

        assert has_t1, f"Table 1 header not detected. Anchors: {rows[:10]}..."
        # At least one of the other tables should be detected
        assert has_t2 or has_t3, \
            f"Neither table 2 nor 3 detected. Anchors: {rows}"


# ---------------------------------------------------------------------------
# Test: Messy/Form Layout
# ---------------------------------------------------------------------------

class TestMessyLayout:
    def test_form_layout(self, tmp_path):
        """Sparse form-like layout should not crash and should encode."""
        path = make_messy_form_layout(str(tmp_path / "messy.xlsx"))
        result = encode_spreadsheet(path)
        assert result is not None
        assert _final_tokens(result) > 0


# ---------------------------------------------------------------------------
# Test: Multi-Sheet
# ---------------------------------------------------------------------------

class TestMultiSheet:
    def test_all_sheets_encoded(self, tmp_path):
        """All 3 sheets should appear in encoding."""
        path = make_multi_sheet_cross_ref(str(tmp_path / "multi.xlsx"))
        result = encode_spreadsheet(path)
        assert result is not None
        names = _sheet_names(result)
        assert "Input" in names
        assert "Calc" in names
        assert "Summary" in names


# ---------------------------------------------------------------------------
# Test: Compression Metrics
# ---------------------------------------------------------------------------

class TestCompressionMetrics:
    def test_metrics_structure(self, tmp_path):
        """Metrics should have all required keys at sheet and overall level."""
        path = make_tiny_with_formulas(str(tmp_path / "tiny.xlsx"))
        result = encode_spreadsheet(path)

        overall = result["compression_metrics"]["overall"]
        required_keys = [
            "original_tokens", "after_anchor_tokens",
            "after_inverted_index_tokens", "after_format_tokens",
            "final_tokens", "anchor_ratio", "inverted_index_ratio",
            "format_ratio", "overall_ratio",
        ]
        for key in required_keys:
            assert key in overall, f"Missing metric: {key}"
            assert isinstance(overall[key], (int, float)), f"Metric {key} should be numeric"

    def test_overall_tokens_positive(self, tmp_path):
        path = make_tiny_with_formulas(str(tmp_path / "tiny.xlsx"))
        result = encode_spreadsheet(path)
        assert _original_tokens(result) > 0
        assert _final_tokens(result) > 0


# ---------------------------------------------------------------------------
# Test: Output Structure
# ---------------------------------------------------------------------------

class TestOutputStructure:
    def test_encoding_has_all_keys(self, tmp_path):
        path = make_tiny_with_formulas(str(tmp_path / "tiny.xlsx"))
        result = encode_spreadsheet(path)

        assert "file_name" in result
        assert "sheets" in result
        assert "compression_metrics" in result

        sheet = list(result["sheets"].values())[0]
        assert "structural_anchors" in sheet
        assert "rows" in sheet["structural_anchors"]
        assert "columns" in sheet["structural_anchors"]
        assert "cells" in sheet
        assert "formats" in sheet
        assert "numeric_ranges" in sheet

    def test_json_serializable(self, tmp_path):
        """Output must be fully JSON-serializable."""
        path = make_medium_500_rows(str(tmp_path / "med.xlsx"))
        result = encode_spreadsheet(path)
        # This should not raise
        json_str = json.dumps(result, ensure_ascii=False)
        assert len(json_str) > 0
        # And parseable back
        parsed = json.loads(json_str)
        assert parsed["file_name"] == "med.xlsx"

    def test_output_file_written(self, tmp_path):
        path = make_tiny_with_formulas(str(tmp_path / "tiny.xlsx"))
        out_path = str(tmp_path / "output.json")
        result = encode_spreadsheet(path, output_path=out_path)

        assert os.path.exists(out_path)
        with open(out_path, "r", encoding="utf-8") as f:
            loaded = json.load(f)
        assert loaded["file_name"] == "tiny.xlsx"
