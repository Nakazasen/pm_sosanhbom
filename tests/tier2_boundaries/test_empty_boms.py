"""Tier 2 Boundary Tests: Empty and Malformed BOM Handling.

Verifies:
1. Zero-byte files raise structured errors without unhandled crashes.
2. Header-only files parse to empty trees cleanly.
3. Missing or empty sheets handled cleanly.
4. R3 files with empty tables produce empty DataFrames with standard schema.
5. Three-way reconciliation handles empty PLM/R3/CTTT inputs fail-safely.
"""

from pathlib import Path
import openpyxl
import pandas as pd
import pytest

from src.core.tree_parser import PLMTreeParser
from src.automation.sap.parser import ResilientR3Parser, SAPParseError
from src.core.reconciliation import ReconciliationEngine


class TestEmptyBOMBoundaries:
    """Boundary test suite for empty, zero-byte, and header-only BOM structures."""

    def test_zero_byte_plm_file_raises_parse_error(self, tmp_path: Path):
        """0-byte file raises exception when loaded."""
        zero_file = tmp_path / "empty_plm.xlsx"
        zero_file.write_bytes(b"")

        parser = PLMTreeParser()
        with pytest.raises(Exception):
            parser.parse_excel(zero_file)

    def test_zero_byte_r3_file_raises_parse_error(self, tmp_path: Path):
        """0-byte file raises SAPParseError."""
        zero_file = tmp_path / "empty_r3.xls"
        zero_file.write_bytes(b"")

        parser = ResilientR3Parser()
        with pytest.raises(SAPParseError):
            parser.parse(zero_file)

    def test_plm_header_only_returns_empty_tree(self, tmp_path: Path):
        """PLM file with 14 standard header columns but 0 data rows yields empty tree."""
        f = tmp_path / "header_only_plm.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "PLM"
        headers = [
            "No", "Level", "Item Type", "Item ID", "Has Children", "Quantity",
            "First Parts", "Second BOM flag", "Occurrence Effectivities",
            "Item Revision:Projects List", "Item Name", "Notice No", "Revision",
            "Item Rev:Release Status"
        ]
        ws.append(headers)
        wb.save(f)

        parser = PLMTreeParser()
        tree = parser.parse_excel(f)
        assert len(tree.roots) == 0
        assert len(tree.flatten()) == 0

    def test_r3_empty_table_returns_empty_dataframe(self, tmp_path: Path):
        """R3 HTML file with empty table tags returns empty DataFrame with schema."""
        f = tmp_path / "empty_table_r3.xls"
        f.write_text("<html><body><table><tr><th>Component</th><th>Quantity</th></tr></table></body></html>", encoding="utf-8")

        parser = ResilientR3Parser()
        df = parser.parse(f)
        assert len(df) == 0
        assert "part_code" in df.columns
        assert "quantity" in df.columns

    def test_reconciliation_engine_with_empty_inputs(self):
        """Reconciliation engine gracefully handles empty PLM, R3, or CTTT datasets."""
        engine = ReconciliationEngine()
        result_df = engine.reconcile_three_way(
            cttt_data=pd.DataFrame(columns=["part_code", "quantity_cttt", "delivery_date"]),
            plm_data=pd.DataFrame(columns=["part_code", "quantity_plm", "unit"]),
            r3_data=pd.DataFrame(columns=["part_code", "quantity_r3", "rev_r3"]),
        )
        assert isinstance(result_df, pd.DataFrame)
        assert len(result_df) == 0
