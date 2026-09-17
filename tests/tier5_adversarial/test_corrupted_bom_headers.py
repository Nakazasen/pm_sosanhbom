"""Tier 5 Adversarial Tests: Corrupted BOM Headers & Column Scrambling.

Verifies:
1. Missing mandatory columns (missing Level, Item ID, Quantity).
2. Duplicated column names with conflicting values.
3. Random / unknown headers interspersed with real ones.
4. Scrambled column order and bizarre casing variations.
5. BOM header placed deep down in the sheet (e.g. after title banners/empty rows).
6. Invisible unicode characters and BOM marks in header text.
"""

from __future__ import annotations

from pathlib import Path
import pandas as pd
import pytest

from src.core.tree_parser import BOMTreeParser
from src.automation.sap.parser import ResilientR3Parser
from src.automation.sap.models import SAPParseError


class TestAdversarialCorruptedBOMHeaders:
    """Test suite for corrupted and scrambled headers."""

    def test_missing_mandatory_level_and_id_columns(self):
        """DataFrame missing 'level' or 'item_id' produces an empty tree safely without unhandled key error."""
        df_corrupt = pd.DataFrame({
            "Random_Column_A": ["Val1", "Val2"],
            "Description": ["Desc1", "Desc2"],
            "Quantity": [1.0, 2.0],
        })

        parser = BOMTreeParser()
        tree = parser.parse_dataframe(df_corrupt)
        assert tree is not None
        # Without level and item_id, no valid roots can be constructed
        assert len(tree.roots) == 0

    def test_duplicate_column_headers(self):
        """Data with duplicated column headers (e.g. two 'Item ID' columns) parses stably."""
        df_dups = pd.DataFrame(
            [
                ["1", "PART_A", "PART_A_ALT", "5.0", "MAIN_UNIT"],
                ["..2", "PART_B", "PART_B_ALT", "10.0", "SUB_UNIT"],
            ],
            columns=["Level", "Item ID", "Item ID", "Quantity", "Item Name"]
        )

        parser = BOMTreeParser()
        tree = parser.parse_dataframe(df_dups)
        assert tree is not None
        assert len(tree.roots) == 1
        assert tree.roots[0].item_id in ("PART_A", "PART_A_ALT")

    def test_scrambled_casing_and_leading_trailing_whitespace(self):
        """Headers with arbitrary casing and heavy whitespace match aliases correctly."""
        headers = [
            "  lEvEL  ",
            "\tItEm_Id\n",
            "   HaS_cHiLdReN   ",
            " QuAnTiTy ",
            "   itEM_nAMe\r\n",
        ]
        parser = BOMTreeParser()
        mapping = parser.detect_column_mapping(headers)
        assert "level" in mapping
        assert "item_id" in mapping
        assert "has_children" in mapping
        assert "quantity" in mapping
        assert "item_name" in mapping

    def test_invisible_bom_mark_and_unicode_space_in_header(self):
        """Headers polluted with UTF-8 BOM (\\ufeff) and zero-width spaces (\\u200b) match successfully."""
        headers = [
            "\ufeffLevel",
            "Item\u200bID",
            "Quantity\u00a0",  # Non-breaking space
            "Item Name",
        ]
        parser = BOMTreeParser()
        mapping = parser.detect_column_mapping(headers)
        assert "level" in mapping
        assert "item_id" in mapping
        assert "quantity" in mapping
        assert "item_name" in mapping

    def test_sap_r3_corrupted_header_missing_part_code(self):
        """SAP R3 file with completely empty or non-tabular content raises SAPParseError."""
        parser = ResilientR3Parser()
        with pytest.raises(SAPParseError):
            parser.parse_content("", format_hint="tsv")

    def test_header_offset_deep_in_matrix(self):
        """BOM where metadata banners occupy first 4 rows before actual headers."""
        matrix = [
            ["KYOCERA DOCUMENT SOLUTIONS", None, None, None, None],
            ["PROJECT: VENUS_2026", None, None, None, None],
            ["CONFIDENTIAL FACTORY BOM", None, None, None, None],
            [None, None, None, None, None],  # Blank line
            ["Level", "Item ID", "Has Children", "Quantity", "Item Name"],  # Actual header at index 4
            ["1", "TOP_ASSY", "True", "1.0", "MAIN ASSY"],
            ["2", "SUB_PART", "False", "4.0", "SUB COMPONENT"],
        ]
        df = pd.DataFrame(matrix)
        parser = BOMTreeParser()
        header_idx = -1
        mapping = {}
        for i, row in df.iterrows():
            norm = [str(c).strip().lower() for c in row if c is not None]
            if "level" in norm and ("item id" in norm or "item_id" in norm):
                header_idx = i
                mapping = parser.detect_column_mapping(list(row))
                break

        assert header_idx == 4
        data_rows = df.iloc[header_idx + 1:].values.tolist()
        tree = parser._build_tree_from_matrix(data_rows, mapping)
        assert len(tree.roots) == 1
        assert tree.roots[0].item_id == "TOP_ASSY"
        assert len(tree.roots[0].children) == 1
