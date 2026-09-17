"""Feature F21: Dynamic R3 Header Resolution Isolation Tests.

Verifies:
1. Dynamic detection of Part Code, Quantity, and Revision across ALV layout variants.
2. Graceful fallback when Revision column is omitted from SAP layout.
3. Parsing Vietnamese and alternate header translations.
4. Parsing varied numeric quantity formats (European comma decimals, trailing minus).
5. Filtering out decorative lines, borders (e.g. '-------'), and repeated header artifacts.
"""

from pathlib import Path
import pytest

from src.automation.sap.parser import ResilientR3Parser, parse_r3_cs12_file


class TestF21DynamicR3Header:
    """Test suite for Feature F21: Dynamic R3 Header Resolution."""

    def test_f21_parse_with_revlev_column(self, tmp_path: Path):
        """Test 1: Layout with RevLev column present."""
        html_content = """
        <table>
            <tr><th>Component</th><th>Quantity</th><th>RevLev</th></tr>
            <tr><td>PART_100</td><td>5.0</td><td>02</td></tr>
        </table>
        """
        f = tmp_path / "r3_with_rev.xls"
        f.write_text(html_content, encoding="utf-8")

        parser = ResilientR3Parser()
        df = parser.parse(f)

        assert len(df) == 1
        assert df.iloc[0]["part_code"] == "PART_100"
        assert df.iloc[0]["quantity"] == 5.0
        # Revision string is retained
        assert str(df.iloc[0]["rev_r3"]) in ("02", "2")

    def test_f21_parse_without_revlev_column(self, tmp_path: Path):
        """Test 2: Layout without RevLev column present fills rev_r3 with empty string."""
        html_content = """
        <table>
            <tr><th>Material</th><th>Component Qty</th></tr>
            <tr><td>PART_200</td><td>10.0</td></tr>
        </table>
        """
        f = tmp_path / "r3_no_rev.xls"
        f.write_text(html_content, encoding="utf-8")

        parser = ResilientR3Parser()
        df = parser.parse(f)

        assert len(df) == 1
        assert df.iloc[0]["part_code"] == "PART_200"
        assert df.iloc[0]["quantity"] == 10.0
        assert df.iloc[0]["rev_r3"] == ""

    def test_f21_vietnamese_header_variants(self, tmp_path: Path):
        """Test 3: Detect Vietnamese column headers in TSV exports."""
        tsv_content = (
            "MÃ LINH KIỆN\tSỐ LƯỢNG\tREV\n"
            "302FP02010\t1.0\tA\n"
        )
        f = tmp_path / "r3_vn.txt"
        f.write_text(tsv_content, encoding="utf-8")

        df = parse_r3_cs12_file(f)
        assert len(df) == 1
        assert df.iloc[0]["part_code"] == "302FP02010"
        assert df.iloc[0]["quantity"] == 1.0
        assert df.iloc[0]["rev_r3"] == "A"

    def test_f21_quantity_formats_and_decimal_notations(self, tmp_path: Path):
        """Test 4: Handle German decimal notation, standard thousands separators, and trailing minus."""
        tsv_content = (
            "Component\tQuantity\n"
            "P1\t1.250,50\n"
            "P2\t1,250.50\n"
            "P3\t10-\n"
        )
        f = tmp_path / "r3_nums.txt"
        f.write_text(tsv_content, encoding="utf-8")

        df = parse_r3_cs12_file(f)
        assert len(df) == 3
        assert df.iloc[0]["quantity"] == 1250.5
        assert df.iloc[1]["quantity"] == 1250.5
        assert df.iloc[2]["quantity"] == -10.0

    def test_f21_filter_decorative_borders_and_empty_rows(self, tmp_path: Path):
        """Test 5: Filter out decorative border rows (e.g. '-------') and repeated headers."""
        tsv_content = (
            "Component\tQuantity\tRev\n"
            "--------------------\t-----\t---\n"
            "302FP93010\t1.0\t01\n"
            "====================\t=====\t===\n"
            "Component\tQuantity\tRev\n"
            "302FP93020\t2.0\t02\n"
        )
        f = tmp_path / "r3_decorations.txt"
        f.write_text(tsv_content, encoding="utf-8")

        df = parse_r3_cs12_file(f)
        assert len(df) == 2
        assert set(df["part_code"]) == {"302FP93010", "302FP93020"}
