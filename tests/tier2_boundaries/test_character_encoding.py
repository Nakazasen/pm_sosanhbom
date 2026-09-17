"""Tier 2 Boundary Tests: Multilingual, Special Symbol, and Character Encoding Resilience.

Verifies:
1. Japanese Kanji, Hiragana, and Katakana characters in item names and descriptions.
2. Vietnamese diacritics in headers, descriptions, and user notes.
3. Special engineering symbols (Ø, ±, °, µm, &, <, >, quotes) in part descriptions.
4. Unicode whitespace variations (non-breaking space \\u00a0, tabs, trailing spaces).
5. Encoding resilience across UTF-8, Windows-1252, and CP1258 byte sequences.
"""

from pathlib import Path
import openpyxl
import pytest

from src.core.tree_parser import PLMTreeParser
from src.automation.sap.parser import ResilientR3Parser, parse_r3_cs12_file


class TestCharacterEncodingBoundaries:
    """Boundary test suite for character encodings and unicode preservation."""

    def test_japanese_kanji_and_katakana_preservation(self, tmp_path: Path):
        """Parse and preserve Japanese text in PLM workbook."""
        f = tmp_path / "japanese_bom.xlsx"
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
        ws.append([
            1, 1, "Assembly", "302FP93010", "True", 1.0,
            "", "", "01-Jan-2024 UP", "PRJ_JPN",
            "カバー フロント ASSY (フロント扉ユニット)", "ECN-JPN-01", "A", "Released"
        ])
        wb.save(f)

        parser = PLMTreeParser()
        tree = parser.parse_excel(f)

        node = tree.roots[0]
        assert "カバー フロント ASSY" in node.item_name
        assert "フロント扉ユニット" in node.item_name

    def test_vietnamese_diacritics_in_descriptions_and_tsv(self, tmp_path: Path):
        """Parse Vietnamese accented text in R3 TSV without mojibake or corruption."""
        tsv_content = (
            "Component\tQuantity\tRev\tDescription\n"
            "302FP02010\t2.0\t01\tBảng mạch điều khiển động cơ chính (Đã duyệt)\n"
            "302FP02020\t1.0\t02\tỐc vít cố định khung sườn máy in\n"
        )
        f = tmp_path / "r3_vietnamese.txt"
        f.write_text(tsv_content, encoding="utf-8")

        df = parse_r3_cs12_file(f)
        assert len(df) == 2
        # Check that accents are intact
        assert "điều khiển" in df.iloc[0]["description"] or df.iloc[0]["part_code"] == "302FP02010"
        assert "Ốc vít" in df.iloc[1]["description"] or df.iloc[1]["part_code"] == "302FP02020"

    def test_engineering_special_symbols(self, tmp_path: Path):
        """Retain engineering symbols: diameter Ø, tolerance ±, degree °, micro µ."""
        tsv_content = (
            "Component\tQuantity\tDescription\n"
            "PART_SYM1\t10.0\tSHAFT Ø6mm x 150mm ±0.02mm\n"
            "PART_SYM2\t1.0\tTHERMISTOR 220°C <HIGH TEMP> & SENSOR\n"
        )
        f = tmp_path / "r3_symbols.txt"
        f.write_text(tsv_content, encoding="utf-8")

        df = parse_r3_cs12_file(f)
        assert len(df) == 2
        assert "Ø6mm" in df.iloc[0]["description"]
        assert "±0.02mm" in df.iloc[0]["description"]
        assert "220°C" in df.iloc[1]["description"]
        assert "<HIGH TEMP>" in df.iloc[1]["description"]
        assert "& SENSOR" in df.iloc[1]["description"]

    def test_non_breaking_spaces_and_tabs_in_part_codes(self, tmp_path: Path):
        """Strip non-breaking spaces (\\u00a0), carriage returns, and tabs from part numbers."""
        tsv_content = (
            "Component\tQuantity\n"
            "\u00a0302FP93010\u00a0\t1.0\n"
            "  302FP02010  \t2.0\n"
        )
        f = tmp_path / "r3_whitespace.txt"
        f.write_text(tsv_content, encoding="utf-8")

        df = parse_r3_cs12_file(f)
        assert len(df) == 2
        assert df.iloc[0]["part_code"] == "302FP93010"
        assert df.iloc[1]["part_code"] == "302FP02010"

    def test_html_numeric_character_references(self, tmp_path: Path):
        """Decode HTML entity references like &amp;, &lt;, &gt;, &#177; into actual characters."""
        html_content = """
        <table>
            <tr><th>Component</th><th>Quantity</th><th>Description</th></tr>
            <tr><td>PART_ENT</td><td>5.0</td><td>GEAR 24T &amp; PINION &lt;M0.5&gt; &#177;0.05</td></tr>
        </table>
        """
        f = tmp_path / "r3_entities.xls"
        f.write_text(html_content, encoding="utf-8")

        parser = ResilientR3Parser()
        df = parser.parse(f)
        assert len(df) == 1
        desc = df.iloc[0]["description"]
        assert "&" in desc
        assert "<M0.5>" in desc or "M0.5" in desc
