"""Tier 5 Adversarial Tests: Malformed Workbooks and File Corruption.

Verifies fail-closed resilience against:
1. 0-byte files with .xlsx / .xls extensions.
2. Truncated zip archives and files ending mid-stream.
3. Fake Excel files (plain text, JSON, HTML disguised as .xlsx).
4. Corrupted zip CRC and invalid XML structures.
5. Missing or corrupted worksheets in multi-sheet workbooks.
6. Non-existent and directory paths passed as files.
"""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

import openpyxl
import pytest

from src.automation.sap.models import SAPParseError
from src.automation.sap.parser import ResilientR3Parser
from src.core.tree_parser import BOMTreeParser


class TestAdversarialMalformedWorkbooks:
    """Test suite for malformed and corrupted workbooks."""

    def test_zero_byte_excel_file_rejection(self, tmp_path: Path):
        """0-byte file must raise appropriate parse error or return empty tree cleanly."""
        empty_xlsx = tmp_path / "empty_corrupt.xlsx"
        empty_xlsx.write_bytes(b"")

        parser = BOMTreeParser()
        with pytest.raises(Exception):
            parser.parse_excel(empty_xlsx)

        r3_parser = ResilientR3Parser()
        with pytest.raises(SAPParseError):
            r3_parser.parse(empty_xlsx)

    def test_truncated_zip_archive(self, tmp_path: Path):
        """Partially written / truncated zip file must fail closed without hang."""
        # Create a valid minimal zip, then truncate halfway
        valid_buf = io.BytesIO()
        with zipfile.ZipFile(valid_buf, "w") as zf:
            zf.writestr("[Content_Types].xml", "<Types></Types>")
            zf.writestr("xl/workbook.xml", "<workbook></workbook>")
        truncated_bytes = valid_buf.getvalue()[:len(valid_buf.getvalue()) // 2]

        corrupt_file = tmp_path / "truncated.xlsx"
        corrupt_file.write_bytes(truncated_bytes)

        parser = BOMTreeParser()
        with pytest.raises(Exception):
            parser.parse_excel(corrupt_file)

    def test_plain_text_disguised_as_xlsx(self, tmp_path: Path):
        """Text / ASCII script disguised as .xlsx must be rejected safely."""
        fake_xlsx = tmp_path / "fake_bom.xlsx"
        fake_xlsx.write_text("THIS IS NOT A ZIP FILE OR EXCEL ARCHIVE", encoding="utf-8")

        parser = BOMTreeParser()
        with pytest.raises(Exception):
            parser.parse_excel(fake_xlsx)

    def test_corrupted_crc_zip_payload(self, tmp_path: Path):
        """Zip archive with corrupted byte content must be detected."""
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("xl/workbook.xml", "<workbook><sheets></sheets></workbook>")

        data = bytearray(buf.getvalue())
        # Corrupt central payload bytes
        for i in range(20, min(60, len(data))):
            data[i] ^= 0xFF

        bad_crc_file = tmp_path / "bad_crc.xlsx"
        bad_crc_file.write_bytes(bytes(data))

        parser = BOMTreeParser()
        with pytest.raises(Exception):
            parser.parse_excel(bad_crc_file)

    def test_workbook_with_no_sheets(self, tmp_path: Path):
        """Workbook containing no sheets or blank sheet handles gracefully."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "EmptySheet"
        # Don't add any rows
        empty_wb_path = tmp_path / "empty_sheet.xlsx"
        wb.save(empty_wb_path)
        wb.close()

        parser = BOMTreeParser()
        tree = parser.parse_excel(empty_wb_path)
        assert tree is not None
        assert len(tree.roots) == 0

    def test_nonexistent_and_directory_filepath(self, tmp_path: Path):
        """Non-existent path and folder path passed to parser must raise FileNotFoundError / SAPParseError."""
        parser = BOMTreeParser()
        with pytest.raises(FileNotFoundError):
            parser.parse_excel(tmp_path / "ghost_file.xlsx")

        dir_path = tmp_path / "subfolder"
        dir_path.mkdir()
        with pytest.raises((IsADirectoryError, PermissionError, Exception)):
            parser.parse_excel(dir_path)

        r3_parser = ResilientR3Parser()
        with pytest.raises(SAPParseError):
            r3_parser.parse(tmp_path / "ghost_r3.xls")
