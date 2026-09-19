"""Unit tests for src/core/jig_manager.py.

Verifies:
1. Canonical 15 machine series matching Sheet 'List JIG' V3:V17.
2. JIG catalog loading across series from master Excel workbook.
3. Sheet population preserving formulas, cell formatting, merged cells, and column widths (A3:G50).
4. Sheet clearing logic (A3:H50 unmerge & clear).
5. 4M assessment checklist (Man, Machine, Material, Method) write and read-back (V36:V41).
6. KTSX sign-off confirmation recording and parsing.
7. Resilience against missing files, corrupt paths, and edge cases.
"""

from __future__ import annotations

from pathlib import Path
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import pytest

from src.core.jig_manager import (
    STANDARD_JIG_SERIES,
    Assessment4M,
    Decision4M,
    JIGManager,
    Presence4M,
)


@pytest.fixture
def sample_master_jig_file(tmp_path: Path) -> Path:
    """Create a realistic multi-sheet JIG master Excel workbook."""
    file_path = tmp_path / "List JIG thay doi, khi bo sung ma hang.xlsx"
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    thin_border = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000"),
    )
    yellow_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
    bold_font = Font(name="Arial", size=10, bold=True)

    # Populate 3 sample series sheets: Virgo, Polaris, 6thA4
    for series in ["Virgo", "Polaris", "6thA4"]:
        ws = wb.create_sheet(title=series)

        # Header rows 1-2
        ws["A1"] = f"DANH MUC JIG CHO DONG MAY {series.upper()}"
        ws["A2"] = "STT"
        ws["B2"] = "TEN JIG"
        ws["C2"] = "MA JIG"
        ws["D2"] = "VI TRI"
        ws["E2"] = "SO LUONG"
        ws["F2"] = "DON GIA"
        ws["G2"] = "THANH TIEN"

        # Rows 3 to 10 with realistic data and formulas
        for r in range(3, 8):
            ws.cell(row=r, column=1, value=r - 2)
            ws.cell(row=r, column=2, value=f"JIG Lap Rap {series} #{r-2}")
            ws.cell(row=r, column=3, value=f"JIG-{series[:3].upper()}-00{r-2}")
            ws.cell(row=r, column=4, value="Công đoạn 1")
            ws.cell(row=r, column=5, value=r)
            ws.cell(row=r, column=6, value=150000.0)
            # Formula preserving test: Col G = Col E * Col F
            ws.cell(row=r, column=7, value=f"=E{r}*F{r}")

            # Apply styling to verify style preservation
            cell = ws.cell(row=r, column=2)
            cell.font = bold_font
            cell.fill = yellow_fill
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="left", vertical="center")

        # Merge cells in row 8 (e.g. A8:B8 note)
        ws.cell(row=8, column=1, value="Ghi chú chung")
        ws.merge_cells("A8:B8")

        # Column widths
        ws.column_dimensions["B"].width = 30.0
        ws.column_dimensions["C"].width = 20.0

    wb.save(file_path)
    wb.close()
    return file_path


@pytest.fixture
def target_ssbom_file(tmp_path: Path) -> Path:
    """Create a mock target workbook with an empty 'List JIG' sheet."""
    target_path = tmp_path / "BOM_Virgo_Test.xlsm"
    wb = openpyxl.Workbook()
    ws_jig = wb.active
    ws_jig.title = "List JIG"

    # Pre-populate V3:V17 and V24
    for idx, s_name in enumerate(STANDARD_JIG_SERIES, start=3):
        ws_jig[f"V{idx}"] = s_name

    ws_jig["V24"] = "dummy_path_to_master.xlsx"
    ws_jig["V36"] = "Hãy xác nhận"
    ws_jig["V37"] = "Hãy xác nhận"
    ws_jig["V39"] = "Hãy xác nhận"
    ws_jig["V40"] = "Hãy xác nhận"

    wb.save(target_path)
    wb.close()
    return target_path


class TestJIGManager:
    """Comprehensive test suite for JIGManager."""

    def test_canonical_15_series(self) -> None:
        """Verify the 15 standard machine series match List JIG specification exactly."""
        expected_15 = [
            "6thA4",
            "Polaris",
            "6thA3",
            "Brazil",
            "Libra",
            "Libra2",
            "Mebius",
            "Kairos",
            "Virgo",
            "Iris",
            "Sirius",
            "EH Iris",
            "Spica",
            "DF Iris",
            "Pictor",
        ]
        assert len(STANDARD_JIG_SERIES) == 15
        assert STANDARD_JIG_SERIES == expected_15

        manager = JIGManager()
        supported = manager.get_supported_series()
        assert len(supported) == 15
        assert supported == expected_15

    def test_is_supported_series_casing_and_whitespace(self) -> None:
        """Verify series matching is case-insensitive and trims whitespace."""
        manager = JIGManager()
        assert manager.is_supported_series("virgo") is True
        assert manager.is_supported_series("  Polaris  ") is True
        assert manager.is_supported_series("6THA4") is True
        assert manager.is_supported_series("EH IRIS") is True
        assert manager.is_supported_series("UnknownModel") is False
        assert manager.is_supported_series("") is False

    def test_load_series_catalog_as_dataframe(
        self, sample_master_jig_file: Path
    ) -> None:
        """Verify load_series_catalog loads rows 3:50 as a DataFrame."""
        manager = JIGManager(master_catalog_path=sample_master_jig_file)
        df = manager.load_series_catalog("Virgo")

        assert not df.empty
        # Verify first data row corresponds to row 3 of sheet
        assert "TEN JIG" in df.columns or df.shape[1] >= 6
        assert len(df) >= 4

    def test_load_series_catalog_missing_file_or_sheet(
        self, tmp_path: Path
    ) -> None:
        """Verify load_series_catalog gracefully returns empty DataFrame on invalid input."""
        manager = JIGManager(master_catalog_path=tmp_path / "non_existent.xlsx")
        df1 = manager.load_series_catalog("Virgo")
        assert isinstance(df1, pd.DataFrame)
        assert df1.empty

        # Valid file, missing sheet
        empty_file = tmp_path / "empty.xlsx"
        wb = openpyxl.Workbook()
        wb.save(empty_file)
        wb.close()

        manager2 = JIGManager(master_catalog_path=empty_file)
        df2 = manager2.load_series_catalog("NonExistentSeries")
        assert df2.empty

    def test_populate_jig_sheet_preserves_formulas(
        self, sample_master_jig_file: Path, target_ssbom_file: Path
    ) -> None:
        """Verify populating JIG sheet preserves Excel formulas verbatim (=E3*F3)."""
        manager = JIGManager(master_catalog_path=sample_master_jig_file)
        success = manager.populate_jig_sheet(
            target_workbook_path=target_ssbom_file,
            series_name="Virgo",
        )
        assert success is True

        # Open target workbook with data_only=False to verify raw formula
        wb = openpyxl.load_workbook(target_ssbom_file, data_only=False)
        ws = wb["List JIG"]

        # Cell G3 should be the exact formula '=E3*F3'
        assert ws["G3"].value == "=E3*F3"
        assert ws["G4"].value == "=E4*F4"
        assert ws["B3"].value == "JIG Lap Rap Virgo #1"
        assert ws["C3"].value == "JIG-VIR-001"

        wb.close()

    def test_populate_jig_sheet_preserves_styles_and_merges(
        self, sample_master_jig_file: Path, target_ssbom_file: Path
    ) -> None:
        """Verify populating JIG sheet replicates fonts, fills, borders, and merged ranges."""
        manager = JIGManager(master_catalog_path=sample_master_jig_file)
        manager.populate_jig_sheet(
            target_workbook_path=target_ssbom_file,
            series_name="Virgo",
        )

        wb = openpyxl.load_workbook(target_ssbom_file)
        ws = wb["List JIG"]

        # Check B3 styling
        b3 = ws["B3"]
        assert b3.font is not None and b3.font.bold is True
        assert b3.fill is not None and b3.fill.fill_type == "solid"
        assert b3.border is not None

        # Check merged cell range A8:B8
        merged_str = [str(r) for r in ws.merged_cells.ranges]
        assert any("A8:B8" in m for m in merged_str)

        # Check column width
        assert ws.column_dimensions["B"].width == 30.0

        wb.close()

    def test_populate_jig_sheet_unmerges_and_clears_prior_content(
        self, sample_master_jig_file: Path, target_ssbom_file: Path
    ) -> None:
        """Verify existing contents and merged cells in A3:H50 are cleared before new load."""
        # Pre-seed target with old data and old merged cell
        wb = openpyxl.load_workbook(target_ssbom_file)
        ws = wb["List JIG"]
        ws["A3"] = "OLD DATA"
        ws["B3"] = "OLD DATA"
        ws.merge_cells("A3:B3")
        wb.save(target_ssbom_file)
        wb.close()

        manager = JIGManager(master_catalog_path=sample_master_jig_file)
        manager.populate_jig_sheet(
            target_workbook_path=target_ssbom_file,
            series_name="Polaris",
        )

        wb = openpyxl.load_workbook(target_ssbom_file)
        ws = wb["List JIG"]
        # Old merged cell A3:B3 must be unmerged
        merged_str = [str(r) for r in ws.merged_cells.ranges]
        assert not any("A3:B3" in m for m in merged_str)
        # New data must be present
        assert "Polaris" in str(ws["B3"].value)
        wb.close()

    def test_clear_jig_sheet(self, target_ssbom_file: Path) -> None:
        """Verify clear_jig_sheet completely wipes A3:H50 contents and unmerges."""
        wb = openpyxl.load_workbook(target_ssbom_file)
        ws = wb["List JIG"]
        ws["A3"] = "DATA 1"
        ws["C5"] = "DATA 2"
        ws.merge_cells("C5:D5")
        wb.save(target_ssbom_file)
        wb.close()

        manager = JIGManager()
        success = manager.clear_jig_sheet(target_ssbom_file)
        assert success is True

        wb = openpyxl.load_workbook(target_ssbom_file)
        ws = wb["List JIG"]
        assert ws["A3"].value is None
        assert ws["C5"].value is None
        assert not any("C5:D5" in str(r) for r in ws.merged_cells.ranges)
        wb.close()

    def test_write_and_read_4m_assessment_with_changes(
        self, target_ssbom_file: Path
    ) -> None:
        """Verify writing and reading 4M assessment with changes (Có, OK)."""
        manager = JIGManager()

        assessment = Assessment4M(
            man_changed=Presence4M.NO,
            machine_changed=Presence4M.YES,  # JIG modified
            material_changed=Presence4M.NO,
            method_changed=Presence4M.NO,
            overall_evaluation=Decision4M.OK,
            ktsx_confirmed=True,
            ktsx_confirmer="Nguyen Van B",
            confirmation_date="19/09/2026",
            notes="Bo sung JIG han moi",
        )

        write_ok = manager.write_4m_assessment(target_ssbom_file, assessment)
        assert write_ok is True

        wb = openpyxl.load_workbook(target_ssbom_file, data_only=True)
        ws = wb["List JIG"]
        assert ws["V36"].value == "Hãy xác nhận"
        assert ws["V37"].value == "Có"
        assert ws["V39"].value == "Hãy xác nhận"
        assert ws["V40"].value == "OK"
        assert "Nguyen Van B" in str(ws["V41"].value)
        assert "19/09/2026" in str(ws["V41"].value)
        wb.close()

        # Read back
        read_back = manager.read_4m_assessment(target_ssbom_file)
        assert read_back is not None
        assert read_back.has_any_change() is True
        assert read_back.overall_evaluation == Decision4M.OK
        assert read_back.ktsx_confirmed is True
        assert read_back.ktsx_confirmer == "Nguyen Van B"

    def test_write_and_read_4m_assessment_no_change(
        self, target_ssbom_file: Path
    ) -> None:
        """Verify 4M assessment with no change (Không, OK)."""
        manager = JIGManager()

        assessment = Assessment4M(
            man_changed=Presence4M.NO,
            machine_changed=Presence4M.NO,
            material_changed=Presence4M.NO,
            method_changed=Presence4M.NO,
            overall_evaluation=Decision4M.OK,
            ktsx_confirmed=False,
        )

        manager.write_4m_assessment(target_ssbom_file, assessment)

        wb = openpyxl.load_workbook(target_ssbom_file, data_only=True)
        ws = wb["List JIG"]
        assert ws["V37"].value == "Không"
        assert ws["V40"].value == "OK"
        assert ws["V41"].value == "Chưa xác nhận KTSX"
        wb.close()

        read_back = manager.read_4m_assessment(target_ssbom_file)
        assert read_back is not None
        assert read_back.has_any_change() is False
        assert read_back.overall_evaluation == Decision4M.OK
        assert read_back.ktsx_confirmed is False

    def test_extract_master_path_from_sheet(
        self, target_ssbom_file: Path
    ) -> None:
        """Verify extracting master catalog path from cell V24."""
        manager = JIGManager()
        extracted = manager.extract_master_path_from_sheet(target_ssbom_file)
        assert extracted is not None
        assert "dummy_path_to_master.xlsx" in str(extracted)
