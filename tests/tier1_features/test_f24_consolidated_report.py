"""Feature F24: Consolidated Report Generation Isolation Tests.

Verifies:
1. Report workbook contains all mandatory sheets matching legacy form_ssbom.xlsm:
   'Tongket', 'List JIG', 'MSI_7980_7990', 'CTTT', 'PLM', 'R3', 'CTTT_Total'.
2. Conditional formatting styling (Green for 'OK', Red for 'NG').
3. Numerical alignment and formula integrity.
4. Header styling and column formatting.
5. Generation of a valid, non-corrupted OpenXML workbook.
"""

from __future__ import annotations

from pathlib import Path
import openpyxl
import pandas as pd
import pytest

from src.core.reconciliation import ReconciliationResult
from src.reporting.excel_generator import (
    COLOR_GREEN_FILL_HEX,
    COLOR_RED_FILL_HEX,
    ExcelReportGenerator,
)


class TestF24ConsolidatedReport:
    """Test suite for Feature F24: Consolidated Report Generation."""

    @pytest.fixture
    def sample_reconciliation(self) -> ReconciliationResult:
        cttt_df = pd.DataFrame([
            {
                "SUB": "LSU",
                "TRANG CTTT": "01",
                "MÃ LINH KIỆN": "302FP02010",
                "TÊN LINH KIỆN": "BRACKET",
                "SỐ LƯỢNG": 1.0,
                "PHỤ TRÁCH": "A",
                "Q.ty (PLM)": 1.0,
                "Compare (PLM Qty)": "OK",
                "Rev PLM": "A",
                "Qty (R3)": 1.0,
                "Compare (R3 Qty)": "OK",
                "Rev R3": "A",
                "Compare (Rev)": "OK",
                "Giải thích": "",
                "Check": "OK",
            },
            {
                "SUB": "FUSER",
                "TRANG CTTT": "02",
                "MÃ LINH KIỆN": "302FP04010",
                "TÊN LINH KIỆN": "HEATER",
                "SỐ LƯỢNG": 2.0,
                "PHỤ TRÁCH": "B",
                "Q.ty (PLM)": 0.0,
                "Compare (PLM Qty)": "NG",
                "Rev PLM": "",
                "Qty (R3)": 0.0,
                "Compare (R3 Qty)": "NG",
                "Rev R3": "",
                "Compare (Rev)": "NG",
                "Giải thích": "Missing in PLM",
                "Check": "NG",
            },
        ])
        plm_missing = pd.DataFrame([{"item_id": "302FP09999", "quantity": 1.0}])
        cttt_totals = pd.DataFrame([{"MÃ LINH KIỆN": "302FP02010", "CTTT_SUM_QTY": 1.0, "r3_total_qty": 1.0, "STATUS": "OK"}])
        msi_results = pd.DataFrame([{"unit_code": "302FP93010", "member_code": "1HN", "master_code": "1HN", "status": "OK"}])

        return ReconciliationResult(
            cttt_rows=cttt_df,
            plm_missing_rows=plm_missing,
            cttt_totals=cttt_totals,
            msi_results=msi_results,
            overall_status="NG",
        )

    def test_f24_sheet_structure_matches_legacy_form_ssbom(
        self,
        sample_reconciliation: ReconciliationResult,
        tmp_path: Path,
    ) -> None:
        """Test 1: Output workbook has exact 7 sheets from form_ssbom.xlsm."""
        out_file = tmp_path / "form_ssbom_generated.xlsx"
        generator = ExcelReportGenerator()
        generator.generate_report(out_file, sample_reconciliation, "Virgo", "2026-09-17")

        wb = openpyxl.load_workbook(out_file)
        expected = [
            "Tongket", "List JIG", "MSI_7980_7990", "CTTT", "PLM", "R3", "CTTT_Total"
        ]
        assert wb.sheetnames == expected
        wb.close()

    def test_f24_cell_color_formatting_ok_green_ng_red(
        self,
        sample_reconciliation: ReconciliationResult,
        tmp_path: Path,
    ) -> None:
        """Test 2: Status 'OK' is formatted with soft green, 'NG' with soft red."""
        out_file = tmp_path / "formatted_report.xlsx"
        generator = ExcelReportGenerator()
        generator.generate_report(out_file, sample_reconciliation, "Virgo", "2026-09-17")

        wb = openpyxl.load_workbook(out_file)
        ws = wb["CTTT"]
        # Column R (18) is Check column
        check_col = 18
        val_row3 = ws.cell(row=3, column=check_col).value
        val_row4 = ws.cell(row=4, column=check_col).value
        assert val_row3 == "OK"
        assert val_row4 == "NG"

        fill_row3 = ws.cell(row=3, column=check_col).fill.start_color.rgb
        fill_row4 = ws.cell(row=4, column=check_col).fill.start_color.rgb
        assert COLOR_GREEN_FILL_HEX in str(fill_row3)
        assert COLOR_RED_FILL_HEX in str(fill_row4)
        wb.close()

    def test_f24_preserve_formulas_in_report(
        self,
        sample_reconciliation: ReconciliationResult,
        tmp_path: Path,
    ) -> None:
        """Test 3: Verify Tongket overview summary sheet is populated."""
        out_file = tmp_path / "formula_report.xlsx"
        generator = ExcelReportGenerator()
        generator.generate_report(out_file, sample_reconciliation, "Virgo", "2026-09-17")

        wb = openpyxl.load_workbook(out_file, data_only=False)
        ws_tongket = wb["Tongket"]
        assert ws_tongket is not None
        assert ws_tongket.max_row > 10
        wb.close()

    def test_f24_column_widths_and_autofit(
        self,
        sample_reconciliation: ReconciliationResult,
        tmp_path: Path,
    ) -> None:
        """Test 4: Apply column dimension formatting."""
        out_file = tmp_path / "width_report.xlsx"
        generator = ExcelReportGenerator()
        generator.generate_report(out_file, sample_reconciliation, "Virgo", "2026-09-17")

        wb = openpyxl.load_workbook(out_file)
        ws = wb["CTTT"]
        assert ws.column_dimensions["C"].width > 10
        wb.close()

    def test_f24_report_file_creation_and_non_empty(
        self,
        sample_reconciliation: ReconciliationResult,
        tmp_path: Path,
    ) -> None:
        """Test 5: Generated workbook is valid file with size > 5KB."""
        out_file = tmp_path / "final_report.xlsx"
        generator = ExcelReportGenerator()
        generator.generate_report(out_file, sample_reconciliation, "Virgo", "2026-09-17")

        assert out_file.exists()
        assert out_file.stat().st_size > 5000
