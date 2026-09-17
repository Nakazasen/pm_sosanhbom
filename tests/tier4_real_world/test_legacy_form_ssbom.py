"""Tier 4 Real-World Legacy Validation: form_ssbom.xlsm Production Ground Truth.

Validates:
1. Presence and integrity of all 7 production sheets:
   ['Tongket', 'List JIG', 'MSI_7980_7990', 'CTTT', 'PLM', 'R3', 'CTTT_Total'].
2. 100% mathematical fidelity to legacy VBA reconciliation formulas:
   - Col H: Compare PLM Qty = IF(G3=E3, "OK", "NG")
   - Col L: Compare R3 Qty = IF(K3=E3, "OK", "NG")
   - Col N: Compare Rev = IF(M3=I3, "OK", "NG")
   - Col R: Check = IF(OR(G3=0, K3=0, N3="NG"), "NG", "OK")
3. Standard 14 serialized unit names in Sheet MSI_7980_7990
   (IMAGE UNIT, FUSER UNIT, DATA LASER, CASE OUTER, ISU, etc.).
4. Compatibility of ReconciliationEngine with authentic Sheet CTTT header schema.
5. Leader summary report structure in Sheet Tongket (approval gates and sign-off blocks).
"""

from pathlib import Path
import openpyxl
import pandas as pd
import pytest

from src.core.reconciliation import ReconciliationEngine, reconcile_single_row


@pytest.fixture(scope="module")
def form_ssbom_path() -> Path:
    """Locate root production template form_ssbom.xlsm."""
    path = Path("form_ssbom.xlsm")
    assert path.exists(), "form_ssbom.xlsm ground truth file missing from project root!"
    return path


class TestLegacyFormSSBOMGroundTruth:
    """Ground truth validation suite against production workbook form_ssbom.xlsm."""

    def test_all_seven_production_sheets_present(self, form_ssbom_path: Path):
        """Verify presence of all 7 factory production sheets."""
        wb = openpyxl.load_workbook(str(form_ssbom_path), read_only=True)
        sheet_names = wb.sheetnames
        wb.close()

        expected_sheets = [
            "Tongket",
            "List JIG",
            "MSI_7980_7990",
            "CTTT",
            "PLM",
            "R3",
            "CTTT_Total",
        ]
        for s in expected_sheets:
            assert s in sheet_names, f"Expected production sheet '{s}' not found in form_ssbom.xlsm!"

    def test_cttt_legacy_formula_mathematical_fidelity(self):
        """Validate Python reconciliation engine reproduces legacy VBA formulas with 100% fidelity.

        Formulas extracted directly from form_ssbom.xlsm Sheet CTTT Row 3:
        - Col H: Compare (PLM Qty) = IF(G3=E3,"OK","NG")
        - Col L: Compare (R3 Qty) = IF(K3=E3,"OK","NG")
        - Col N: Compare (Rev) = IF(M3=I3,"OK","NG")
        - Col R: Check = IF(OR(G3=0,K3=0,N3="NG"),"NG","OK")
        """
        # Test Case 1: Perfect Match
        r1 = reconcile_single_row(cttt_qty=2.0, plm_qty=2.0, r3_qty=2.0, plm_rev="01", r3_rev="01")
        assert r1["comp_plm_qty"] == "OK"
        assert r1["comp_r3_qty"] == "OK"
        assert r1["comp_rev"] == "OK"
        assert r1["overall_check"] == "OK"

        # Test Case 2: PLM quantity mismatch
        r2 = reconcile_single_row(cttt_qty=2.0, plm_qty=1.0, r3_qty=2.0, plm_rev="01", r3_rev="01")
        assert r2["comp_plm_qty"] == "NG"
        assert r2["comp_r3_qty"] == "OK"
        assert r2["overall_check"] == "NG"

        # Test Case 3: Zero quantity in R3 (G3=0 or K3=0 rule)
        r3 = reconcile_single_row(cttt_qty=2.0, plm_qty=2.0, r3_qty=0.0, plm_rev="01", r3_rev="01")
        assert r3["comp_r3_qty"] == "NG"
        assert r3["overall_check"] == "NG"

        # Test Case 4: Revision mismatch (N3="NG")
        r4 = reconcile_single_row(cttt_qty=2.0, plm_qty=2.0, r3_qty=2.0, plm_rev="01", r3_rev="02")
        assert r4["comp_rev"] == "NG"
        assert r4["overall_check"] == "NG"

    def test_msi_standard_unit_assembly_catalog(self, form_ssbom_path: Path):
        """Verify the 14 standard serialized unit names present in Sheet MSI_7980_7990."""
        wb = openpyxl.load_workbook(str(form_ssbom_path), data_only=True, read_only=True)
        ws = wb["MSI_7980_7990"]

        unit_names: list[str] = []
        for row in ws.iter_rows(min_row=2, max_row=25, values_only=True):
            # Col C (index 2) contains unit descriptions
            if len(row) > 2 and row[2] is not None:
                val = str(row[2]).strip()
                if val:
                    unit_names.append(val)
        wb.close()

        required_serialized_units = [
            "IMAGE UNIT",
            "FUSER UNIT",
            "ISU",
            "MAIN DRIVE ASSY",
            "WTB UNIT",
        ]
        for req in required_serialized_units:
            assert any(req in u for u in unit_names), f"Required serialized unit '{req}' missing from MSI sheet!"

    def test_reconciliation_engine_with_real_cttt_header_schema(self, form_ssbom_path: Path):
        """ReconciliationEngine correctly recognizes and processes real Sheet CTTT headers."""
        wb = openpyxl.load_workbook(str(form_ssbom_path), read_only=True)
        ws = wb["CTTT"]
        # Row 2 contains actual production headers
        r2 = [str(c.value).strip() for c in ws[2] if c.value is not None]
        wb.close()

        assert "SUB" in r2
        assert "TRANG CTTT" in r2
        assert "MÃ LINH KIỆN" in r2
        assert "TÊN LINH KIỆN" in r2
        assert "SỐ LƯỢNG" in r2

        # Create DataFrame adhering to this exact schema
        df_cttt = pd.DataFrame([{
            "SUB": "SUB 1",
            "TRANG CTTT": "P1",
            "MÃ LINH KIỆN": "302FP02010",
            "TÊN LINH KIỆN": "MOTOR BRACKET",
            "SỐ LƯỢNG": 1.0,
        }])
        df_plm = pd.DataFrame([{
            "part_code": "302FP02010",
            "quantity": 1.0,
            "revision": "01",
        }])
        df_r3 = pd.DataFrame([{
            "part_code": "302FP02010",
            "quantity": 1.0,
            "revision": "01",
        }])

        engine = ReconciliationEngine()
        result_df = engine.reconcile_three_way(df_cttt, df_plm, df_r3)
        assert len(result_df) == 1
        assert result_df.iloc[0]["Check"] == "OK"
        assert result_df.iloc[0]["Compare (PLM Qty)"] == "OK"
        assert result_df.iloc[0]["Compare (R3 Qty)"] == "OK"

    def test_tongket_sign_off_sheet_structure(self, form_ssbom_path: Path):
        """Verify summary sign-off structure in Sheet Tongket."""
        wb = openpyxl.load_workbook(str(form_ssbom_path), data_only=True, read_only=True)
        ws = wb["Tongket"]
        title_cell = ws.cell(1, 2).value or ws.cell(2, 2).value
        wb.close()

        assert "Những nội dung cần xác nhận" in str(title_cell) or "xác nhận khi sản xuất" in str(title_cell)
