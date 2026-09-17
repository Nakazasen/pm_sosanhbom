"""Feature F23: Member Input Workspace Isolation Tests.

Verifies:
1. Member CTTT entry validation and addition via MemberWorkspaceView.
2. Member MSI 3-character fixed code format validation.
3. Member preliminary self-check execution prior to submission.
4. LCP Label 7980/7990 input validation.
5. Packaging member submission into structured records.
"""

from __future__ import annotations

from pathlib import Path
import pandas as pd
import pytest
from PyQt6.QtWidgets import QApplication, QMessageBox

from src.core.msi_engine import evaluate_msi_branch
from src.core.reconciliation import ReconciliationEngine
from src.gui.member_view import MemberWorkspaceView


@pytest.fixture
def qapp() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


class TestF23MemberWorkspace:
    """Test suite for Feature F23: Member Input Workspace."""

    def test_f23_member_cttt_input_validation(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """Test 1: Validate member work instruction input fields via MemberWorkspaceView."""
        view = MemberWorkspaceView(base_dir=tmp_path)
        view.clear_cttt_table()

        row_idx = view.add_cttt_row(
            page="01",
            part_code="302FP02010",
            part_name="MOTOR BRACKET",
            quantity=1.0,
            explanation="Normal part",
        )
        assert view.cttt_table.rowCount() == 1
        assert view.cttt_table.item(row_idx, 1).text() == "302FP02010"

    def test_f23_member_msi_code_entry(self) -> None:
        """Test 2: Validate 3-character MSI fixed code constraint via evaluate_msi_branch."""
        res_valid = evaluate_msi_branch(
            in_plm=True,
            member_code="2NL",
            master_code="2NL",
            member_service="-",
            master_service="",
        )
        assert res_valid["status"] == "OK"

        res_invalid = evaluate_msi_branch(
            in_plm=True,
            member_code="2NL",
            master_code="XYZ",
            member_service="-",
            master_service="",
        )
        assert res_invalid["status"] == "NG"
        assert res_invalid["branch"] == 8

    def test_f23_preliminary_self_check_execution(
        self,
        qapp: QApplication,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """Test 3: Member runs self-check comparing CTTT parts against local PLM sub-unit."""
        monkeypatch.setattr(QMessageBox, "information", lambda *args, **kwargs: None)
        view = MemberWorkspaceView(base_dir=tmp_path)
        view.clear_cttt_table()
        view.add_cttt_row(page="01", part_code="302FP02010", quantity=2.0)

        plm_df = pd.DataFrame([{"item_id": "302FP02010", "quantity": 2.0}])
        r3_df = pd.DataFrame([{"part_code": "302FP02010", "quantity": 2.0}])
        view.set_reference_data(plm_df, r3_df)

        result = view.run_preliminary_self_check()
        assert result["status"] == "OK"
        assert result["ok_count"] == 1
        assert result["ng_count"] == 0

    def test_f23_label_7980_7990_management(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """Test 4: Validate Label 7980/7990 UI fields."""
        view = MemberWorkspaceView(base_dir=tmp_path)
        assert view.label_spec_combo.count() >= 2
        view.label_code_edit.setText("7980-VN-001")
        assert view.label_code_edit.text() == "7980-VN-001"

    def test_f23_export_member_workbook(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        """Test 5: Export member data to structured records."""
        view = MemberWorkspaceView(base_dir=tmp_path)
        view.clear_cttt_table()
        view.add_cttt_row(page="01", part_code="302FP02010", quantity=1.0)

        data = view.get_cttt_table_data()
        assert len(data) == 1
        assert data[0]["MÃ LINH KIỆN"] == "302FP02010"
        assert data[0]["SỐ LƯỢNG"] == 1.0
