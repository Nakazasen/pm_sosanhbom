"""Automated Screenshot Capture Engine for Kyocera SSBOM User Guide.

Instantiates each core GUI window and dialog with authentic sample data for
Model C3MAKO / Virgo (Machine Code 110C3M3AK0), applies appropriate styling,
renders the widgets, and saves high-resolution PNG screenshots into docs/images/.
"""

from __future__ import annotations

import datetime
import os
import sys
from pathlib import Path

# Ensure project root is in sys.path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

# Ensure Windows Qt platform
os.environ["QT_QPA_PLATFORM"] = "windows"
# Disable real Excel calls during testing/capturing
os.environ["SSBOM_DISABLE_EXCEL_COM"] = "1"

from PyQt6.QtCore import QDate, QSize, Qt
from PyQt6.QtGui import QColor, QFont
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDateEdit,
    QLineEdit,
    QTableWidgetItem,
)

from src.gui.app import SSBOMMainWindow
from src.gui.bom_filter_dialog import BOMFilterConfigDialog
from src.gui.bom_update_dialog import BOMUpdateDialog
from src.gui.bom_visual_builder_dialog import BOMVisualRuleBuilderDialog
from src.gui.member_roster_dialog import MemberRosterDialog
from src.gui.pcd_plan_dialog import PCDPlanScanDialog
from src.gui.plm_download_dialog import PLMDownloadDialog
from src.gui.settings_dialog import SettingsDialog
from src.gui.standalone_bom_filter_dialog import StandaloneBOMFilterDialog
from src.gui.styles import get_theme_manager
from src.gui.update_dialog import UpdateDialog
from src.services.app_updates import AppUpdateManager
from src.services.update_delivery import UpdateCandidate, UpdateDeliveryService


def save_widget_screenshot(widget, output_path: Path, width: int = 1200, height: int = 800) -> None:
    """Resize widget, force layout update, render to QPixmap, and save as PNG."""
    widget.resize(width, height)
    widget.show()
    QApplication.processEvents()
    QApplication.processEvents()
    pixmap = widget.grab()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    pixmap.save(str(output_path), "PNG")
    print(f"[OK] Saved screenshot: {output_path.name} ({pixmap.width()}x{pixmap.height()})")
    widget.close()


def setup_sample_leader_view(leader, sample_dir: Path) -> None:
    """Populate LeaderWorkspaceView with realistic data for model 110C3M3AK0 (C3MAKO / Virgo)."""
    try:
        # Step 1: Model & Machine Setup
        if hasattr(leader, "model_combo"):
            leader.model_combo.setCurrentText("Virgo")
        if hasattr(leader, "stage_combo"):
            leader.stage_combo.setCurrentIndex(0)  # MP (ma1 - tiền tố 110)
        if hasattr(leader, "edit_base_dir"):
            leader.edit_base_dir.setText(str(sample_dir))

        if hasattr(leader, "machine_table"):
            leader.machine_table.setRowCount(0)
            leader._add_machine_row(code="110C3M3AK0", note="Mã máy C3MAKO xuất EU/US (Chính thức)", model_name="Virgo")
            leader._add_machine_row(code="110C3M3NL0", note="Mã máy C3MAKO xuất Nhật/Châu Á", model_name="Virgo")

        if hasattr(leader, "staff_table"):
            # Set up sample assignments
            sample_assignments = [
                ("Son_mecha1", "Cơ 1", "FUSER"),
                ("Loc_mecha2", "Cơ 2", "PAPER FEED"),
                ("Vinh_mecha3", "Cơ 3", "FRAME"),
                ("Viet_mecha3", "Cơ 3", "DRUM"),
                ("Duy_mecha1", "Cơ 1", "LASER"),
            ]
            for r, (eng, dept, subunit) in enumerate(sample_assignments):
                if r < leader.staff_table.rowCount():
                    chk_w = leader.staff_table.cellWidget(r, 0)
                    if chk_w:
                        chk = chk_w.findChild(QCheckBox)
                        if chk:
                            chk.setChecked(True)
                    combo_m = leader.staff_table.cellWidget(r, 3)
                    if isinstance(combo_m, QComboBox):
                        combo_m.setCurrentText("110C3M3AK0")
                    combo_s = leader.staff_table.cellWidget(r, 4)
                    if isinstance(combo_s, QComboBox):
                        combo_s.setCurrentText(subunit)

        # Step 2: Sourcing
        step2 = getattr(leader, "step2_widget", None)
        if step2 and hasattr(step2, "sourcing_table"):
            step2.sourcing_table.setRowCount(0)
            step2.sourcing_table.insertRow(0)
            step2.sourcing_table.setItem(0, 0, QTableWidgetItem("1"))
            step2.sourcing_table.setItem(0, 1, QTableWidgetItem("110C3M3AK0"))
            step2.sourcing_table.setItem(0, 2, QTableWidgetItem(datetime.date.today().strftime("%Y/%m/%d")))

            item_plm = QTableWidgetItem("PLM_110C3M3AK0.xlsx (Đã sẵn sàng)")
            item_plm.setBackground(QColor("#D1E7DD"))
            item_plm.setForeground(QColor("#0F5132"))
            step2.sourcing_table.setItem(0, 3, item_plm)

            item_r3 = QTableWidgetItem("R3_110C3M3AK0.xlsx (CS12 Plant 2200)")
            item_r3.setBackground(QColor("#D1E7DD"))
            item_r3.setForeground(QColor("#0F5132"))
            step2.sourcing_table.setItem(0, 4, item_r3)

            item_flt = QTableWidgetItem("✓ Đã lọc xong (Lv1..6)")
            item_flt.setBackground(QColor("#D1E7DD"))
            item_flt.setForeground(QColor("#0F5132"))
            item_flt.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            step2.sourcing_table.setItem(0, 5, item_flt)

        # Step 3: Tracking & Consolidation
        step3 = getattr(leader, "step3_widget", None)
        if step3 and hasattr(step3, "status_table"):
            step3.status_table.setRowCount(0)
            sample_subs = [
                ("110C3M3AK0", "FUSER", "Son_mecha1", 42, "✓ OK"),
                ("110C3M3AK0", "PAPER FEED", "Loc_mecha2", 68, "✓ OK"),
                ("110C3M3AK0", "FRAME", "Vinh_mecha3", 115, "✓ OK"),
                ("110C3M3AK0", "DRUM", "Viet_mecha3", 34, "✓ OK"),
                ("110C3M3AK0", "LASER", "Duy_mecha1", 28, "✓ OK"),
            ]
            for r, (mach, unit, eng, items, stt) in enumerate(sample_subs):
                step3.status_table.insertRow(r)
                step3.status_table.setItem(r, 0, QTableWidgetItem(str(r + 1)))
                step3.status_table.setItem(r, 1, QTableWidgetItem(mach))
                step3.status_table.setItem(r, 2, QTableWidgetItem(unit))
                step3.status_table.setItem(r, 3, QTableWidgetItem(eng))
                it_stt = QTableWidgetItem(stt)
                it_stt.setBackground(QColor("#D1E7DD"))
                it_stt.setForeground(QColor("#0F5132"))
                it_stt.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                step3.status_table.setItem(r, 4, it_stt)
                step3.status_table.setItem(r, 5, QTableWidgetItem(str(items)))
                step3.status_table.setItem(r, 6, QTableWidgetItem(datetime.datetime.now().strftime("%d/%m/%Y %H:%M")))

        # Step 4: Master BOM
        step4 = getattr(leader, "step4_widget", None)
        if step4:
            if hasattr(step4, "combo_jig_model"):
                step4.combo_jig_model.setCurrentText("Virgo")
            if hasattr(step4, "chk_4m"):
                step4.chk_4m.setChecked(True)
            if hasattr(step4, "edit_ktsx_name"):
                step4.edit_ktsx_name.setText("Bùi Đức Vinh (PE Dept)")
            if hasattr(step4, "lbl_bom_status"):
                step4.lbl_bom_status.setText("Đã tạo thành công: BOM_110C3M3AK0.xlsm")
    except Exception as ex:
        print(f"[WARN] Error setting up leader view: {ex}")


def setup_sample_member_view(member, sample_dir: Path) -> None:
    """Populate MemberWorkspaceView with realistic data for model 110C3M3AK0 (FUSER unit)."""
    try:
        if hasattr(member, "combo_machine"):
            member.combo_machine.setCurrentText("110C3M3AK0")
        if hasattr(member, "combo_sub_unit"):
            member.combo_sub_unit.setCurrentText("FUSER")
        if hasattr(member, "combo_engineer"):
            member.combo_engineer.setCurrentText("Vinh_mecha3")

        # Populate sample CTTT parts into member table
        table = getattr(member, "table_cttt", None)
        if table:
            sample_parts = [
                ("1", "302FP02010", "PULLEY EXIT", "1", "PC", "EXIT UNIT", "Linh kiện lắp ráp chính"),
                ("2", "302M294010", "ROLLER HEAT", "1", "PC", "FUSER HEATER", "Đã đối chiếu khớp R3"),
                ("3", "302V394010", "THERMISTOR ASSY", "2", "PC", "FUSER SENSOR", "Bản vẽ Rev B"),
                ("4", "302ND94010", "GEAR Z32R", "1", "PC", "FUSER DRIVE", "Khớp cấu trúc PLM"),
                ("5", "302H794010", "BUSHING BEARING", "2", "PC", "FUSER FRAME", "Linh kiện chuẩn"),
            ]
            table.setRowCount(0)
            for r, row_data in enumerate(sample_parts):
                table.insertRow(r)
                for c, val in enumerate(row_data):
                    table.setItem(r, c, QTableWidgetItem(val))

        # Show status Q2 = OK
        if hasattr(member, "lbl_submission_status"):
            member.lbl_submission_status.setText("Trạng thái hiện tại: Đã nộp (CTTT!Q2 = OK) - Sẵn sàng tổng hợp")
            member.lbl_submission_status.setStyleSheet("color: #0F5132; font-weight: bold; background: #D1E7DD; padding: 4px;")
    except Exception as ex:
        print(f"[WARN] Error setting up member view: {ex}")


def main() -> None:
    app = QApplication.instance() or QApplication(sys.argv)
    theme_mgr = get_theme_manager()
    theme_mgr.set_theme("light", apply_stylesheet=True)

    img_dir = Path("docs/images")
    img_dir.mkdir(parents=True, exist_ok=True)
    sample_dir = Path("d:/Sandbox/pm_sosanhbom/sample_data/C3MAKO - Copy").resolve()

    # 1. Main Window - Step 1: Lập Dự Án (Model 110C3M3AK0 / C3MAKO)
    main_win = SSBOMMainWindow()
    main_win.resize(1300, 850)

    # Disable background worker to avoid network hanging
    if hasattr(main_win, "_update_check_worker") and main_win._update_check_worker:
        try:
            main_win._update_check_worker.quit()
        except Exception:
            pass

    leader = getattr(main_win, "leader_view", None)
    if leader:
        setup_sample_leader_view(leader, sample_dir)

    member = getattr(main_win, "member_view", None)
    if member:
        setup_sample_member_view(member, sample_dir)

    # Ensure Tab Leader is active
    if hasattr(main_win, "tabs"):
        main_win.tabs.setCurrentIndex(0)

    if leader:
        # Step 1
        if hasattr(leader, "wizard_stack"):
            leader.wizard_stack.setCurrentIndex(0)
        save_widget_screenshot(main_win, img_dir / "01_main_window_overview.png", 1300, 850)

        # Step 2
        main_win.show()
        if hasattr(leader, "wizard_stack"):
            leader.wizard_stack.setCurrentIndex(1)
        save_widget_screenshot(main_win, img_dir / "02_step2_data_sourcing.png", 1300, 850)

        # Step 3
        main_win.show()
        if hasattr(leader, "wizard_stack"):
            leader.wizard_stack.setCurrentIndex(2)
        save_widget_screenshot(main_win, img_dir / "03_step3_tracking_consolidation.png", 1300, 850)

        # Step 4
        main_win.show()
        if hasattr(leader, "wizard_stack"):
            leader.wizard_stack.setCurrentIndex(3)
        save_widget_screenshot(main_win, img_dir / "04_step4_master_bom_reporting.png", 1300, 850)

    # 5. Member Workspace Tab
    main_win.show()
    if hasattr(main_win, "tabs"):
        main_win.tabs.setCurrentIndex(1)
    save_widget_screenshot(main_win, img_dir / "05_member_workspace.png", 1300, 850)
    main_win.close()

    # 6. Dialog BOM Filter
    dlg_filter = BOMFilterConfigDialog(initial_model="Virgo", parent=None)
    save_widget_screenshot(dlg_filter, img_dir / "06_dialog_bom_filter.png", 1000, 680)

    # 7. Dialog PLM / SAP Download
    dlg_plm = PLMDownloadDialog(parent=None, initial_model="Virgo", project_parts=["110C3M3AK0"])
    save_widget_screenshot(dlg_plm, img_dir / "07_dialog_plm_download.png", 1050, 720)

    # 8. Dialog Settings
    dlg_settings = SettingsDialog(parent=None)
    save_widget_screenshot(dlg_settings, img_dir / "08_dialog_settings.png", 850, 620)

    # 9. Dialog Update (Theme-aware v1.0.6) - Correct UpdateCandidate arguments!
    cand = UpdateCandidate(
        version="1.0.6",
        package_name="SSBOM_Manager-1.0.6.mpupdate",
        source_location="file://///fstvn01/Data/.../SSBOM_Manager-1.0.6.mpupdate",
        package_sha256="670150211fd6dccc2957f42f5a6badc52d3b0113380e351bcff3f29bfcb85117",
        package_size=134643082,
        release_notes=(
            "SSBOM v1.0.6 Release Notes:\n\n"
            "• Sửa triệt để lỗi file BOM .xlsm không mở được sau khi Cập nhật BOM mới (PLM & R3)\n"
            "• Bảo toàn 100% VBA macro, 3 nút bấm tương tác, 9 shape vẽ và Pivot Table W1\n"
            "• Tự động gỡ và thiết lập lại AutoFilter khi ghi dữ liệu, chống lỗi lệch dòng\n"
            "• Dark Theme: Giao diện nền tối hiển thị sắc nét, tương phản đạt chuẩn WCAG AA\n"
            "• Kế thừa giải trình tự động: Cột O (Giải thích), Cột P (Phụ trách), Cột Q (Quản lý check)"
        ),
    )
    delivery = UpdateDeliveryService(base_dir=Path("."), current_version="1.0.5")
    mgr = AppUpdateManager(install_root=Path("."))
    dlg_upd = UpdateDialog(parent=None, candidate=cand, current_version="1.0.5", delivery_service=delivery, update_manager=mgr)
    save_widget_screenshot(dlg_upd, img_dir / "09_dialog_update.png", 680, 520)

    # 10. Dialog BOM Update (PLM & R3) - Core v1.0.6 Feature with Real Sample Data
    bom_file = sample_dir / "BOM_110C3M3AK0.xlsm"
    dlg_bom_upd = BOMUpdateDialog(initial_dir=sample_dir, initial_bom_file=bom_file if bom_file.exists() else None, parent=None)
    save_widget_screenshot(dlg_bom_upd, img_dir / "10_dialog_bom_update.png", 980, 660)

    # 11. Dialog Standalone BOM Filter
    plm_file = sample_dir / "PLM_110C3M3AK0.xlsx"
    dlg_std_filter = StandaloneBOMFilterDialog(
        initial_model="Virgo",
        initial_path=plm_file if plm_file.exists() else sample_dir,
        parent=None,
    )
    save_widget_screenshot(dlg_std_filter, img_dir / "11_dialog_standalone_filter.png", 980, 640)

    # 12. Dialog PCD Plan Scan SharePoint
    dlg_pcd = PCDPlanScanDialog(parent=None)
    save_widget_screenshot(dlg_pcd, img_dir / "12_dialog_pcd_plan.png", 950, 650)

    # 13. Dialog Member Roster / Leader Management
    dlg_roster = MemberRosterDialog(parent=None)
    save_widget_screenshot(dlg_roster, img_dir / "13_dialog_member_roster.png", 950, 650)

    # 14. Dialog BOM Visual Rule Builder
    dlg_visual = BOMVisualRuleBuilderDialog(
        initial_model="Virgo",
        initial_file=plm_file if plm_file.exists() else None,
        parent=None,
    )
    save_widget_screenshot(dlg_visual, img_dir / "14_dialog_bom_visual_builder.png", 1000, 680)

    # 15. Dark Mode Overview (Main Window in Dark Theme)
    theme_mgr.set_theme("dark", apply_stylesheet=True)
    main_win_dark = SSBOMMainWindow()
    if hasattr(main_win_dark, "_update_check_worker") and main_win_dark._update_check_worker:
        try:
            main_win_dark._update_check_worker.quit()
        except Exception:
            pass
    leader_dark = getattr(main_win_dark, "leader_view", None)
    if leader_dark:
        setup_sample_leader_view(leader_dark, sample_dir)
        if hasattr(leader_dark, "wizard_stack"):
            leader_dark.wizard_stack.setCurrentIndex(0)
    save_widget_screenshot(main_win_dark, img_dir / "15_dark_mode_overview.png", 1300, 850)
    main_win_dark.close()

    # Reset theme back to light
    theme_mgr.set_theme("light", apply_stylesheet=True)

    print("\n[SUCCESS] All 15 authentic screenshots captured cleanly into docs/images/!")
    sys.exit(0)


if __name__ == "__main__":
    main()
