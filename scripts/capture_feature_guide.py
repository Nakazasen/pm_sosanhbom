"""Feature capture driver: drives the real SSBOM dialogs / BOM-update flow for machine 110C3M3AK0.

Screenshots -> docs/images_feat/f*.png ; log -> sample_data/run_feat/feature_log.txt
Excel COM is ENABLED here (BOM update writes through a private Excel instance).

Run:  python scripts/capture_feature_guide.py [part ...]     parts: dialogs update theme
"""

from __future__ import annotations

import json
import os
import shutil
import sys
import time
import traceback
import warnings
import zipfile
from pathlib import Path

warnings.filterwarnings("ignore")
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
os.environ["QT_QPA_PLATFORM"] = "windows"
os.environ["SSBOM_DISABLE_EXCEL_COM"] = "0"

import openpyxl
from PyQt6.QtCore import QEventLoop, QTimer
from PyQt6.QtWidgets import QApplication, QCheckBox, QDialog, QFileDialog, QMessageBox, QTabWidget

SAMPLE = ROOT / "sample_data" / "C3MAKO - Copy"
RUN = ROOT / "sample_data" / "run_feat"
OUT = ROOT / "docs" / "images_feat"
MACHINE = "110C3M3AK0"
MODEL = "Virgo"
CFG = ROOT / "config" / "settings.json"
LOG: list[str] = []
LOG_ENTRIES: list[tuple[str, str, str]] = []
W, H = 1500, 900
NEXT: list[str] = []


def wait(ms: int) -> None:
    loop = QEventLoop()
    QTimer.singleShot(ms, loop.quit)
    loop.exec()


def shot(widget, name: str, w: int = W, h: int = H, resize: bool = True) -> None:
    if resize:
        widget.resize(w, h)
    widget.show()
    wait(500)
    for _ in range(4):
        QApplication.processEvents()
    OUT.mkdir(parents=True, exist_ok=True)
    pm = widget.grab()
    pm.save(str(OUT / name), "PNG")
    print(f"[shot] {name} {pm.width()}x{pm.height()}", flush=True)


def patch_dialogs() -> None:
    def rec(kind):
        def _f(parent, title="", text="", *a, **k):
            LOG.append(f"[{kind}] {title} | {str(text)}")
            LOG_ENTRIES.append((kind, str(title), str(text)))
            return QMessageBox.StandardButton.Yes if kind == "question" else QMessageBox.StandardButton.Ok
        return _f

    QMessageBox.information = staticmethod(rec("info"))
    QMessageBox.warning = staticmethod(rec("warn"))
    QMessageBox.critical = staticmethod(rec("crit"))
    QMessageBox.question = staticmethod(rec("question"))

    def mask_secrets(dlg) -> None:
        """Replace real account IDs / passwords shown in dialogs; screenshots go into a shared document."""
        from PyQt6.QtWidgets import QComboBox, QLineEdit

        repl = {"vn_pe02": "<tài khoản TC24>", "v130474": "<tài khoản SAP>"}
        for e in dlg.findChildren(QLineEdit):
            if e.echoMode() != QLineEdit.EchoMode.Normal:
                e.setText("********")
                continue
            for k, v in repl.items():
                if k in e.text():
                    e.setText(e.text().replace(k, v))
        for c in dlg.findChildren(QComboBox):
            for i in range(c.count()):
                t = c.itemText(i)
                for k, v in repl.items():
                    if k in t:
                        c.setItemText(i, t.replace(k, v))
            if c.isEditable():
                for k, v in repl.items():
                    if k in c.currentText():
                        c.setCurrentText(c.currentText().replace(k, v))

    globals()["mask_secrets"] = mask_secrets

    def fake_exec(self):
        base = NEXT.pop(0) if NEXT else f"f_unnamed_{self.__class__.__name__}"
        print("[dialog]", self.__class__.__name__, "->", base, flush=True)
        try:
            mask_secrets(self)
            self.resize(max(self.width(), 1100), 1000)
            shot(self, f"{base}.png", resize=False)
            tabs = [t for t in self.findChildren(QTabWidget)]
            for ti, tw in enumerate(tabs):
                if tw.count() > 1:
                    cur = tw.currentIndex()
                    for i in range(tw.count()):
                        tw.setCurrentIndex(i)
                        shot(self, f"{base}_tab{ti}_{i}.png", resize=False)
                    tw.setCurrentIndex(cur)
        except Exception:
            traceback.print_exc()
        self.close()
        return 0

    QDialog.exec = fake_exec


def build_state(win):
    """Same project setup as the e2e run (real widgets), in a separate work dir."""
    leader = win.leader_view
    s1 = leader.step1_widget
    if RUN.exists():
        shutil.rmtree(RUN)
    (RUN).mkdir(parents=True)
    shutil.copy2(ROOT / "formnguoidung.xlsm", RUN / "formnguoidung.xlsm")
    s1.edit_base_dir.setText(str(RUN))
    s1.state.base_dir = RUN
    s1.model_combo.setCurrentText(MODEL)
    s1.stage_combo.setCurrentIndex(0)
    s1.machine_table.setRowCount(0)
    s1._add_machine_row(code=MACHINE, model_name=MODEL)
    s1._reindex_machine_stt()
    s1._set_all_staff_checked(False)
    wanted = {"Duong_mecha1.2", "Ly_mecha1.2", "Nhung_mecha1.2", "Hung_mecha1.1", "VietE_mecha1.2",
              "Tu_mecha1.2", "Thai_mecha1.2", "Chien_mecha1.2", "Thuong_mecha1.2"}
    for r, a in enumerate(s1.state.staff_roster):
        if a.engineer_name in wanted:
            s1.staff_table.cellWidget(r, 0).findChild(QCheckBox).setChecked(True)
    s1.chk_only_assigned_model.setChecked(True)
    s1._filter_staff_table()
    s1.sync_state_from_ui()
    s1.execute_create_folders_and_packages()
    leader.refresh_kpi_cards()
    mdir = RUN / MODEL / MACHINE
    # real sample files: PLM export, R3 export (from BOM R3 sheet), finished BOM with macros
    shutil.copy2(SAMPLE / f"PLM_{MACHINE}.xlsx", mdir / f"PLM_{MACHINE}.xlsx")
    shutil.copy2(SAMPLE / f"BOM_{MACHINE}.xlsm", mdir / f"BOM_{MACHINE}.xlsm")
    src_wb = openpyxl.load_workbook(SAMPLE / f"BOM_{MACHINE}.xlsm", read_only=True, data_only=True)
    out = openpyxl.Workbook()
    ows = out.active
    ows.title = "R3"
    for row in src_wb["R3"].iter_rows(values_only=True):
        ows.append(list(row))
    src_wb.close()
    out.save(mdir / f"R3_{MACHINE}.xlsx")
    s1.sync_state_from_ui()
    for m in s1.state.machines:
        if m.machine_code == MACHINE:
            m.plm_file = mdir / f"PLM_{MACHINE}.xlsx"
            m.r3_file = mdir / f"R3_{MACHINE}.xlsx"
    leader.step2_widget.refresh_sourcing_table()
    leader.state.last_bom_file = mdir / f"BOM_{MACHINE}.xlsm"
    return leader, mdir


def bom_report(path: Path) -> str:
    lines = [f"{path.name}: {path.stat().st_size:,} bytes"]
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        lines.append("vbaProject.bin: " + str("xl/vbaProject.bin" in names))
        lines.append("pivotTables: " + str(sum(1 for n in names if n.startswith("xl/pivotTables/pivotTable"))))
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    for ws in wb.worksheets:
        lines.append(f"  sheet {ws.title}: rows={ws.max_row}")
    wb.close()
    return "\n".join(lines)


def part_dialogs(win, leader, mdir):
    s1, s2, s3, s4 = leader.step1_widget, leader.step2_widget, leader.step3_widget, leader.step4_widget
    jobs = [
        ("f01_dlg_plm_download", lambda: leader._open_plm_download_dialog()),
        ("f02_dlg_filter_config", lambda: s2.open_bom_filter_dialog()),
        ("f03_dlg_visual_builder", lambda: s2.open_visual_bom_builder_dialog()),
        ("f04_dlg_standalone_filter", lambda: s2.open_standalone_bom_filter_dialog()),
        ("f05_dlg_bom_update", lambda: win.open_bom_update_dialog(initial_dir=mdir)),
        ("f06_dlg_settings", lambda: win.open_settings_dialog()),
        ("f07_dlg_roster", lambda: s1._open_manage_roster_dialog()),
        ("f08_dlg_model_manager", lambda: s1._open_model_machine_manager_dialog()),
        ("f09_dlg_pcd_scan", lambda: s1._open_pcd_scan_dialog()),
    ]
    for name, fn in jobs:
        NEXT.append(name)
        t = time.time()
        try:
            fn()
        except Exception:
            traceback.print_exc()
        print(f"   {name} done in {time.time() - t:.1f}s", flush=True)
        NEXT.clear()


def render_msg(title: str, text: str, name: str) -> None:
    box = QMessageBox(QMessageBox.Icon.Information, title, text)
    box.setMinimumWidth(560)
    shot(box, name, resize=False)
    box.close()


def part_update(win, leader, mdir):
    s4 = leader.step4_widget
    bom = mdir / f"BOM_{MACHINE}.xlsm"
    Path(RUN / "bom_before.txt").write_text(bom_report(bom), encoding="utf-8")

    # --- dialog "Cap nhat BOM moi (PLM & R3)" run for real
    from src.gui.bom_update_dialog import BOMUpdateDialog

    dlg = BOMUpdateDialog(parent=win, initial_dir=mdir)
    dlg.resize(980, 820)
    shot(dlg, "f10_update_dialog_scanned.png", resize=False)
    dlg._start_update(True, True)
    t0 = time.time()
    last = 0
    while dlg._worker is not None and dlg._worker.isRunning() and time.time() - t0 < 900:
        wait(500)
        if time.time() - last > 20:
            print(f"   updating... {int(time.time() - t0)}s", flush=True)
            last = time.time()
    wait(800)
    print(f"[update] finished in {time.time() - t0:.1f}s", flush=True)
    sb = dlg.txt_log.verticalScrollBar()
    sb.setValue(sb.maximum())
    dlg.resize(980, 900)
    shot(dlg, "f11_update_dialog_done.png", resize=False)
    Path(RUN / "update_log.txt").write_text(dlg.txt_log.toPlainText(), encoding="utf-8")
    dlg.close()
    Path(RUN / "bom_after_update.txt").write_text(bom_report(bom), encoding="utf-8")
    for kind, title, text in [e for e in LOG_ENTRIES if "Hoàn Tất" in e[1] or "Hoan" in e[1]][-1:]:
        render_msg(title, text, "f12_msg_update_done.png")

    # --- Step 4 discrepancy filter on the same BOM
    leader.switch_to_step(3)
    s4.combo_active_machine.setCurrentText(MACHINE)
    n_before = len(LOG_ENTRIES)
    s4._on_filter_discrepancies_clicked()
    for kind, title, text in LOG_ENTRIES[n_before:]:
        render_msg(title, text, "f13_msg_filter_done.png")
    Path(RUN / "bom_after_filter.txt").write_text(bom_report(bom), encoding="utf-8")


def part_filterrun(win, leader, mdir):
    """Run the standalone BOM filter for real on a copy of the PLM file."""
    from src.gui.standalone_bom_filter_dialog import StandaloneBOMFilterDialog

    demo = RUN / "filter_demo"
    demo.mkdir(parents=True, exist_ok=True)
    f = demo / f"PLM_{MACHINE}.xlsx"
    shutil.copy2(SAMPLE / f"PLM_{MACHINE}.xlsx", f)
    wb = openpyxl.load_workbook(f, read_only=True)
    rows_before = wb.worksheets[0].max_row
    wb.close()
    d = StandaloneBOMFilterDialog(initial_model=MODEL, initial_path=f, parent=win)
    d.resize(1100, 760)
    shot(d, "f14_filter_standalone_before.png", resize=False)
    d.btn_start.click()
    t0 = time.time()
    while d.worker is not None and d.worker.isRunning() and time.time() - t0 < 300:
        wait(400)
    wait(800)
    shot(d, "f15_filter_standalone_after.png", resize=False)
    wb = openpyxl.load_workbook(f, read_only=True)
    rows_after = wb.worksheets[0].max_row
    wb.close()
    (RUN / "filter_demo_result.txt").write_text(
        f"rows before={rows_before} after={rows_after}\n" + "\n".join(p.name for p in sorted(demo.iterdir())),
        encoding="utf-8")
    print(f"[filter] rows {rows_before} -> {rows_after}; files: {[p.name for p in demo.iterdir()]}", flush=True)
    d.close()


def part_extras(win, leader, mdir):
    s1, s2, s4 = leader.step1_widget, leader.step2_widget, leader.step4_widget
    from src.gui.plm_download_dialog import BOMScanDialog
    from src.gui.update_dialog import UpdateDialog
    from src.services.app_updates import AppUpdateManager
    from src.services.update_delivery import UpdateCandidate, UpdateDeliveryService

    d = BOMScanDialog(parent=win, dest_dir=mdir, project_parts=[MACHINE])
    mask_secrets(d)
    d.resize(1100, 800)
    shot(d, "f30_dlg_bom_scan.png", resize=False)
    d.close()

    for name, fn in [
        ("f31_dlg_mail_manager", lambda: s4._on_send_management_review_email()),
        ("f32_dlg_mail_assign", lambda: s1._on_send_assignment_email()),
    ]:
        NEXT.append(name)
        try:
            fn()
        except Exception:
            traceback.print_exc()
        NEXT.clear()

    cand = UpdateCandidate(
        version="1.0.7",
        package_name="SSBOM_Manager-1.0.7.mpupdate",
        source_location="LAN: release_update\\SSBOM_Manager-1.0.7.mpupdate",
        package_sha256="ee6c0582e1ed74929eda2b4b8224bc4d0a707475f3f16ec918f5b83548a7af72",
        package_size=134648009,
        release_notes=json.loads((ROOT / "release.json").read_text(encoding="utf-8"))["notes"],
    )
    delivery = UpdateDeliveryService(base_dir=ROOT, current_version="1.0.6")
    mgr = AppUpdateManager(install_root=ROOT)
    u = UpdateDialog(parent=win, candidate=cand, current_version="1.0.6", delivery_service=delivery, update_manager=mgr)
    u.resize(720, 620)
    shot(u, "f33_dlg_software_update.png", resize=False)
    u.close()


def part_theme(win, leader, mdir):
    from src.gui.styles import get_theme_manager

    tm = get_theme_manager()
    leader.switch_to_step(0)
    tm.set_theme("dark", apply_stylesheet=True)
    wait(600)
    shot(win, "f20_theme_dark_step1.png")
    leader.switch_to_step(3)
    shot(win, "f21_theme_dark_step4.png")
    tm.set_theme("light", apply_stylesheet=True)
    wait(600)


def main() -> None:
    parts = sys.argv[1:] or ["dialogs", "update", "extras", "theme"]
    cfg_backup = CFG.read_bytes() if CFG.exists() else None
    app = QApplication.instance() or QApplication(sys.argv)
    try:
        from src.gui.app import SSBOMMainWindow
        from src.gui.styles import get_theme_manager

        get_theme_manager().set_theme("light", apply_stylesheet=True)
        patch_dialogs()
        win = SSBOMMainWindow()
        win.resize(W, H)
        win.show()
        app.processEvents()
        if hasattr(win, "tabs"):
            win.tabs.setCurrentIndex(0)
        leader, mdir = build_state(win)
        for p in parts:
            print(f"=== part {p}", flush=True)
            try:
                {"dialogs": part_dialogs, "update": part_update, "filterrun": part_filterrun, "extras": part_extras, "theme": part_theme}[p](win, leader, mdir)
            except Exception:
                traceback.print_exc()
        (RUN / "feature_log.txt").write_text("\n".join(LOG), encoding="utf-8")
        print("log entries:", len(LOG))
    finally:
        if cfg_backup is not None:
            CFG.write_bytes(cfg_backup)


if __name__ == "__main__":
    main()
    os._exit(0)
