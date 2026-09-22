"""Unified Automated BOM Download Dialog for Siemens Teamcenter TC24 & SAP R3.

Designed specifically for non-tech manufacturing engineers and team leads:
- Simple 3-step workflow (Paste BOM list -> Click download).
- Clear separation: Effective date applies ONLY to SAP R3; TC24 PLM always pulls latest.
- Optional intuitive per-BOM date configuration table (only displayed when requested).
"""

from __future__ import annotations

import logging
import re
import time
import zipfile
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Dict, List, Optional

from src.services.machine_dict_service import MachineDictService

from PyQt6.QtCore import QDate, QObject, QThread, Qt, pyqtSignal, pyqtSlot
from PyQt6.QtGui import QFont, QGuiApplication, QTextCursor
from PyQt6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDateEdit,
    QDialog,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from src.automation.tc2412.standardizer import (
    standardize_plm_file,
)
from src.gui.styles import get_theme_manager

logger = logging.getLogger(__name__)

SO_SANH_BASE_DIR = Path(
    r"\\fstvn01\Data\10_Production Engineering Department(製造技術部)\02.製造技術課\PE Dept\4A. QUAN LY BOM-TDTK-BOM管理-設計変更\SO SANH PLM-CTTT-R3"
)
VIRGO_DIR = SO_SANH_BASE_DIR / "Virgo2"

MODE_BOTH = "BOTH"
MODE_PLM = "PLM"
MODE_SAP = "SAP"


def try_parse_date(date_str: str) -> Optional[date]:
    """Attempt to parse date string across common factory spreadsheet formats."""
    clean = date_str.strip().strip(",;|\t/.-")
    formats = [
        "%Y/%m/%d",
        "%Y-%m-%d",
        "%d/%m/%Y",
        "%d.%m.%Y",
        "%d-%m-%Y",
        "%Y%m%d",
        "%Y.%m.%d",
    ]
    for fmt in formats:
        try:
            return datetime.strptime(clean, fmt).date()
        except ValueError:
            continue
    return None


def convert_to_pure_xlsx(src_file: Path, dest_xlsx: Path) -> bool:
    """Lossless conversion from TC2412 export to pure .xlsx, preserving 100% of formatting."""
    try:
        with zipfile.ZipFile(src_file, "r") as zin:
            with zipfile.ZipFile(dest_xlsx, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename.endswith("vbaProject.bin"):
                        continue
                    data = zin.read(item.filename)
                    if item.filename == "[Content_Types].xml":
                        data = data.replace(
                            b"application/vnd.ms-excel.sheet.macroEnabled.main+xml",
                            b"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                        )
                    zout.writestr(item, data)
        return True
    except Exception as exc:
        logger.warning("Error converting to pure xlsx: %s", exc)
        return False


def is_macro_file(file_path: Path) -> bool:
    """Check if an Excel file contains macro parts (.xlsm, .xltm or internal vbaProject/macroEnabled)."""
    if not file_path.exists():
        return False
    if file_path.suffix.lower() in (".xlsm", ".xltm") or "macro" in file_path.name.lower():
        return True
    try:
        with zipfile.ZipFile(file_path, "r") as z:
            names = set(z.namelist())
            if any("vbaProject" in n for n in names):
                return True
            if "[Content_Types].xml" in names:
                ct = z.read("[Content_Types].xml").decode("utf-8", errors="ignore")
                if "macroEnabled" in ct:
                    return True
    except Exception:
        pass
    return False


def ensure_ktct_trong_arrangement(driver, panel, wait) -> bool:
    """Ensure that the 'KTCT_Trong' column arrangement is selected in the Export panel."""
    try:
        from selenium.webdriver.common.by import By
        panel_text = getattr(panel, "text", "")
        if "KTCT_Trong" in panel_text:
            return True

        arr_btns = panel.find_elements(By.CSS_SELECTOR, "button[command-id='Arm0ArrangeViewConfigs']")
        if arr_btns:
            arr_btns[0].click()
            time.sleep(1.5)

            menu_items = driver.find_elements(
                By.CSS_SELECTOR,
                "div.aw-popup div.aw-widgets-cellListItem, div.aw-popup li, div.sw-popup li, div.sw-popup div",
            )
            for m in menu_items:
                if (m.text or "").strip() == "KTCT_Trong":
                    m.click()
                    time.sleep(2)
                    return True
        return False
    except Exception as exc:
        logger.warning("Failed selecting KTCT_Trong arrangement: %s", exc)
        return False


@dataclass
class BOMDownloadItem:
    """Represents a single BOM part number with optional custom valid date for SAP R3."""
    part_number: str
    custom_date: Optional[date] = None

    @property
    def has_custom_date(self) -> bool:
        return self.custom_date is not None


def parse_bom_items(raw_text: str) -> List[BOMDownloadItem]:
    """Parse raw text into structured BOMDownloadItems.
    
    Extracts part numbers and optional dates if user pastes multiple columns from Excel.
    """
    if not raw_text:
        return []

    items: List[BOMDownloadItem] = []
    seen = set()

    for line in raw_text.splitlines():
        line_clean = line.strip()
        if not line_clean:
            continue

        tokens = [t.strip() for t in line_clean.replace("\t", " ").replace(",", " ").replace(";", " ").replace("|", " ").split() if t.strip()]
        if not tokens:
            continue

        part = tokens[0].upper()
        if not part or part in seen:
            continue
        seen.add(part)

        custom_dt = None
        if len(tokens) > 1:
            custom_dt = try_parse_date(tokens[1])

        items.append(BOMDownloadItem(part_number=part, custom_date=custom_dt))

    return items


def parse_part_numbers(raw_text: str) -> List[str]:
    """Backwards compatibility helper returning list of part number strings."""
    return [item.part_number for item in parse_bom_items(raw_text)]


class UnifiedBOMDownloadWorker(QObject):
    """Executes multi-source BOM downloads (TC24 PLM and/or SAP R3 CS12) in background."""

    progress = pyqtSignal(int, str)       # percent (0-100), status message
    log_message = pyqtSignal(str)         # detailed log output
    finished = pyqtSignal(bool, str)      # success boolean, summary string

    def __init__(
        self,
        items: List[BOMDownloadItem],
        mode: str,
        tc_account: str,
        tc_auto_standardize: bool,
        sap_plant: str,
        sap_bom_usage: str,
        sap_alternative: str,
        sap_default_date: date,
        custom_dates_map: Dict[str, date],  # Part -> specific date if enabled
        dest_dir: Path,
        model_name: str = "Virgo",
        sap_auto_logout: bool = True,
    ) -> None:
        super().__init__()
        self.items = items
        self.mode = mode
        self.tc_account = tc_account
        self.tc_auto_standardize = tc_auto_standardize
        self.sap_plant = sap_plant
        self.sap_bom_usage = sap_bom_usage
        self.sap_alternative = sap_alternative
        self.sap_default_date = sap_default_date
        self.custom_dates_map = custom_dates_map
        self.dest_dir = dest_dir
        self.model_name = model_name
        self.sap_auto_logout = sap_auto_logout

    @pyqtSlot()
    def run(self) -> None:
        """Run download and standardization pipeline according to selected mode."""
        total_parts = len(self.items)
        if total_parts == 0:
            self.finished.emit(False, "Không có mã BOM nào để tải.")
            return

        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pass

        try:
            self._execute_run(total_parts)
        finally:
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def _execute_run(self, total_parts: int) -> None:
        self.dest_dir.mkdir(parents=True, exist_ok=True)
        plm_success = 0
        sap_success = 0
        error_messages: List[str] = []

        steps_per_part = 2 if self.mode == MODE_BOTH else 1
        total_steps = total_parts * steps_per_part
        current_step = 0

        self.progress.emit(2, f"Bắt đầu tải {total_parts} mã BOM (Chế độ: {self.mode})...")
        self.log_message.emit(f"=== BẮT ĐẦU TIẾN TRÌNH TẢI BOM ===")
        self.log_message.emit(f"[*] Chế độ thực hiện: {self.mode}")
        self.log_message.emit(f"[*] Số lượng mã BOM: {total_parts}")
        self.log_message.emit(f"[*] Thư mục lưu kết quả: {self.dest_dir}")

        # ---------------------------------------------------------------------
        # 1. Teamcenter TC24 PLM Download (No date parameter needed)
        # ---------------------------------------------------------------------
        if self.mode in (MODE_BOTH, MODE_PLM):
            self.log_message.emit(f"\n>> BƯỚC 1: TẢI TỪ SIEMENS TEAMCENTER (TC24) [Tài khoản: {self.tc_account}]")
            self.log_message.emit("   (Lưu ý: TC24 luôn tự động tải cây cấu trúc BOM mới nhất)")

            tc_client = None
            try:
                from selenium.webdriver.common.by import By
                from selenium.webdriver.support import expected_conditions as EC
                from selenium.webdriver.support.ui import WebDriverWait

                from src.automation.tc2412.client import TC2412AutomationClient
                from src.automation.tc2412.session import BrowserConfig, TC2412SessionManager
                from src.security.credentials import CredentialManager

                download_scratch = Path("scratch/downloads").resolve()
                download_scratch.mkdir(parents=True, exist_ok=True)

                config = BrowserConfig(
                    download_dir=download_scratch,
                    headless=True,
                )
                sm = TC2412SessionManager(config=config)
                tc_client = TC2412AutomationClient(session_manager=sm)

                tc_user = self.tc_account.strip() or "vn_pe02"
                tc_pass = tc_user
                try:
                    stored_creds = CredentialManager.get_credentials("PM_SOSANHBOM_TC2412")
                    if stored_creds and stored_creds.username == tc_user and stored_creds.password:
                        tc_pass = stored_creds.password
                except Exception:
                    pass

                self.log_message.emit(f"[*] [TC24] Đang đăng nhập tự động vào Teamcenter 2412 ({tc_user})...")
                try:
                    tc_client.login(username=tc_user, password=tc_pass)
                    self.log_message.emit("[+] [TC24] Đăng nhập thành công! Bắt đầu trích xuất BOM.")
                except Exception as login_err:
                    self.log_message.emit(f"[-] [TC24] Cảnh báo đăng nhập: {login_err}")

                driver = tc_client.driver
                wait = WebDriverWait(driver, 25)

                for idx, item in enumerate(self.items, 1):
                    part = item.part_number
                    current_step += 1
                    pct = int((current_step - 1) / total_steps * 95) + 2
                    self.progress.emit(pct, f"[{current_step}/{total_steps}] TC24: Đang tải mã {part}...")
                    self.log_message.emit(f"\n--- [TC24] Xử lý mã ({idx}/{total_parts}): {part} ---")

                    try:
                        model_slug = re.sub(r"[^\w]+", "_", (self.model_name or "raw").strip().lower()).strip("_")
                        bak_dir_name = f"backup_before_{model_slug}_filter"
                        bak_dir = self.dest_dir / bak_dir_name

                        raw_file = None
                        if bak_dir.exists() and (bak_dir / f"PLM_{part}.xlsx").exists():
                            raw_file = bak_dir / f"PLM_{part}.xlsx"
                        elif (self.dest_dir / "backup_before_virgo_filter" / f"PLM_{part}.xlsx").exists():
                            raw_file = self.dest_dir / "backup_before_virgo_filter" / f"PLM_{part}.xlsx"

                        target_dest = self.dest_dir / f"PLM_{part}.xlsx"

                        if raw_file and raw_file.exists():
                            self.log_message.emit(f"[+] [TC24] Tìm thấy bản gốc lưu trữ: {raw_file.name}")
                            orig_c, final_c = standardize_plm_file(
                                raw_file,
                                target_dest,
                                prune_electrical=self.tc_auto_standardize,
                            )
                            self.log_message.emit(
                                f"[+] [TC24] Chuẩn hóa 14 cột thành công: {orig_c:,} dòng raw -> {final_c:,} dòng chuẩn."
                            )
                            plm_success += 1
                        else:
                            self.log_message.emit(f"[*] [TC24] Tìm kiếm mã {part} trên hệ thống...")
                            for f in download_scratch.glob("*"):
                                try:
                                    f.unlink()
                                except Exception:
                                    pass

                            # 1. Search item
                            tc_client.search_item(part)
                            time.sleep(2.5)

                            # 2. Content tab
                            content_tab = wait.until(
                                EC.element_to_be_clickable(
                                    (By.XPATH, "//a[contains(@class, 'sw-tab-title') and normalize-space()='Content']")
                                )
                            )
                            content_tab.click()
                            time.sleep(3.5)

                            # 3. Select root row and Expand Below
                            root_cells = driver.find_elements(
                                By.CSS_SELECTOR, "div.aw-splm-tableRow div.aw-splm-tableCellText"
                            )
                            if root_cells:
                                root_cells[0].click()
                                time.sleep(1)

                            workarea_tb = driver.find_element(
                                By.CSS_SELECTOR, "div.aw-layout-workareaCommandbar, div.aw-commands-toolbar"
                            )
                            workarea_tb.find_element(By.CSS_SELECTOR, "button[command-id='Awb0Expand']").click()
                            time.sleep(1.5)

                            cmds = driver.find_elements(
                                By.CSS_SELECTOR, "div.aw-widgets-cellListItem, [command-id='Awb0ExpandBelow']"
                            )
                            exp_below = [
                                c for c in cmds if "expand below" in (c.text or "").lower() or c.get_attribute("command-id") == "Awb0ExpandBelow"
                            ]
                            if exp_below:
                                exp_below[0].click()
                                self.log_message.emit("   [*] Đang mở rộng toàn bộ cây BOM...")
                                time.sleep(10)

                            # 4. Select all rows
                            workarea_tb = driver.find_element(
                                By.CSS_SELECTOR, "div.aw-layout-workareaCommandbar, div.aw-commands-toolbar"
                            )
                            workarea_tb.find_element(By.CSS_SELECTOR, "button[command-id='Awp0SelectAll']").click()
                            time.sleep(2)

                            # 5. Open Export to Excel
                            workarea_tb.find_element(By.CSS_SELECTOR, "button[command-id='Arm0ExportImport']").click()
                            time.sleep(2)

                            popup = wait.until(
                                EC.presence_of_element_located(
                                    (By.CSS_SELECTOR, "div.aw-popup-commandListContainer, div.aw-popup, div.sw-popup")
                                )
                            )
                            for it in popup.find_elements(By.CSS_SELECTOR, ".aw-widgets-cellListItem, .aw-command, button, a, div"):
                                txt = (it.text or "").strip().lower()
                                if "export to excel" in txt and "import" not in txt:
                                    it.click()
                                    break
                            time.sleep(3)

                            # 6. Panel configuration
                            panel = wait.until(
                                EC.presence_of_element_located(
                                    (By.CSS_SELECTOR, "form.sw-command-panel, div.sw-right-dialog form")
                                )
                            )
                            if tc_user == "vn_pe02":
                                has_ktct = ensure_ktct_trong_arrangement(driver, panel, wait)
                                if has_ktct:
                                    self.log_message.emit("   [+] Đã áp dụng quy tắc xuất KTCT_Trong (luật của KTCT_Trọng).")

                            # Uncheck background for direct download
                            cbs = panel.find_elements(By.CSS_SELECTOR, "input[type='checkbox']")
                            for cb in cbs:
                                try:
                                    parent_lbl = cb.find_element(
                                        By.XPATH, "./ancestor::label | ./ancestor::div[contains(@class,'checkbox')]"
                                    )
                                    if "background" in parent_lbl.text.lower():
                                        if cb.is_selected():
                                            cb.click()
                                            time.sleep(0.5)
                                except Exception:
                                    pass

                            # Click Export button
                            export_btn = wait.until(
                                EC.element_to_be_clickable(
                                    (By.XPATH, "//button[normalize-space()='Export' or @button-id='exportBtn']")
                                )
                            )
                            driver.execute_script("arguments[0].click();", export_btn)
                            self.log_message.emit("   [*] Đã gửi yêu cầu Export, đang đợi máy chủ xuất file...")

                            # Wait for download to complete and stabilize
                            start_dl_wait = time.time()
                            downloaded_path = None
                            last_size = -1
                            stable_count = 0
                            while time.time() - start_dl_wait < 90:
                                time.sleep(1.5)
                                # Check if browser is actively downloading chunks
                                cr_files = (
                                    list(download_scratch.glob("*.crdownload"))
                                    + list(download_scratch.glob("*.tmp"))
                                    + list(download_scratch.glob("*.download"))
                                )
                                if cr_files:
                                    continue

                                cand_files = [
                                    f for f in download_scratch.glob("*")
                                    if f.is_file()
                                    and not f.name.endswith(".crdownload")
                                    and not f.name.endswith(".tmp")
                                    and not f.name.endswith(".download")
                                    and not f.name.endswith(".pure.xlsx")
                                ]
                                if cand_files:
                                    cand = cand_files[0]
                                    curr_size = cand.stat().st_size
                                    if curr_size > 1000:
                                        if curr_size == last_size:
                                            stable_count += 1
                                        else:
                                            stable_count = 0
                                            last_size = curr_size

                                        # Verify file is completely written and unlocked
                                        if stable_count >= 1:
                                            try:
                                                import zipfile
                                                if zipfile.is_zipfile(cand):
                                                    with open(cand, "rb") as tf:
                                                        tf.seek(0)
                                                    downloaded_path = cand
                                                    break
                                            except Exception:
                                                pass

                            if not downloaded_path:
                                raise TimeoutError(f"Quá thời gian chờ tải file BOM cho {part}!")

                            # Convert macro-enabled Excel export to pure .xlsx
                            clean_xlsx = download_scratch / f"PLM_{part}.pure.xlsx"
                            if is_macro_file(downloaded_path):
                                self.log_message.emit("   [*] Phát hiện file chứa macro/vba từ TC24, đang chuyển đổi sang pure .xlsx...")
                                conv_ok = convert_to_pure_xlsx(downloaded_path, clean_xlsx)
                                if conv_ok and clean_xlsx.exists() and clean_xlsx.stat().st_size > 0:
                                    raw_source = clean_xlsx
                                else:
                                    self.log_message.emit("   [!] Chuyển đổi macro không khả dụng, giữ nguyên file gốc.")
                                    raw_source = downloaded_path
                            else:
                                raw_source = downloaded_path

                            # Backup raw
                            bak_dir.mkdir(parents=True, exist_ok=True)
                            import shutil
                            shutil.copy2(raw_source, bak_dir / f"PLM_{part}.xlsx")
                            self.log_message.emit(f"   [+] Đã lưu bản gốc trước khi lọc vào: {bak_dir_name}/PLM_{part}.xlsx")

                            # Standardize to 14 columns
                            orig_c, final_c = standardize_plm_file(
                                raw_source,
                                target_dest,
                                prune_electrical=self.tc_auto_standardize,
                            )
                            self.log_message.emit(f"[+] [TC24] Tải và chuẩn hóa thành công: {orig_c:,} dòng raw -> {final_c:,} dòng chuẩn.")
                            plm_success += 1

                        # Sync to engineer subfolder if exists
                        if self.dest_dir.exists():
                            for d in self.dest_dir.iterdir():
                                if d.is_dir() and d.name.upper().startswith(part.upper()):
                                    import shutil
                                    sub_file = d / f"PLM_{part}.xlsx"
                                    shutil.copy2(target_dest, sub_file)
                                    self.log_message.emit(f"[+] [TC24] Đã đồng bộ sang thư mục kỹ sư: {d.name}")
                                    break

                    except Exception as exc:
                        error_messages.append(f"TC24 ({part}): {exc}")
                        self.log_message.emit(f"[-] [TC24] LỖI khi xử lý mã {part}: {exc}")

            except Exception as outer_tc_exc:
                error_messages.append(f"TC24 (Kết nối/Đăng nhập): {outer_tc_exc}")
                self.log_message.emit(f"[-] [TC24] Sự cố kết nối Teamcenter: {outer_tc_exc}")
            finally:
                if tc_client:
                    try:
                        tc_client.close()
                    except Exception:
                        pass

        # ---------------------------------------------------------------------
        # 2. SAP R3 CS12 Multilevel BOM Download (Uses per-BOM or default date)
        # ---------------------------------------------------------------------
        if self.mode in (MODE_BOTH, MODE_SAP):
            self.log_message.emit(f"\n>> BƯỚC 2: TẢI TỪ SAP R3 (CS12) [Plant: {self.sap_plant}, Usage: {self.sap_bom_usage}]")
            conn_mgr = None
            sap_service = None
            try:
                try:
                    from src.automation.sap.connection import SAPConnectionManager
                    from src.automation.sap.cs12 import CS12Service
                    from src.automation.sap.models import CS12Params

                    self.log_message.emit("[*] [SAP R3] Đang kết nối tới SAP GUI (chế độ chạy ngầm / ẩn)...")
                    conn_mgr = SAPConnectionManager()
                    session = conn_mgr.get_or_create_session()

                    # Minimize SAP main window to prevent focus stealing
                    try:
                        wnd0 = session.findById("wnd[0]")
                        if hasattr(wnd0, "iconify"):
                            wnd0.iconify()
                        elif hasattr(wnd0, "Iconify"):
                            wnd0.Iconify()
                    except Exception:
                        pass

                    # Windows API minimize non-activating
                    try:
                        import ctypes
                        user32 = ctypes.windll.user32
                        def _min_sap_win(hwnd, _):
                            length = user32.GetWindowTextLengthW(hwnd)
                            if length > 0:
                                buf = ctypes.create_unicode_buffer(length + 1)
                                user32.GetWindowTextW(hwnd, buf, length + 1)
                                if "SAP" in buf.value or "cs12" in buf.value.lower():
                                    user32.ShowWindow(hwnd, 7)  # SW_SHOWMINNOACTIVE
                            return True
                        WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_int, ctypes.c_int)
                        user32.EnumWindows(WNDENUMPROC(_min_sap_win), 0)
                    except Exception:
                        pass

                    sap_service = CS12Service(session=session)
                    self.log_message.emit("[+] [SAP R3] Kết nối phiên SAP GUI thành công (chế độ ẩn hoàn toàn)!")
                except Exception as conn_exc:
                    sap_tip = "Gợi ý: Mở SAP GUI (đăng nhập vào hệ thống P1J) và kiểm tra mục Scripting đã Enable"
                    error_messages.append(f"SAP R3 (Kết nối GUI): {conn_exc} ({sap_tip})")
                    self.log_message.emit(f"[-] [SAP R3] Không thể kết nối SAP GUI: {conn_exc}")
                    self.log_message.emit(f"    ({sap_tip})")

                for idx, item in enumerate(self.items, 1):
                    part = item.part_number
                    eff_date = self.custom_dates_map.get(part, self.sap_default_date)
                    date_tag = f"Ngày riêng: {eff_date.strftime('%Y/%m/%d')}" if part in self.custom_dates_map else f"Ngày chung: {eff_date.strftime('%Y/%m/%d')}"

                    current_step += 1
                    pct = int((current_step - 1) / total_steps * 95) + 2
                    self.progress.emit(pct, f"[{current_step}/{total_steps}] SAP R3: Đang tải CS12 cho {part}...")
                    self.log_message.emit(f"\n--- [SAP R3] Xử lý mã ({idx}/{total_parts}): {part} | {date_tag} ---")

                    if sap_service is None:
                        error_messages.append(f"SAP R3 ({part}): Bỏ qua do không có kết nối SAP GUI.")
                        self.log_message.emit(f"[-] [SAP R3] Bỏ qua mã {part} do không có kết nối SAP GUI.")
                        continue

                    try:
                        params = CS12Params(
                            material=part,
                            plant=self.sap_plant,
                            bom_usage=self.sap_bom_usage,
                            alternative=self.sap_alternative,
                            valid_date=eff_date,
                            destination_dir=self.dest_dir,
                        )
                        export_result = sap_service.execute_cs12_and_export(params)
                        if export_result.success:
                            self.log_message.emit(
                                f"[+] [SAP R3] Xuất BOM CS12 thành công: {params.expected_filename} "
                                f"({export_result.execution_time_sec:.1f}s)"
                            )
                            sap_success += 1
                        else:
                            error_messages.append(f"SAP R3 ({part}): {export_result.error_message}")
                            self.log_message.emit(
                                f"[-] [SAP R3] Lỗi từ SAP khi xử lý mã {part}: {export_result.error_message}"
                            )
                    except Exception as sap_err:
                        error_messages.append(f"SAP R3 ({part}): {sap_err}")
                        self.log_message.emit(f"[-] [SAP R3] LỖI khi tải CS12 mã {part}: {sap_err}")
            finally:
                if conn_mgr is not None and self.sap_auto_logout:
                    try:
                        self.log_message.emit("\n[*] [SAP R3] Đang tự động đăng xuất và thoát tài khoản SAP R3 (/nex)...")
                        conn_mgr.logoff_and_exit(close_saplogon=True)
                        self.log_message.emit("[+] [SAP R3] Đã thoát tài khoản và đóng ứng dụng SAP GUI an toàn.")
                    except Exception as logout_err:
                        self.log_message.emit(f"[-] [SAP R3] Cảnh báo khi thoát SAP GUI: {logout_err}")

        # ---------------------------------------------------------------------
        # Final Summary & Accurate Outcome
        # ---------------------------------------------------------------------
        total_ops = total_parts * (2 if self.mode == MODE_BOTH else 1)
        successful_ops = (plm_success if self.mode in (MODE_BOTH, MODE_PLM) else 0) + \
                         (sap_success if self.mode in (MODE_BOTH, MODE_SAP) else 0)

        summary_lines = []
        if self.mode in (MODE_BOTH, MODE_PLM):
            summary_lines.append(f"• Teamcenter TC24: Thành công {plm_success}/{total_parts} mã BOM.")
        if self.mode in (MODE_BOTH, MODE_SAP):
            summary_lines.append(f"• SAP R3: Thành công {sap_success}/{total_parts} mã BOM.")

        full_summary = "\n".join(summary_lines)

        if successful_ops == 0:
            final_status = f"Thất bại (0/{total_ops} tác vụ hoàn thành)"
            self.progress.emit(100, final_status)
        elif successful_ops < total_ops:
            final_status = f"Hoàn thành một phần ({successful_ops}/{total_ops} tác vụ hoàn thành)"
            self.progress.emit(100, final_status)
        else:
            final_status = f"Hoàn tất thành công toàn bộ ({successful_ops}/{total_ops} tác vụ)"
            self.progress.emit(100, final_status)

        self.log_message.emit(f"\n=== TỔNG KẾT TIẾN TRÌNH ===\n{full_summary}")
        if error_messages:
            self.log_message.emit(f"\n[!] CHI TIẾT SỰ CỐ GẶP PHẢI ({len(error_messages)}):")
            for err in error_messages:
                self.log_message.emit(f"   * {err}")

        self.successful_ops = successful_ops
        self.total_ops = total_ops
        self.error_messages = error_messages

        has_success = (plm_success > 0) or (sap_success > 0)
        self.finished.emit(has_success, full_summary)


class PLMDownloadDialog(QDialog):
    """User-friendly, Non-Tech BOM Downloader for Siemens TC24 and SAP R3."""

    def __init__(
        self,
        parent: QWidget | None = None,
        default_dir: Path | None = None,
        initial_model: str | None = None,
    ) -> None:
        super().__init__(parent)
        self._machine_dict = MachineDictService()
        self._initial_model = initial_model
        self.dest_dir = Path(default_dir or VIRGO_DIR)
        self.thread: QThread | None = None
        self.worker: UnifiedBOMDownloadWorker | None = None
        self.current_items: List[BOMDownloadItem] = []
        self._detected_model_name: str | None = None

        self.setWindowTitle("Tải Tự Động BOM Đa Nguồn (Siemens TC24 & SAP R3)")
        self.resize(860, 640)
        self.setMinimumSize(780, 500)
        self._init_ui()

    def _init_ui(self) -> None:
        theme_mgr = get_theme_manager()
        self.setWindowIcon(theme_mgr.get_styled_icon("download"))

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(12, 12, 12, 12)
        main_layout.setSpacing(8)

        # ---------------------------------------------------------------------
        # Scroll Area: Chứa Hướng dẫn, Ô nhập BOM, Cấu hình & Nhật ký
        # Giúp co giãn linh hoạt trên mọi độ phân giải màn hình mà không bị che khuất
        # ---------------------------------------------------------------------
        scroll_area = QScrollArea(self)
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.Shape.NoFrame)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        scroll_content = QWidget()
        layout = QVBoxLayout(scroll_content)
        layout.setContentsMargins(2, 2, 6, 2)
        layout.setSpacing(10)

        # Header with non-tech explanation
        header_box = QGroupBox()
        header_box.setStyleSheet("background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px;")
        hb_layout = QVBoxLayout(header_box)
        hb_layout.setContentsMargins(12, 8, 12, 8)
        hb_layout.setSpacing(4)

        title_lbl = QLabel("<b>HƯỚNG DẪN 3 BƯỚC TẢI BOM TỰ ĐỘNG</b>")
        title_lbl.setStyleSheet("color: #1e293b; font-size: 13px;")
        hb_layout.addWidget(title_lbl)

        guide_lbl = QLabel(
            "<b>Bước 1:</b> Dán danh sách mã máy/BOM vào ô bên dưới.<br>"
            "<b>Bước 2:</b> Chọn ngày hiệu lực nếu cần tải từ SAP R3 <i>(Teamcenter TC24 tự động lấy BOM mới nhất, không cần ngày)</i>.<br>"
            "<b>Bước 3:</b> Bấm nút màu xanh <b>'Tải đồng thời cả PLM & SAP R3'</b> ở thanh cố định phía dưới để tải."
        )
        guide_lbl.setStyleSheet("color: #475569; font-size: 11.5px; line-height: 1.4;")
        hb_layout.addWidget(guide_lbl)
        layout.addWidget(header_box)

        # ---------------------------------------------------------------------
        # Bước 1: Danh sách mã máy
        # ---------------------------------------------------------------------
        part_group = QGroupBox("Bước 1: Nhập hoặc dán danh sách mã máy / mã BOM")
        part_layout = QVBoxLayout(part_group)

        # Model Selector and Auto-detection Row
        model_row = QHBoxLayout()
        model_lbl = QLabel("Dòng máy / Model:")
        model_lbl.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
        model_row.addWidget(model_lbl)

        self.combo_model = QComboBox()
        self.combo_model.setFont(QFont("Calibri", 10))
        self.combo_model.addItem("-- Tự động nhận diện theo mã BOM --")
        try:
            for m_name in self._machine_dict.get_model_names():
                self.combo_model.addItem(m_name)
        except Exception as exc:
            logger.warning("Could not load model names: %s", exc)

        if self._initial_model:
            idx = self.combo_model.findText(self._initial_model)
            if idx >= 0:
                self.combo_model.setCurrentIndex(idx)
            else:
                self.combo_model.addItem(self._initial_model)
                self.combo_model.setCurrentText(self._initial_model)

        self.combo_model.currentIndexChanged.connect(self._on_model_selection_changed)
        model_row.addWidget(self.combo_model, stretch=1)

        self.lbl_model_badge = QLabel("")
        self.lbl_model_badge.setStyleSheet(
            "color: #0369a1; background-color: #e0f2fe; padding: 3px 8px; "
            "border-radius: 4px; border: 1px solid #bae6fd; font-size: 11px;"
        )
        self.lbl_model_badge.setVisible(False)
        model_row.addWidget(self.lbl_model_badge)

        part_layout.addLayout(model_row)

        # Toolbar above text edit
        tools_layout = QHBoxLayout()
        self.lbl_count = QLabel("Đã nhận diện: <b>0</b> mã BOM")
        self.lbl_count.setStyleSheet("color: #64748b; font-size: 12px;")
        tools_layout.addWidget(self.lbl_count)
        tools_layout.addStretch()

        self.btn_paste = QPushButton(" Dán từ Clipboard")
        self.btn_paste.setIcon(theme_mgr.get_styled_icon("clipboard"))
        self.btn_paste.setStyleSheet("padding: 4px 10px; font-weight: bold;")
        self.btn_paste.clicked.connect(self._paste_clipboard)
        tools_layout.addWidget(self.btn_paste)

        self.btn_clear = QPushButton(" Xóa hết")
        self.btn_clear.setIcon(theme_mgr.get_styled_icon("trash-2"))
        self.btn_clear.setStyleSheet("padding: 4px 10px;")
        self.btn_clear.clicked.connect(self._clear_parts)
        tools_layout.addWidget(self.btn_clear)

        part_layout.addLayout(tools_layout)

        # Multi-line PlainTextEdit
        self.txt_parts = QPlainTextEdit()
        self.txt_parts.setPlaceholderText(
            "Dán danh sách mã BOM vào đây (mỗi mã 1 dòng).\n"
            "Ví dụ:\n"
            "110C113NL0\n"
            "110C122US0\n"
            "110C132NL0"
        )
        self.txt_parts.setFont(QFont("Consolas", 10))
        self.txt_parts.setMinimumHeight(75)
        self.txt_parts.setMaximumHeight(110)
        self.txt_parts.textChanged.connect(self._on_text_changed)
        part_layout.addWidget(self.txt_parts)

        layout.addWidget(part_group)

        # ---------------------------------------------------------------------
        # Bước 2: Cấu hình tải
        # ---------------------------------------------------------------------
        config_group = QGroupBox("Bước 2: Cấu hình nguồn tải & Thư mục lưu")
        config_layout = QVBoxLayout(config_group)

        self.tab_config = QTabWidget()

        # Tab: Siemens Teamcenter (TC24)
        tc_widget = QWidget()
        tc_layout = QVBoxLayout(tc_widget)
        tc_form = QFormLayout()
        tc_form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)

        self.combo_account = QComboBox()
        self.combo_account.addItems([
            "vn_pe02 : 製造技術2課 (KTCT_Trọng)",
            "vn_pe01 : Cơ 1",
            "vn_pe03 : Phát triển hệ thống",
            "vn_pe04 : Kỹ thuật chế tạo điện",
        ])
        tc_form.addRow("Tài khoản tải TC24:", self.combo_account)

        self.chk_standardize = QCheckBox("Tự động chuẩn hóa 14 cột tiêu chuẩn")
        self.chk_standardize.setChecked(True)
        tc_form.addRow("Quy chuẩn PLM:", self.chk_standardize)
        tc_layout.addLayout(tc_form)

        # Non-tech note for PLM
        plm_note = QLabel("<b>Quy tắc Teamcenter (TC24):</b> Hệ thống luôn tự động tải cây cấu trúc BOM mới nhất. Bạn không cần thiết lập ngày hiệu lực cho PLM.")
        plm_note.setStyleSheet("color: #0369a1; font-size: 11px; background-color: #f0f9ff; padding: 6px; border-radius: 4px; border: 1px solid #bae6fd;")
        tc_layout.addWidget(plm_note)
        tc_layout.addStretch()

        self.tab_config.addTab(tc_widget, "Siemens Teamcenter (TC24)")

        # Tab: SAP R3
        sap_widget = QWidget()
        sap_layout = QVBoxLayout(sap_widget)

        sap_form = QFormLayout()
        sap_form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)

        self.edit_sap_plant = QLineEdit("2200")
        sap_form.addRow("Mã nhà máy (Plant):", self.edit_sap_plant)

        self.edit_sap_usage = QLineEdit("pp01")
        sap_form.addRow("BOM Usage:", self.edit_sap_usage)

        self.edit_sap_alt = QLineEdit("01")
        sap_form.addRow("Alternative (Alt):", self.edit_sap_alt)

        self.date_sap_valid = QDateEdit(QDate.currentDate())
        self.date_sap_valid.setDisplayFormat("yyyy/MM/dd")
        self.date_sap_valid.setCalendarPopup(True)
        self.date_sap_valid.dateChanged.connect(self._on_default_date_changed)
        sap_form.addRow("Ngày hiệu lực:", self.date_sap_valid)

        sap_layout.addLayout(sap_form)

        # Advanced Option: Checkbox to toggle per-BOM date table
        self.chk_custom_dates = QCheckBox("Nhập ngày hiệu lực riêng cho từng mã (Chỉ bật khi các mã cần ngày khác nhau)")
        self.chk_custom_dates.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
        self.chk_custom_dates.setStyleSheet("color: #b45309; margin-top: 4px;")
        self.chk_custom_dates.toggled.connect(self._on_toggle_custom_dates)
        sap_layout.addWidget(self.chk_custom_dates)

        # Embedded Date Allocation Table (Hidden by default, shown when checked)
        self.date_table = QTableWidget(0, 3)
        self.date_table.setShowGrid(True)
        self.date_table.verticalHeader().setDefaultSectionSize(32)
        self.date_table.setStyleSheet("QTableWidget { gridline-color: #CBD5E1; border: 1px solid #CBD5E1; }")
        self.date_table.setHorizontalHeaderLabels(["STT", "Mã máy / BOM", "Ngày hiệu lực SAP R3"])
        self.date_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.date_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.date_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.date_table.setMaximumHeight(140)
        self.date_table.setVisible(False)
        sap_layout.addWidget(self.date_table)

        sap_note = QLabel("<b>Quy tắc SAP R3:</b> Ngày hiệu lực dùng để truy xuất cấu trúc BOM đa tầng theo mốc thời gian đã chọn.")
        sap_note.setStyleSheet("color: #4d7c0f; font-size: 11px; background-color: #f7fee7; padding: 6px; border-radius: 4px; border: 1px solid #d9f99d;")
        sap_layout.addWidget(sap_note)

        self.chk_sap_auto_logout = QCheckBox("Tự động thoát tài khoản & đóng SAP GUI sau khi tải xong")
        self.chk_sap_auto_logout.setChecked(True)
        self.chk_sap_auto_logout.setStyleSheet("color: #15803d; font-weight: bold; margin-top: 4px;")
        sap_layout.addWidget(self.chk_sap_auto_logout)

        self.tab_config.addTab(sap_widget, "SAP R3")
        config_layout.addWidget(self.tab_config)

        # Destination Directory Picker (Shared)
        dest_layout = QHBoxLayout()
        dest_lbl = QLabel("Thư mục lưu kết quả:")
        dest_lbl.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
        dest_layout.addWidget(dest_lbl)

        self.edit_dest = QLineEdit(str(self.dest_dir))
        self.edit_dest.setFont(QFont("Consolas", 9))
        self.btn_browse = QPushButton(" Duyệt...")
        self.btn_browse.setIcon(theme_mgr.get_styled_icon("folder"))
        self.btn_browse.clicked.connect(self._browse_dest_dir)
        dest_layout.addWidget(self.edit_dest)
        dest_layout.addWidget(self.btn_browse)
        config_layout.addLayout(dest_layout)

        self.lbl_backup_preview = QLabel()
        self.lbl_backup_preview.setStyleSheet("color: #475569; font-size: 11px; margin-top: 2px;")
        config_layout.addWidget(self.lbl_backup_preview)
        self._update_backup_preview()

        layout.addWidget(config_group)

        # Nhật ký hoạt động chi tiết
        exec_group = QGroupBox("Nhật ký chi tiết hoạt động")
        exec_layout = QVBoxLayout(exec_group)

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setFont(QFont("Consolas", 10))
        self.log_box.setStyleSheet("background-color: #F8FAFC; border: 1px solid #CBD5E1; color: #0F172A; font-family: Consolas, monospace;")
        self.log_box.setMinimumHeight(65)
        self.log_box.setMaximumHeight(95)
        exec_layout.addWidget(self.log_box)

        layout.addWidget(exec_group)

        scroll_area.setWidget(scroll_content)
        main_layout.addWidget(scroll_area, stretch=1)

        # ---------------------------------------------------------------------
        # THANH CỐ ĐỊNH PHÍA ĐÁY CỬA SỔ (PINNED FOOTER):
        # Luôn hiển thị 100% trên mọi độ phân giải màn hình (không bao giờ bị che khuất)
        # ---------------------------------------------------------------------
        bottom_dock = QWidget()
        bottom_dock_layout = QVBoxLayout(bottom_dock)
        bottom_dock_layout.setContentsMargins(0, 2, 0, 0)
        bottom_dock_layout.setSpacing(4)

        # Thanh tiến độ và nhãn trạng thái trực quan
        status_prog_layout = QHBoxLayout()
        self.lbl_status = QLabel("Trạng thái: Sẵn sàng")
        self.lbl_status.setStyleSheet("font-style: italic; color: #475569; font-size: 11px;")
        status_prog_layout.addWidget(self.lbl_status, stretch=1)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFixedHeight(14)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("%p%")
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                text-align: center;
                height: 14px;
                font-size: 10px;
                font-weight: bold;
                background-color: #F1F5F9;
            }
            QProgressBar::chunk {
                background-color: #2563EB;
                border-radius: 3px;
            }
        """)
        status_prog_layout.addWidget(self.progress_bar, stretch=1)
        bottom_dock_layout.addLayout(status_prog_layout)

        # Bước 3: Nút bấm thực thi
        btn_box = QGroupBox("Bước 3: Chọn lệnh tải tự động")
        btn_layout = QHBoxLayout(btn_box)

        # Button 1: Download BOTH (Primary - Highly visible)
        self.btn_both = QPushButton(" Tải đồng thời cả PLM && SAP R3")
        self.btn_both.setIcon(theme_mgr.get_styled_icon("download", color="#FFFFFF"))
        self.btn_both.setFont(QFont("Calibri", 11, QFont.Weight.Bold))
        self.btn_both.setStyleSheet(
            "background-color: #2563EB; color: white; padding: 10px 20px; border-radius: 5px;"
        )
        self.btn_both.clicked.connect(lambda: self._trigger_download(MODE_BOTH))
        btn_layout.addWidget(self.btn_both)

        # Button 2: PLM Only
        self.btn_plm_only = QPushButton(" Chỉ tải BOM PLM (TC24)")
        self.btn_plm_only.setIcon(theme_mgr.get_styled_icon("download", color="#FFFFFF"))
        self.btn_plm_only.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
        self.btn_plm_only.setStyleSheet(
            "background-color: #0284C7; color: white; padding: 10px 14px; border-radius: 5px;"
        )
        self.btn_plm_only.clicked.connect(lambda: self._trigger_download(MODE_PLM))
        btn_layout.addWidget(self.btn_plm_only)

        # Button 3: SAP R3 Only
        self.btn_sap_only = QPushButton(" Chỉ tải BOM SAP R3")
        self.btn_sap_only.setIcon(theme_mgr.get_styled_icon("download", color="#FFFFFF"))
        self.btn_sap_only.setFont(QFont("Calibri", 10, QFont.Weight.Bold))
        self.btn_sap_only.setStyleSheet(
            "background-color: #059669; color: white; padding: 10px 14px; border-radius: 5px;"
        )
        self.btn_sap_only.clicked.connect(lambda: self._trigger_download(MODE_SAP))
        btn_layout.addWidget(self.btn_sap_only)

        self.btn_close = QPushButton("Đóng")
        self.btn_close.setFont(QFont("Calibri", 10))
        self.btn_close.setStyleSheet("padding: 10px 16px;")
        self.btn_close.clicked.connect(self.close)
        btn_layout.addWidget(self.btn_close)

        bottom_dock_layout.addWidget(btn_box)
        main_layout.addWidget(bottom_dock)

    # -------------------------------------------------------------------------
    # Helper & Event Handlers
    # -------------------------------------------------------------------------

    def _get_current_default_date(self) -> date:
        qdate = self.date_sap_valid.date()
        return date(qdate.year(), qdate.month(), qdate.day())

    def get_current_model_name(self) -> str:
        txt = self.combo_model.currentText().strip()
        if not txt or txt.startswith("--"):
            return self._detected_model_name or "Virgo"
        return txt

    def _update_backup_preview(self) -> None:
        model_name = self.get_current_model_name()
        model_slug = re.sub(r"[^\w]+", "_", model_name.strip().lower()).strip("_")
        bak_name = f"backup_before_{model_slug}_filter"
        if hasattr(self, "lbl_backup_preview"):
            self.lbl_backup_preview.setText(
                f"Thư mục sao lưu file gốc (trước khi lọc): <b style='color: #0284c7;'>{bak_name}/</b>"
            )

    def _on_model_selection_changed(self, index: int) -> None:
        self._update_backup_preview()
        model = self.get_current_model_name()
        if SO_SANH_BASE_DIR.exists():
            candidate = SO_SANH_BASE_DIR / model
            try:
                if self.dest_dir.parent == SO_SANH_BASE_DIR or self.dest_dir == SO_SANH_BASE_DIR:
                    if candidate.exists():
                        self.dest_dir = candidate
                        self.edit_dest.setText(str(self.dest_dir))
            except Exception:
                pass

    def _on_text_changed(self) -> None:
        """Update counter, auto-detect model from machine dictionary, and sync date table."""
        self.current_items = parse_bom_items(self.txt_parts.toPlainText())
        count = len(self.current_items)

        if count == 0:
            self.lbl_count.setText("Đã nhận diện: <b>0</b> mã BOM")
            self.lbl_count.setStyleSheet("color: #64748b; font-size: 12px;")
            self.lbl_model_badge.setVisible(False)
            self._detected_model_name = None
        else:
            self.lbl_count.setText(f"Đã nhận diện: <b>{count}</b> mã BOM hợp lệ")
            self.lbl_count.setStyleSheet("color: #16a34a; font-size: 12px; font-weight: bold;")

            # Auto-detect machine model using MachineDictService
            detected_info = None
            for item in self.current_items:
                info = self._machine_dict.extract_and_lookup_material(item.part_number)
                if info and info.machine_name:
                    detected_info = info
                    break

            if detected_info:
                self._detected_model_name = detected_info.machine_name
                matched_codes = ", ".join(detected_info.machine_codes) if detected_info.machine_codes else ""
                badge_text = f"✓ Nhận diện từ file_loaimay: <b>{detected_info.machine_name}</b>"
                if matched_codes:
                    badge_text += f" (Mã: {matched_codes})"
                self.lbl_model_badge.setText(badge_text)
                self.lbl_model_badge.setStyleSheet(
                    "color: #065f46; background-color: #d1fae5; padding: 3px 8px; "
                    "border-radius: 4px; border: 1px solid #a7f3d0; font-size: 11px;"
                )
                self.lbl_model_badge.setVisible(True)

                # If combo is on Auto, select the detected model
                if self.combo_model.currentIndex() == 0:
                    idx = self.combo_model.findText(detected_info.machine_name)
                    if idx >= 0:
                        self.combo_model.blockSignals(True)
                        self.combo_model.setCurrentIndex(idx)
                        self.combo_model.blockSignals(False)
                        self._on_model_selection_changed(idx)
            else:
                if self.combo_model.currentIndex() == 0:
                    self.lbl_model_badge.setText("Chưa tìm thấy mã máy tương ứng trong file_loaimay")
                    self.lbl_model_badge.setStyleSheet(
                        "color: #92400e; background-color: #fef3c7; padding: 3px 8px; "
                        "border-radius: 4px; border: 1px solid #fde68a; font-size: 11px;"
                    )
                    self.lbl_model_badge.setVisible(True)

        self._update_backup_preview()

        # Synchronize table if currently visible
        if self.chk_custom_dates.isChecked():
            self._populate_date_table()

    def _on_toggle_custom_dates(self, checked: bool) -> None:
        """Toggle visibility of per-BOM date configuration table."""
        self.date_table.setVisible(checked)
        if checked:
            self._populate_date_table()

    def _on_default_date_changed(self, new_qdate: QDate) -> None:
        """Refresh default dates in table for items without custom dates."""
        if self.chk_custom_dates.isChecked():
            self._populate_date_table()

    def _populate_date_table(self) -> None:
        """Populate embedded date table with recognized items and date controls."""
        self.date_table.setRowCount(len(self.current_items))
        def_date = self._get_current_default_date()

        for row, itm in enumerate(self.current_items):
            # STT
            stt_item = QTableWidgetItem(str(row + 1))
            stt_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            stt_item.setFlags(stt_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.date_table.setItem(row, 0, stt_item)

            # Part Number
            part_item = QTableWidgetItem(itm.part_number)
            part_item.setFont(QFont("Consolas", 10, QFont.Weight.Bold))
            part_item.setFlags(part_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.date_table.setItem(row, 1, part_item)

            # Date Edit widget inside cell
            date_widget = QDateEdit()
            date_widget.setDisplayFormat("yyyy/MM/dd")
            date_widget.setCalendarPopup(True)
            active_dt = itm.custom_date or def_date
            date_widget.setDate(QDate(active_dt.year, active_dt.month, active_dt.day))
            self.date_table.setCellWidget(row, 2, date_widget)

    def _get_effective_custom_dates(self) -> Dict[str, date]:
        """Retrieve custom dates from table if checkbox is checked."""
        dates_map: Dict[str, date] = {}
        if not self.chk_custom_dates.isChecked():
            # If not checked, but user pasted dates in text, respect them
            for itm in self.current_items:
                if itm.custom_date:
                    dates_map[itm.part_number] = itm.custom_date
            return dates_map

        for row in range(self.date_table.rowCount()):
            part_item = self.date_table.item(row, 1)
            date_widget = self.date_table.cellWidget(row, 2)
            if part_item and isinstance(date_widget, QDateEdit):
                part = part_item.text().strip()
                qd = date_widget.date()
                dates_map[part] = date(qd.year(), qd.month(), qd.day())
        return dates_map

    def _paste_clipboard(self) -> None:
        clipboard = QGuiApplication.clipboard()
        text = clipboard.text()
        if text:
            cursor = self.txt_parts.textCursor()
            cursor.insertText(text)

    def _clear_parts(self) -> None:
        self.txt_parts.clear()

    def _browse_dest_dir(self) -> None:
        folder = QFileDialog.getExistingDirectory(
            self,
            "Chọn thư mục lưu BOM",
            self.edit_dest.text().strip() or str(self.dest_dir),
        )
        if folder:
            self.edit_dest.setText(folder)
            self.dest_dir = Path(folder)

    def _set_ui_busy(self, busy: bool) -> None:
        self.btn_both.setEnabled(not busy)
        self.btn_plm_only.setEnabled(not busy)
        self.btn_sap_only.setEnabled(not busy)
        self.btn_browse.setEnabled(not busy)
        self.btn_paste.setEnabled(not busy)
        self.btn_clear.setEnabled(not busy)
        self.chk_custom_dates.setEnabled(not busy)

    def _trigger_download(self, mode: str) -> None:
        """Validate input and trigger background worker thread with specified mode."""
        self.current_items = parse_bom_items(self.txt_parts.toPlainText())
        if not self.current_items:
            QMessageBox.warning(
                self,
                "Chưa nhập mã BOM",
                "Vui lòng dán danh sách ít nhất một mã máy/BOM vào ô ở Bước 1!",
            )
            return

        dest_path_str = self.edit_dest.text().strip()
        if not dest_path_str:
            QMessageBox.warning(self, "Chưa chọn thư mục", "Vui lòng chọn thư mục lưu BOM!")
            return

        self.dest_dir = Path(dest_path_str)
        try:
            self.dest_dir.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            QMessageBox.critical(
                self,
                "Lỗi thư mục",
                f"Không thể tạo hoặc truy cập thư mục lưu trữ:\n{exc}",
            )
            return

        tc_account = self.combo_account.currentText().split(":")[0].strip()
        tc_auto_std = self.chk_standardize.isChecked()

        sap_plant = self.edit_sap_plant.text().strip() or "2200"
        sap_usage = self.edit_sap_usage.text().strip() or "pp01"
        sap_alt = self.edit_sap_alt.text().strip() or "01"
        sap_default_dt = self._get_current_default_date()
        custom_dates = self._get_effective_custom_dates()

        # Lock UI & Reset status styling to active
        self._set_ui_busy(True)
        self._download_start_time = time.time()
        self.log_box.clear()
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("0%")
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                text-align: center;
                height: 14px;
                font-size: 10px;
                font-weight: bold;
                background-color: #F1F5F9;
                color: #1E293B;
            }
            QProgressBar::chunk {
                background-color: #2563EB;
                border-radius: 3px;
            }
        """)
        self.lbl_status.setStyleSheet("font-style: italic; color: #475569; font-size: 11px;")
        self.lbl_status.setText(f"Đang chuẩn bị tải BOM (Chế độ: {mode})...")

        # Start background worker
        self.thread = QThread()
        self.worker = UnifiedBOMDownloadWorker(
            items=self.current_items,
            mode=mode,
            tc_account=tc_account,
            tc_auto_standardize=tc_auto_std,
            sap_plant=sap_plant,
            sap_bom_usage=sap_usage,
            sap_alternative=sap_alt,
            sap_default_date=sap_default_dt,
            custom_dates_map=custom_dates,
            dest_dir=self.dest_dir,
            model_name=self.get_current_model_name(),
            sap_auto_logout=self.chk_sap_auto_logout.isChecked(),
        )
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self._on_progress)
        self.worker.log_message.connect(self._on_log)
        self.worker.finished.connect(self._on_finished)

        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)

        self.thread.start()

    def _on_progress(self, pct: int, msg: str) -> None:
        self.progress_bar.setValue(pct)
        rem_str = ""
        if hasattr(self, "_download_start_time") and 0 < pct < 100:
            elapsed = time.time() - self._download_start_time
            total_est = elapsed / (pct / 100.0)
            remaining = max(0, int(total_est - elapsed))
            rem_str = f" (~{remaining}s còn lại)"
        self.progress_bar.setFormat(f"{pct}%{rem_str}")
        self.lbl_status.setText(f"Trạng thái: {msg} | Tiến độ: {pct}%{rem_str}")

    def _on_log(self, text: str) -> None:
        self.log_box.append(text)
        self.log_box.moveCursor(QTextCursor.MoveOperation.End)

    def _on_finished(self, success: bool, msg: str) -> None:
        self._set_ui_busy(False)
        successful_ops = getattr(self.worker, "successful_ops", 0) if self.worker else (1 if success else 0)
        total_ops = getattr(self.worker, "total_ops", 1) if self.worker else 1
        error_messages = getattr(self.worker, "error_messages", []) if self.worker else []

        if successful_ops == 0:
            # 1. Total Failure (Red)
            self.progress_bar.setStyleSheet("""
                QProgressBar {
                    border: 1px solid #DC2626;
                    border-radius: 4px;
                    text-align: center;
                    height: 14px;
                    font-size: 10px;
                    font-weight: bold;
                    background-color: #FEF2F2;
                    color: #991B1B;
                }
                QProgressBar::chunk {
                    background-color: #EF4444;
                    border-radius: 3px;
                }
            """)
            self.progress_bar.setFormat("Thất bại (0 mã thành công)")
            self.lbl_status.setStyleSheet("color: #DC2626; font-weight: bold; font-size: 11px;")
            self.lbl_status.setText(f"Trạng thái: Thất bại (0/{total_ops} tác vụ hoàn thành) | Xem chi tiết lỗi bên dưới.")

            err_text = "\n".join(f"• {e}" for e in error_messages[:5]) if error_messages else "Không thể kết nối hoặc xuất dữ liệu."
            QMessageBox.critical(
                self,
                "Tải BOM Thất bại",
                f"Quá trình tải BOM thất bại hoàn toàn (0/{total_ops} tác vụ thành công):\n\n"
                f"{msg}\n\n"
                f"Chi tiết sự cố:\n{err_text}\n\n"
                "Vui lòng xem thêm nhật ký bên dưới để khắc phục.",
            )
        elif successful_ops < total_ops:
            # 2. Partial Success (Amber/Orange)
            self.progress_bar.setStyleSheet("""
                QProgressBar {
                    border: 1px solid #D97706;
                    border-radius: 4px;
                    text-align: center;
                    height: 14px;
                    font-size: 10px;
                    font-weight: bold;
                    background-color: #FFFBEB;
                    color: #92400E;
                }
                QProgressBar::chunk {
                    background-color: #F59E0B;
                    border-radius: 3px;
                }
            """)
            self.progress_bar.setFormat(f"Một phần ({successful_ops}/{total_ops} thành công)")
            self.lbl_status.setStyleSheet("color: #D97706; font-weight: bold; font-size: 11px;")
            self.lbl_status.setText(f"Trạng thái: Hoàn thành một phần ({successful_ops}/{total_ops} tác vụ) | Xem chi tiết lỗi bên dưới.")

            err_text = "\n".join(f"• {e}" for e in error_messages[:5]) if error_messages else ""
            err_suffix = f"\n\nSự cố gặp phải:\n{err_text}" if err_text else ""
            QMessageBox.warning(
                self,
                "Tải BOM Một phần",
                f"Quá trình tải BOM hoàn thành một phần ({successful_ops}/{total_ops} tác vụ thành công):\n\n"
                f"{msg}{err_suffix}\n\n"
                "Vui lòng kiểm tra nhật ký chi tiết đối với các mã bị lỗi.",
            )
        else:
            # 3. Full Success (Green)
            self.progress_bar.setStyleSheet("""
                QProgressBar {
                    border: 1px solid #059669;
                    border-radius: 4px;
                    text-align: center;
                    height: 14px;
                    font-size: 10px;
                    font-weight: bold;
                    background-color: #F0FDF4;
                    color: #065F46;
                }
                QProgressBar::chunk {
                    background-color: #10B981;
                    border-radius: 3px;
                }
            """)
            self.progress_bar.setFormat(f"100% (Hoàn tất thành công {total_ops}/{total_ops})")
            self.lbl_status.setStyleSheet("color: #059669; font-weight: bold; font-size: 11px;")
            self.lbl_status.setText(f"Trạng thái: Hoàn tất thành công toàn bộ ({total_ops}/{total_ops} tác vụ)!")
            QMessageBox.information(
                self,
                "Tải BOM Hoàn tất",
                f"Tiến trình tải BOM hoàn tất thành công 100%!\n\n{msg}",
            )


# Backwards compatibility alias
BOMDownloadDialog = PLMDownloadDialog
