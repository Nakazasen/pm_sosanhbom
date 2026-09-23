"""Unified Automated BOM Download Dialog for Siemens Teamcenter TC24 & SAP R3.

Designed specifically for non-tech manufacturing engineers and team leads:
- Simple 3-step workflow (Paste BOM list -> Click download).
- Clear separation: Effective date applies ONLY to SAP R3; TC24 PLM always pulls latest.
- Optional intuitive per-BOM date configuration table (only displayed when requested).
"""

from __future__ import annotations

import json
import logging
import re
import time
import zipfile
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Dict, List, Optional

from src.services.machine_dict_service import MachineDictService

from PyQt6.QtCore import QDate, QObject, QSize, QThread, QTimer, Qt, pyqtSignal, pyqtSlot
from PyQt6.QtGui import QColor, QFont, QGuiApplication, QTextCursor
from PyQt6.QtWidgets import (
    QButtonGroup,
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
    QRadioButton,
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

DEFAULT_CONFIG_PATH = Path("config/settings.json")
FALLBACK_CONFIG_PATH = Path.home() / ".ssbom" / "config.json"

SO_SANH_BASE_DIR = Path(
    r"\\fstvn01\Data\10_Production Engineering Department(製造技術部)\02.製造技術課\PE Dept\4A. QUAN LY BOM-TDTK-BOM管理-設計変更\SO SANH PLM-CTTT-R3"
)
VIRGO_DIR = SO_SANH_BASE_DIR / "Virgo2"

MODE_BOTH = "BOTH"
MODE_PLM = "PLM"
MODE_SAP = "SAP"


def normalize_model_key(model_name: str | None) -> str:
    """Normalize model string by removing all whitespace and uppercasing."""
    if not model_name:
        return ""
    return "".join(model_name.split()).upper()


def get_persisted_download_dir(
    model_name: str | None = None,
    fallback_dir: Path | None = None,
    config_path: Path | None = None,
) -> Path:
    """Retrieve persisted download directory for a model from settings.json.

    Checks per-model saved directory first (paths.last_download_dirs_by_model),
    then falls back to paths.last_download_dir.
    Verifies that the candidate directory exists and is a directory. If not accessible
    (e.g., offline LAN share, deleted directory), gracefully returns fallback_dir.
    """
    default_fallback = Path(fallback_dir or VIRGO_DIR)
    target_config = config_path or (
        DEFAULT_CONFIG_PATH if DEFAULT_CONFIG_PATH.exists() else FALLBACK_CONFIG_PATH
    )
    if not target_config.exists():
        return default_fallback

    try:
        with open(target_config, "r", encoding="utf-8") as fp:
            data = json.load(fp)
    except Exception as exc:
        logger.warning("Could not read settings for download dir: %s", exc)
        return default_fallback

    paths_sec = data.get("paths", {})
    if not isinstance(paths_sec, dict):
        return default_fallback

    cand_str: Optional[str] = None

    # 1. Check per-model configuration
    if model_name:
        clean_target = "".join(model_name.split())
        norm_target = normalize_model_key(model_name)
        model_dirs = paths_sec.get("last_download_dirs_by_model", {})
        if isinstance(model_dirs, dict):
            if clean_target in model_dirs:
                cand_str = model_dirs[clean_target]
            elif model_name in model_dirs:
                cand_str = model_dirs[model_name]
            else:
                for k, v in model_dirs.items():
                    if normalize_model_key(k) == norm_target:
                        cand_str = v
                        break

    # 2. Check global last_download_dir fallback if no model match
    if not cand_str:
        global_last = paths_sec.get("last_download_dir")
        if isinstance(global_last, str) and global_last.strip():
            cand_str = global_last.strip()

    if cand_str:
        cand_p = Path(cand_str)
        try:
            if cand_p.exists() and cand_p.is_dir():
                return cand_p
            else:
                logger.info(
                    "Persisted directory %s is not accessible, using fallback %s",
                    cand_p,
                    default_fallback,
                )
        except Exception:
            pass

    return default_fallback


def save_persisted_download_dir(
    model_name: str | None,
    dest_dir: Path,
    config_path: Path | None = None,
) -> None:
    """Save the chosen download directory to settings.json atomically.

    Saves under both paths.last_download_dirs_by_model[clean_model] and paths.last_download_dir.
    Preserves all other configuration keys in settings.json.
    """
    target_config = config_path or (
        DEFAULT_CONFIG_PATH if DEFAULT_CONFIG_PATH.exists() else FALLBACK_CONFIG_PATH
    )
    data: dict = {}
    if target_config.exists():
        try:
            with open(target_config, "r", encoding="utf-8") as fp:
                data = json.load(fp)
        except Exception as exc:
            logger.warning("Could not read settings for updating download dir: %s", exc)
            data = {}

    if not isinstance(data, dict):
        data = {}

    if "paths" not in data or not isinstance(data["paths"], dict):
        data["paths"] = {}

    paths_sec = data["paths"]
    dest_str = str(dest_dir).strip()
    paths_sec["last_download_dir"] = dest_str

    if model_name:
        clean_model = "".join(model_name.split())
        if clean_model and not clean_model.startswith("--"):
            if "last_download_dirs_by_model" not in paths_sec or not isinstance(
                paths_sec["last_download_dirs_by_model"], dict
            ):
                paths_sec["last_download_dirs_by_model"] = {}
            paths_sec["last_download_dirs_by_model"][clean_model] = dest_str

    try:
        target_config.parent.mkdir(parents=True, exist_ok=True)
        tmp_file = target_config.with_suffix(".tmp")
        with open(tmp_file, "w", encoding="utf-8") as fp:
            json.dump(data, fp, indent=2, ensure_ascii=False)
        tmp_file.replace(target_config)
        logger.info("Saved persisted download directory %s for model %s", dest_str, model_name)
    except Exception as exc:
        logger.error("Failed to persist download dir to %s: %s", target_config, exc)


def is_bom_scan_suggestion_enabled(config_path: Path | None = None) -> bool:
    """Check if the smart suggestion banner should be shown when opening download dialog."""
    target_config = config_path or (
        DEFAULT_CONFIG_PATH if DEFAULT_CONFIG_PATH.exists() else FALLBACK_CONFIG_PATH
    )
    if not target_config.exists():
        return True
    try:
        with open(target_config, "r", encoding="utf-8") as fp:
            data = json.load(fp)
        return bool(data.get("paths", {}).get("show_bom_scan_suggestion", True))
    except Exception:
        return True


def set_bom_scan_suggestion_enabled(enabled: bool, config_path: Path | None = None) -> None:
    """Save user preference for showing the smart suggestion banner."""
    target_config = config_path or (
        DEFAULT_CONFIG_PATH if DEFAULT_CONFIG_PATH.exists() else FALLBACK_CONFIG_PATH
    )
    data: dict = {}
    if target_config.exists():
        try:
            with open(target_config, "r", encoding="utf-8") as fp:
                data = json.load(fp)
        except Exception:
            data = {}
    if not isinstance(data, dict):
        data = {}
    if "paths" not in data or not isinstance(data["paths"], dict):
        data["paths"] = {}
    data["paths"]["show_bom_scan_suggestion"] = enabled
    try:
        target_config.parent.mkdir(parents=True, exist_ok=True)
        tmp_file = target_config.with_suffix(".tmp")
        with open(tmp_file, "w", encoding="utf-8") as fp:
            json.dump(data, fp, indent=2, ensure_ascii=False)
        tmp_file.replace(target_config)
    except Exception as exc:
        logger.error("Failed to save show_bom_scan_suggestion to %s: %s", target_config, exc)


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
        if not arr_btns:
            arr_btns = driver.find_elements(By.CSS_SELECTOR, "button[command-id='Arm0ArrangeViewConfigs']")
        if arr_btns:
            arr_btns[0].click()
            time.sleep(2)

            menu_items = driver.find_elements(
                By.CSS_SELECTOR,
                "div.aw-popup div.aw-widgets-cellListItem, div.aw-popup li, div.sw-popup li, div.sw-popup div",
            )
            for m in menu_items:
                if "KTCT_Trong" in (m.text or ""):
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


PART_REGEX = re.compile(r"(?:PLM_|R3_)?([0-9]{3}[A-Z0-9]{7})", re.IGNORECASE)


@dataclass
class BOMPartStatus:
    """Represents a BOM part and the availability of its PLM and SAP R3 files."""

    part_number: str
    has_plm: bool = False
    plm_path: Optional[Path] = None
    plm_size: int = 0
    plm_mtime: Optional[float] = None
    has_r3: bool = False
    r3_path: Optional[Path] = None
    r3_size: int = 0
    r3_mtime: Optional[float] = None

    @property
    def needs_r3(self) -> bool:
        """True if PLM file exists but SAP R3 file is missing."""
        return self.has_plm and not self.has_r3

    @property
    def needs_plm(self) -> bool:
        """True if SAP R3 file exists but PLM file is missing."""
        return self.has_r3 and not self.has_plm

    @property
    def is_complete(self) -> bool:
        """True if both PLM and SAP R3 files exist."""
        return self.has_plm and self.has_r3

    @property
    def status_summary(self) -> str:
        if self.is_complete:
            return "✓ Đã đầy đủ cả 2"
        elif self.needs_r3:
            return "⚠️ Cần tải bù SAP R3"
        elif self.needs_plm:
            return "⚠️ Cần tải bù PLM"
        else:
            return "- Chưa tải file nào"


def extract_part_from_name(name: str) -> Optional[str]:
    """Extract 10-character machine part number from file or directory name."""
    if not name:
        return None
    m = PART_REGEX.search(name)
    if m:
        return m.group(1).upper()
    return None


def scan_bom_directory(
    folder: Path,
    extra_parts: Optional[List[str]] = None,
) -> List[BOMPartStatus]:
    """Scan directory for BOM files and subdirectories, analyzing PLM & R3 status.

    Scans:
    1. Direct files in folder: PLM_<part>.xlsx, R3_<part>.xlsx / .xls
    2. Subfolders named <part> or containing part codes
    3. Any extra_parts supplied (e.g. from active Project)
    """
    parts_map: Dict[str, BOMPartStatus] = {}

    # Initialize any extra parts (from project)
    if extra_parts:
        for p in extra_parts:
            clean_p = p.strip().upper()
            if clean_p and clean_p not in parts_map:
                parts_map[clean_p] = BOMPartStatus(part_number=clean_p)

    if not folder.exists() or not folder.is_dir():
        return sorted(parts_map.values(), key=lambda x: x.part_number)

    try:
        entries = list(folder.iterdir())
    except Exception as exc:
        logger.warning("Cannot list directory %s: %s", folder, exc)
        return sorted(parts_map.values(), key=lambda x: x.part_number)

    # 1. Inspect direct files in folder
    for item in entries:
        if item.is_file():
            if item.name.startswith("~$") or item.name.endswith(".tmp") or item.name.endswith(".pure.xlsx"):
                continue
            part = extract_part_from_name(item.name)
            if not part:
                continue
            if part not in parts_map:
                parts_map[part] = BOMPartStatus(part_number=part)
            st = parts_map[part]

            fname_upper = item.name.upper()
            try:
                sz = item.stat().st_size
                mtime = item.stat().st_mtime
            except Exception:
                sz = 0
                mtime = None

            if fname_upper.startswith("PLM") or "PLM_" in fname_upper:
                st.has_plm = True
                st.plm_path = item
                st.plm_size = sz
                st.plm_mtime = mtime
            elif fname_upper.startswith("R3") or "R3_" in fname_upper:
                st.has_r3 = True
                st.r3_path = item
                st.r3_size = sz
                st.r3_mtime = mtime

        elif item.is_dir():
            part = extract_part_from_name(item.name)
            if part:
                if part not in parts_map:
                    parts_map[part] = BOMPartStatus(part_number=part)
                st = parts_map[part]
                try:
                    for sf in item.glob("*.xls*"):
                        if sf.name.startswith("~$"):
                            continue
                        sf_upper = sf.name.upper()
                        try:
                            sz = sf.stat().st_size
                            mtime = sf.stat().st_mtime
                        except Exception:
                            sz = 0
                            mtime = None

                        if (sf_upper.startswith("PLM") or "PLM_" in sf_upper) and not st.has_plm:
                            st.has_plm = True
                            st.plm_path = sf
                            st.plm_size = sz
                            st.plm_mtime = mtime
                        elif (sf_upper.startswith("R3") or "R3_" in sf_upper) and not st.has_r3:
                            st.has_r3 = True
                            st.r3_path = sf
                            st.r3_size = sz
                            st.r3_mtime = mtime
                except Exception:
                    pass

    # 2. For any part in parts_map that doesn't have plm/r3 yet, do a quick targeted check
    for part, st in parts_map.items():
        if not st.has_plm:
            for cand in [folder / f"PLM_{part}.xlsx", folder / f"PLM_{part}.xls"]:
                if cand.exists():
                    st.has_plm = True
                    st.plm_path = cand
                    st.plm_size = cand.stat().st_size
                    st.plm_mtime = cand.stat().st_mtime
                    break
        if not st.has_r3:
            for cand in [folder / f"R3_{part}.xlsx", folder / f"R3_{part}.xls"]:
                if cand.exists():
                    st.has_r3 = True
                    st.r3_path = cand
                    st.r3_size = cand.stat().st_size
                    st.r3_mtime = cand.stat().st_mtime
                    break

    return sorted(parts_map.values(), key=lambda x: x.part_number)


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
        part_model_map: Optional[Dict[str, str]] = None,
        apply_bolocbom: bool = True,
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
        self.part_model_map = part_model_map or {}
        self.apply_bolocbom = apply_bolocbom
        self._machine_dict = MachineDictService()

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
                        part_model = self.part_model_map.get(part)
                        if not part_model:
                            info = self._machine_dict.extract_and_lookup_material(part)
                            if info and info.machine_name:
                                part_model = info.machine_name
                            else:
                                part_model = self.model_name or "raw"

                        model_slug = re.sub(r"[^\w]+", "_", part_model.strip().lower()).strip("_")
                        bak_dir_name = f"backup_before_{model_slug}_filter"
                        bak_dir = self.dest_dir / bak_dir_name

                        raw_file = None
                        if bak_dir.exists() and (bak_dir / f"PLM_{part}.xlsx").exists():
                            raw_file = bak_dir / f"PLM_{part}.xlsx"
                        else:
                            for other_bak in self.dest_dir.glob("backup_before_*_filter"):
                                cand = other_bak / f"PLM_{part}.xlsx"
                                if cand.exists():
                                    raw_file = cand
                                    break

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

                            # 1. Search item (allow up to 45s for heavy BOM queries)
                            self.log_message.emit(f"[*] [TC24] Tìm kiếm mã {part} trên hệ thống...")
                            tc_client.search_item(part, timeout=45)
                            time.sleep(2.5)

                            # 2. Content tab
                            if "page=Content" not in driver.current_url:
                                try:
                                    content_tab = wait.until(
                                        EC.element_to_be_clickable(
                                            (
                                                By.XPATH,
                                                "//a[normalize-space()='Content' or @title='Content' or @data-locator='tab-tc_xrt_Content']"
                                                " | //span[normalize-space()='Content' or @title='Content']"
                                                " | //li[contains(@class,'sw-tab')]//a[normalize-space()='Content']",
                                            )
                                        )
                                    )
                                    content_tab.click()
                                    time.sleep(3.5)
                                except Exception:
                                    tc_client.navigate_to_content_tab(timeout=15)
                            else:
                                self.log_message.emit("   [+] Tab Content đã mở sẵn.")

                            # 3. Select root row and Expand Below
                            try:
                                root_cell = wait.until(
                                    EC.element_to_be_clickable(
                                        (By.CSS_SELECTOR, "div.aw-splm-tableRow div.aw-splm-tableCellText, div.aw-splm-tableRow")
                                    )
                                )
                                root_cell.click()
                                time.sleep(1)
                            except Exception:
                                pass

                            exp_btn = wait.until(
                                EC.element_to_be_clickable(
                                    (By.CSS_SELECTOR, "button[command-id='Awb0Expand'], [command-id='Awb0Expand']")
                                )
                            )
                            exp_btn.click()
                            time.sleep(1.5)

                            cmds = driver.find_elements(
                                By.CSS_SELECTOR, "div.aw-widgets-cellListItem, [command-id='Awb0ExpandBelow'], div.sw-popup div, li[command-id='Awb0ExpandBelow']"
                            )
                            exp_below = [
                                c for c in cmds if "expand below" in (c.text or "").lower() or c.get_attribute("command-id") == "Awb0ExpandBelow"
                            ]
                            if exp_below:
                                exp_below[0].click()
                                self.log_message.emit("   [*] Đang mở rộng toàn bộ cây BOM...")
                                time.sleep(15)

                            # 4. Select all rows
                            sel_all_btn = wait.until(
                                EC.element_to_be_clickable(
                                    (By.CSS_SELECTOR, "button[command-id='Awp0SelectAll'], [command-id='Awp0SelectAll']")
                                )
                            )
                            sel_all_btn.click()
                            time.sleep(2)

                            # 5. Open Export to Excel
                            exp_imp_btn = wait.until(
                                EC.element_to_be_clickable(
                                    (By.CSS_SELECTOR, "button[command-id='Arm0ExportImport'], [command-id='Arm0ExportImport']")
                                )
                            )
                            exp_imp_btn.click()
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
                            time.sleep(4)

                            # 6. Panel configuration
                            panel = wait.until(
                                EC.presence_of_element_located(
                                    (By.CSS_SELECTOR, "form.sw-command-panel, div.sw-right-dialog form, form")
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

                            # Ensure downloaded file has .xlsx extension for openpyxl compatibility
                            if downloaded_path.suffix.lower() not in (".xlsx", ".xlsm", ".xltx", ".xltm"):
                                ext_path = downloaded_path.with_suffix(".xlsx")
                                import shutil
                                shutil.move(downloaded_path, ext_path)
                                downloaded_path = ext_path

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

                        # Apply BolocBom model filtering if requested
                        if self.apply_bolocbom and target_dest.exists() and self.tc_auto_standardize:
                            try:
                                from src.core.model_pruner import ModelPruner
                                self.log_message.emit(f"   [*] [BolocBom] Đang áp dụng bộ lọc theo dòng máy '{part_model}'...")
                                pruner = ModelPruner()
                                pruner.prune_excel_file(
                                    input_path=target_dest,
                                    output_path=target_dest,
                                    model_name=part_model,
                                )
                                import openpyxl
                                wb_p = openpyxl.load_workbook(str(target_dest), data_only=True)
                                pruned_rows = wb_p.active.max_row
                                wb_p.close()
                                self.log_message.emit(
                                    f"[+] [BolocBom] Cắt tỉa theo quy tắc '{part_model}' thành công: "
                                    f"còn lại {pruned_rows:,} dòng chuẩn."
                                )
                            except Exception as prune_err:
                                self.log_message.emit(f"   [!] [BolocBom] Cảnh báo khi áp dụng bộ lọc {part_model}: {prune_err}")

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


class BOMScanDialog(QDialog):
    """Dialog allowing users to inspect and select BOM part numbers based on their PLM and SAP R3 file status."""

    def __init__(
        self,
        parent: QWidget | None = None,
        dest_dir: Path | None = None,
        project_parts: Optional[List[str]] = None,
        config_path: Path | None = None,
    ) -> None:
        super().__init__(parent)
        self.dest_dir = Path(dest_dir or VIRGO_DIR)
        self.project_parts = [p.strip().upper() for p in (project_parts or []) if p.strip()]
        self.config_path = config_path or (
            DEFAULT_CONFIG_PATH if DEFAULT_CONFIG_PATH.exists() else FALLBACK_CONFIG_PATH
        )
        self.selected_parts: List[str] = []
        self._scanned_items: List[BOMPartStatus] = []

        self.setWindowTitle("Quét & Lựa chọn Mã BOM (Phân tích PLM & SAP R3)")
        self.resize(800, 530)
        self.setMinimumSize(720, 440)
        self._init_ui()
        self._refresh_scan()

    def _init_ui(self) -> None:
        theme_mgr = get_theme_manager()
        self.setWindowIcon(theme_mgr.get_styled_icon("search"))

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        # 1. Source selector
        source_group = QGroupBox("Nguồn dữ liệu quét mã máy / BOM")
        sg_layout = QVBoxLayout(source_group)
        sg_layout.setSpacing(6)

        self.btn_group_source = QButtonGroup(self)

        row_folder = QHBoxLayout()
        self.rb_folder = QRadioButton("Quét từ Thư mục:")
        self.rb_folder.setChecked(True)
        self.btn_group_source.addButton(self.rb_folder)
        row_folder.addWidget(self.rb_folder)

        self.edit_scan_dir = QLineEdit(str(self.dest_dir))
        self.edit_scan_dir.setFont(QFont("Consolas", 9))
        self.edit_scan_dir.setReadOnly(True)
        row_folder.addWidget(self.edit_scan_dir, stretch=1)

        self.btn_browse_scan = QPushButton(" Duyệt...")
        self.btn_browse_scan.setIcon(theme_mgr.get_styled_icon("folder"))
        self.btn_browse_scan.clicked.connect(self._browse_scan_dir)
        row_folder.addWidget(self.btn_browse_scan)
        sg_layout.addLayout(row_folder)

        # Project source radio
        row_proj = QHBoxLayout()
        proj_count = len(self.project_parts)
        proj_text = (
            f"Lấy từ Dự án hiện tại ({proj_count} mã máy đã khai báo ở Bước 1)"
            if proj_count > 0
            else "Lấy từ Dự án hiện tại (Chưa có mã nào trong dự án)"
        )
        self.rb_project = QRadioButton(proj_text)
        self.rb_project.setEnabled(proj_count > 0)
        self.btn_group_source.addButton(self.rb_project)
        row_proj.addWidget(self.rb_project)
        row_proj.addStretch()

        self.btn_refresh = QPushButton(" Quét lại")
        self.btn_refresh.setIcon(theme_mgr.get_styled_icon("refresh"))
        self.btn_refresh.clicked.connect(self._refresh_scan)
        row_proj.addWidget(self.btn_refresh)
        sg_layout.addLayout(row_proj)

        self.rb_folder.toggled.connect(self._on_source_toggled)
        layout.addWidget(source_group)

        # 2. Quick Filter Toolbar
        filter_box = QHBoxLayout()
        filter_lbl = QLabel("<b>Chọn nhanh theo trạng thái:</b>")
        filter_lbl.setStyleSheet("color: #1E293B; font-size: 11.5px;")
        filter_box.addWidget(filter_lbl)

        self.btn_sel_all = QPushButton("⚡ Tất cả mã")
        self.btn_sel_all.setStyleSheet("padding: 3px 8px; font-weight: bold;")
        self.btn_sel_all.clicked.connect(self._select_all)
        filter_box.addWidget(self.btn_sel_all)

        self.btn_sel_missing_r3 = QPushButton("⚠️ Chỉ thiếu R3")
        self.btn_sel_missing_r3.setStyleSheet(
            "background-color: #FEF3C7; color: #92400E; border: 1px solid #FDE68A; font-weight: bold; padding: 3px 8px; border-radius: 4px;"
        )
        self.btn_sel_missing_r3.clicked.connect(self._select_missing_r3)
        filter_box.addWidget(self.btn_sel_missing_r3)

        self.btn_sel_missing_plm = QPushButton("⚠️ Chỉ thiếu PLM")
        self.btn_sel_missing_plm.setStyleSheet(
            "background-color: #E0F2FE; color: #0369A1; border: 1px solid #BAE6FD; font-weight: bold; padding: 3px 8px; border-radius: 4px;"
        )
        self.btn_sel_missing_plm.clicked.connect(self._select_missing_plm)
        filter_box.addWidget(self.btn_sel_missing_plm)

        self.btn_sel_complete = QPushButton("✅ Đã đủ cả 2")
        self.btn_sel_complete.setStyleSheet(
            "background-color: #DCFCE7; color: #166534; border: 1px solid #BBF7D0; padding: 3px 8px; border-radius: 4px;"
        )
        self.btn_sel_complete.clicked.connect(self._select_complete)
        filter_box.addWidget(self.btn_sel_complete)

        self.btn_sel_none = QPushButton("Bỏ chọn hết")
        self.btn_sel_none.setStyleSheet("padding: 3px 8px;")
        self.btn_sel_none.clicked.connect(self._select_none)
        filter_box.addWidget(self.btn_sel_none)

        filter_box.addStretch()
        layout.addLayout(filter_box)

        # 3. Status Table
        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels([
            "Chọn",
            "Mã máy / BOM",
            "Trạng thái BOM PLM",
            "Trạng thái BOM SAP R3",
            "Trạng thái / Đề xuất",
        ])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        self.table.verticalHeader().setDefaultSectionSize(28)
        self.table.setShowGrid(True)
        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet("QTableWidget { gridline-color: #CBD5E1; }")
        self.table.itemChanged.connect(self._on_table_item_changed)
        layout.addWidget(self.table, stretch=1)

        # 4. Bottom bar
        bottom_layout = QHBoxLayout()

        self.chk_auto_suggest = QCheckBox("Tự động hiển thị thanh gợi ý nhanh khi mở cửa sổ tải BOM")
        self.chk_auto_suggest.setChecked(is_bom_scan_suggestion_enabled(self.config_path))
        self.chk_auto_suggest.toggled.connect(self._on_toggle_suggestion_pref)
        bottom_layout.addWidget(self.chk_auto_suggest)

        bottom_layout.addStretch()

        self.lbl_selected_summary = QLabel("Đã chọn: 0 mã")
        self.lbl_selected_summary.setStyleSheet("font-weight: bold; color: #0F172A;")
        bottom_layout.addWidget(self.lbl_selected_summary)

        self.btn_confirm = QPushButton(" Nạp vào danh sách tải")
        self.btn_confirm.setIcon(theme_mgr.get_styled_icon("download", color="#FFFFFF"))
        self.btn_confirm.setStyleSheet(
            "background-color: #059669; color: white; font-weight: bold; padding: 6px 14px; border-radius: 4px;"
        )
        self.btn_confirm.clicked.connect(self._confirm_selection)
        bottom_layout.addWidget(self.btn_confirm)

        self.btn_close = QPushButton("Đóng")
        self.btn_close.clicked.connect(self.reject)
        bottom_layout.addWidget(self.btn_close)

        layout.addLayout(bottom_layout)

    def _browse_scan_dir(self) -> None:
        folder = QFileDialog.getExistingDirectory(
            self,
            "Chọn thư mục để quét mã BOM",
            str(self.dest_dir),
        )
        if folder:
            self.dest_dir = Path(folder)
            self.edit_scan_dir.setText(str(self.dest_dir))
            self._refresh_scan()

    def _on_source_toggled(self) -> None:
        self._refresh_scan()

    def _on_toggle_suggestion_pref(self, checked: bool) -> None:
        set_bom_scan_suggestion_enabled(checked, self.config_path)

    def _refresh_scan(self) -> None:
        extra = self.project_parts if self.rb_project.isChecked() else None
        all_items = scan_bom_directory(self.dest_dir, extra_parts=extra)

        if self.rb_project.isChecked() and self.project_parts:
            # Filter strictly to project parts
            proj_set = set(self.project_parts)
            self._scanned_items = [itm for itm in all_items if itm.part_number in proj_set]
        else:
            self._scanned_items = all_items

        self.table.blockSignals(True)
        self.table.setRowCount(len(self._scanned_items))

        count_missing_r3 = 0
        count_missing_plm = 0
        count_complete = 0

        for r, itm in enumerate(self._scanned_items):
            if itm.needs_r3:
                count_missing_r3 += 1
            elif itm.needs_plm:
                count_missing_plm += 1
            elif itm.is_complete:
                count_complete += 1

            # Col 0: Checkbox
            chk_item = QTableWidgetItem()
            chk_item.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
            # Default check items that have missing parts, or all if empty
            if itm.needs_r3 or itm.needs_plm:
                chk_item.setCheckState(Qt.CheckState.Checked)
            else:
                chk_item.setCheckState(Qt.CheckState.Unchecked)
            self.table.setItem(r, 0, chk_item)

            # Col 1: Part number
            p_item = QTableWidgetItem(itm.part_number)
            p_item.setFont(QFont("Consolas", 10, QFont.Weight.Bold))
            p_item.setFlags(p_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(r, 1, p_item)

            # Col 2: PLM status
            if itm.has_plm:
                sz_str = f"{itm.plm_size // 1024:,} KB" if itm.plm_size > 0 else "Có file"
                plm_text = f"✓ Đã có ({sz_str})"
                plm_color = QColor("#059669")
            else:
                plm_text = "✗ Chưa có"
                plm_color = QColor("#DC2626")
            plm_item = QTableWidgetItem(plm_text)
            plm_item.setForeground(plm_color)
            plm_item.setFlags(plm_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(r, 2, plm_item)

            # Col 3: SAP R3 status
            if itm.has_r3:
                sz_str = f"{itm.r3_size // 1024:,} KB" if itm.r3_size > 0 else "Có file"
                r3_text = f"✓ Đã có ({sz_str})"
                r3_color = QColor("#059669")
            else:
                r3_text = "✗ Chưa có"
                r3_color = QColor("#DC2626")
            r3_item = QTableWidgetItem(r3_text)
            r3_item.setForeground(r3_color)
            r3_item.setFlags(r3_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(r, 3, r3_item)

            # Col 4: Summary / Recommendation
            rec_item = QTableWidgetItem(itm.status_summary)
            if itm.is_complete:
                rec_item.setForeground(QColor("#059669"))
            elif itm.needs_r3:
                rec_item.setForeground(QColor("#D97706"))
            elif itm.needs_plm:
                rec_item.setForeground(QColor("#0284C7"))
            else:
                rec_item.setForeground(QColor("#64748B"))
            rec_item.setFlags(rec_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(r, 4, rec_item)

        # If nothing was checked by default (e.g. all complete or all empty), check all
        checked_init = sum(
            1 for r in range(self.table.rowCount())
            if self.table.item(r, 0) and self.table.item(r, 0).checkState() == Qt.CheckState.Checked
        )
        if checked_init == 0 and self.table.rowCount() > 0:
            for r in range(self.table.rowCount()):
                it = self.table.item(r, 0)
                if it:
                    it.setCheckState(Qt.CheckState.Checked)

        self.table.blockSignals(False)

        # Update filter button texts
        self.btn_sel_all.setText(f"⚡ Tất cả ({len(self._scanned_items)})")
        self.btn_sel_missing_r3.setText(f"⚠️ Chỉ thiếu R3 ({count_missing_r3})")
        self.btn_sel_missing_plm.setText(f"⚠️ Chỉ thiếu PLM ({count_missing_plm})")
        self.btn_sel_complete.setText(f"✅ Đã đủ cả 2 ({count_complete})")

        self._update_selected_count()

    def _update_selected_count(self) -> None:
        total = self.table.rowCount()
        checked = sum(
            1 for r in range(total)
            if self.table.item(r, 0) and self.table.item(r, 0).checkState() == Qt.CheckState.Checked
        )
        self.lbl_selected_summary.setText(f"Đã chọn: {checked} / {total} mã")

    def _on_table_item_changed(self, item: QTableWidgetItem) -> None:
        if item.column() == 0:
            self._update_selected_count()

    def _select_all(self) -> None:
        self.table.blockSignals(True)
        for r in range(self.table.rowCount()):
            it = self.table.item(r, 0)
            if it:
                it.setCheckState(Qt.CheckState.Checked)
        self.table.blockSignals(False)
        self._update_selected_count()

    def _select_none(self) -> None:
        self.table.blockSignals(True)
        for r in range(self.table.rowCount()):
            it = self.table.item(r, 0)
            if it:
                it.setCheckState(Qt.CheckState.Unchecked)
        self.table.blockSignals(False)
        self._update_selected_count()

    def _select_missing_r3(self) -> None:
        self.table.blockSignals(True)
        for r in range(self.table.rowCount()):
            it = self.table.item(r, 0)
            if it and r < len(self._scanned_items):
                is_target = self._scanned_items[r].needs_r3
                it.setCheckState(Qt.CheckState.Checked if is_target else Qt.CheckState.Unchecked)
        self.table.blockSignals(False)
        self._update_selected_count()

    def _select_missing_plm(self) -> None:
        self.table.blockSignals(True)
        for r in range(self.table.rowCount()):
            it = self.table.item(r, 0)
            if it and r < len(self._scanned_items):
                is_target = self._scanned_items[r].needs_plm
                it.setCheckState(Qt.CheckState.Checked if is_target else Qt.CheckState.Unchecked)
        self.table.blockSignals(False)
        self._update_selected_count()

    def _select_complete(self) -> None:
        self.table.blockSignals(True)
        for r in range(self.table.rowCount()):
            it = self.table.item(r, 0)
            if it and r < len(self._scanned_items):
                is_target = self._scanned_items[r].is_complete
                it.setCheckState(Qt.CheckState.Checked if is_target else Qt.CheckState.Unchecked)
        self.table.blockSignals(False)
        self._update_selected_count()

    def _confirm_selection(self) -> None:
        chosen = []
        for r in range(self.table.rowCount()):
            it = self.table.item(r, 0)
            p_it = self.table.item(r, 1)
            if it and it.checkState() == Qt.CheckState.Checked and p_it:
                chosen.append(p_it.text().strip())

        if not chosen:
            QMessageBox.warning(self, "Chưa chọn mã", "Vui lòng tích chọn ít nhất một mã BOM để nạp!")
            return

        self.selected_parts = chosen
        self.accept()

    def get_selected_parts(self) -> List[str]:
        return self.selected_parts


class NoScrollComboBox(QComboBox):
    """QComboBox that ignores mouse wheel scrolling when closed.

    Prevents accidental item switching when scrolling through the dialog with the mouse wheel.
    The user must explicitly click to open the dropdown and select an item.
    """

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)

    def wheelEvent(self, e):
        e.ignore()


class PLMDownloadDialog(QDialog):
    """User-friendly, Non-Tech BOM Downloader for Siemens TC24 and SAP R3."""

    def __init__(
        self,
        parent: QWidget | None = None,
        default_dir: Path | None = None,
        initial_model: str | None = None,
        config_path: Path | None = None,
        project_parts: Optional[List[str]] = None,
    ) -> None:
        super().__init__(parent)
        self._config_path = config_path or (
            DEFAULT_CONFIG_PATH if DEFAULT_CONFIG_PATH.exists() else FALLBACK_CONFIG_PATH
        )
        self._machine_dict = MachineDictService()
        self._initial_model = initial_model
        self._default_base_dir = default_dir.parent if default_dir else None
        self._project_parts = [p.strip().upper() for p in (project_parts or []) if p.strip()]
        self._suggestion_dismissed = False
        self._cached_scanned_parts: List[BOMPartStatus] = []

        fallback_dir = Path(default_dir or VIRGO_DIR)
        self.dest_dir = get_persisted_download_dir(
            model_name=initial_model,
            fallback_dir=fallback_dir,
            config_path=self._config_path,
        )
        self.thread: QThread | None = None
        self.worker: UnifiedBOMDownloadWorker | None = None
        self.current_items: List[BOMDownloadItem] = []
        self._detected_model_name: str | None = None
        self._detected_models_map: Dict[str, str] = {}

        self.setWindowTitle("Tải Tự Động BOM Đa Nguồn (Siemens TC24 & SAP R3)")
        self.resize(860, 640)
        self.setMinimumSize(780, 500)
        self._init_ui()
        self._update_smart_suggestion()

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

        self.combo_model = NoScrollComboBox()
        self.combo_model.setFont(QFont("Calibri", 10))
        self.combo_model.addItem("-- Tự động nhận diện theo mã BOM --")
        try:
            for m_name in self._machine_dict.get_model_names():
                self.combo_model.addItem(m_name)
        except Exception as exc:
            logger.warning("Could not load model names: %s", exc)

        if self._initial_model:
            idx = self.combo_model.findText(self._initial_model)
            if idx < 0:
                self.combo_model.addItem(self._initial_model)

        self.combo_model.setCurrentIndex(0)

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

        self.btn_scan = QPushButton(" Quét & Nạp mã...")
        self.btn_scan.setIcon(theme_mgr.get_styled_icon("search"))
        self.btn_scan.setStyleSheet("padding: 4px 10px; font-weight: bold;")
        self.btn_scan.clicked.connect(self._open_scan_dialog)
        tools_layout.addWidget(self.btn_scan)

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

        # Smart Suggestion Banner
        self.banner_frame = QFrame()
        self.banner_frame.setStyleSheet("""
            QFrame {
                background-color: #F0FDF4;
                border: 1px solid #BBF7D0;
                border-radius: 6px;
            }
        """)
        banner_layout = QHBoxLayout(self.banner_frame)
        banner_layout.setContentsMargins(10, 6, 10, 6)
        banner_layout.setSpacing(6)

        self.lbl_banner_icon = QLabel("💡")
        self.lbl_banner_msg = QLabel("")
        self.lbl_banner_msg.setFont(QFont("Calibri", 10))
        self.lbl_banner_msg.setStyleSheet("color: #166534;")
        banner_layout.addWidget(self.lbl_banner_icon)
        banner_layout.addWidget(self.lbl_banner_msg, stretch=1)

        self.btn_banner_missing_r3 = QPushButton("⚡ Nạp mã thiếu R3")
        self.btn_banner_missing_r3.setStyleSheet(
            "background-color: #D97706; color: white; font-weight: bold; padding: 4px 10px; border-radius: 4px;"
        )
        self.btn_banner_missing_r3.clicked.connect(self._on_banner_load_missing_r3)
        banner_layout.addWidget(self.btn_banner_missing_r3)

        self.btn_banner_missing_plm = QPushButton("⚡ Nạp mã thiếu PLM")
        self.btn_banner_missing_plm.setStyleSheet(
            "background-color: #0284C7; color: white; font-weight: bold; padding: 4px 10px; border-radius: 4px;"
        )
        self.btn_banner_missing_plm.clicked.connect(self._on_banner_load_missing_plm)
        banner_layout.addWidget(self.btn_banner_missing_plm)

        self.btn_banner_all = QPushButton("⚡ Nạp tất cả")
        self.btn_banner_all.setStyleSheet(
            "background-color: #2563EB; color: white; font-weight: bold; padding: 4px 10px; border-radius: 4px;"
        )
        self.btn_banner_all.clicked.connect(self._on_banner_load_all)
        banner_layout.addWidget(self.btn_banner_all)

        self.btn_banner_scan = QPushButton("📂 Chi tiết...")
        self.btn_banner_scan.setObjectName("btn_banner_scan")
        self.btn_banner_scan.setStyleSheet("""
            QPushButton#btn_banner_scan {
                background-color: #FFFFFF;
                color: #1E293B;
                border: 1px solid #CBD5E1;
                font-weight: 600;
                padding: 4px 10px;
                border-radius: 4px;
                font-size: 11.5px;
            }
            QPushButton#btn_banner_scan:hover {
                background-color: #F1F5F9;
                border-color: #94A3B8;
            }
            QPushButton#btn_banner_scan:pressed {
                background-color: #E2E8F0;
            }
        """)
        self.btn_banner_scan.clicked.connect(self._open_scan_dialog)
        banner_layout.addWidget(self.btn_banner_scan)

        self.btn_banner_dismiss = QPushButton()
        self.btn_banner_dismiss.setObjectName("btn_banner_dismiss")
        self.btn_banner_dismiss.setIcon(theme_mgr.get_styled_icon("x-circle", color="#475569"))
        self.btn_banner_dismiss.setIconSize(QSize(16, 16))
        self.btn_banner_dismiss.setFixedSize(26, 26)
        self.btn_banner_dismiss.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_banner_dismiss.setToolTip("Đóng thanh gợi ý này")
        self.btn_banner_dismiss.setStyleSheet("""
            QPushButton#btn_banner_dismiss {
                background-color: transparent;
                border: 1px solid transparent;
                border-radius: 4px;
                padding: 0px;
                margin: 0px;
                min-width: 26px;
                max-width: 26px;
                min-height: 26px;
                max-height: 26px;
            }
            QPushButton#btn_banner_dismiss:hover {
                background-color: #FEE2E2;
                border: 1px solid #FCA5A5;
            }
            QPushButton#btn_banner_dismiss:pressed {
                background-color: #FECACA;
                border: 1px solid #F87171;
            }
        """)
        self.btn_banner_dismiss.clicked.connect(self._dismiss_smart_suggestion)
        banner_layout.addWidget(self.btn_banner_dismiss)

        self.banner_frame.setVisible(False)
        part_layout.addWidget(self.banner_frame)

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

        self.combo_account = NoScrollComboBox()
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

        self.chk_apply_bolocbom = QCheckBox(
            "Tự động áp dụng Bộ lọc BOM (BolocBom) theo từng loại máy ngay sau khi tải"
        )
        self.chk_apply_bolocbom.setChecked(True)
        self.chk_apply_bolocbom.setToolTip(
            "Sau khi tải và chuẩn hóa 14 cột, tự động nhận diện dòng máy và chạy cắt tỉa "
            "các cụm con/phantom branch theo đúng quy tắc Sheet BolocBom."
        )
        tc_form.addRow("Bộ lọc BolocBom:", self.chk_apply_bolocbom)
        self.chk_standardize.toggled.connect(self._on_standardize_toggled)
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
        self.edit_dest.editingFinished.connect(self._on_dest_editing_finished)
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
            return self._detected_model_name or self._initial_model or "Virgo"
        return txt

    def _update_backup_preview(self) -> None:
        if not hasattr(self, "lbl_backup_preview"):
            return

        # If on Auto mode and multiple distinct models detected, show breakdown
        if (
            self.combo_model.currentIndex() == 0
            and hasattr(self, "_detected_models_map")
            and self._detected_models_map
        ):
            distinct_models = list(dict.fromkeys(self._detected_models_map.values()))
            if len(distinct_models) > 1:
                slug_tags = [
                    f"backup_before_{re.sub(r'[^\w]+', '_', m.strip().lower()).strip('_')}_filter/"
                    for m in distinct_models
                ]
                preview_list = ", ".join(slug_tags)
                self.lbl_backup_preview.setText(
                    f"Thư mục sao lưu file gốc (trước khi lọc): "
                    f"<b style='color: #0284c7;'>Tự động tách riêng theo từng dòng máy ({preview_list})</b>"
                )
                return

        model_name = self.get_current_model_name()
        model_slug = re.sub(r"[^\w]+", "_", model_name.strip().lower()).strip("_")
        bak_name = f"backup_before_{model_slug}_filter"
        self.lbl_backup_preview.setText(
            f"Thư mục sao lưu file gốc (trước khi lọc): <b style='color: #0284c7;'>{bak_name}/</b>"
        )

    def _on_model_selection_changed(self, index: int) -> None:
        self._update_backup_preview()
        model = self.get_current_model_name()
        if self._default_base_dir:
            model_fallback = self._default_base_dir / model
        elif SO_SANH_BASE_DIR.exists():
            model_fallback = SO_SANH_BASE_DIR / model
        else:
            model_fallback = Path(model)

        persisted = get_persisted_download_dir(
            model_name=model,
            fallback_dir=model_fallback,
            config_path=self._config_path,
        )
        self.dest_dir = persisted
        self.edit_dest.setText(str(self.dest_dir))
        self._update_smart_suggestion()

    def _open_scan_dialog(self) -> None:
        dlg = BOMScanDialog(
            parent=self,
            dest_dir=self.dest_dir,
            project_parts=self._project_parts,
            config_path=self._config_path,
        )
        if dlg.exec():
            chosen = dlg.get_selected_parts()
            if chosen:
                self.txt_parts.setPlainText("\n".join(chosen))

    def _update_smart_suggestion(self) -> None:
        if not hasattr(self, "banner_frame"):
            return

        if self._suggestion_dismissed:
            self.banner_frame.setVisible(False)
            return

        if not is_bom_scan_suggestion_enabled(self._config_path):
            self.banner_frame.setVisible(False)
            return

        if self.txt_parts.toPlainText().strip():
            self.banner_frame.setVisible(False)
            return

        scanned = scan_bom_directory(self.dest_dir, extra_parts=self._project_parts)
        if not scanned:
            self.banner_frame.setVisible(False)
            return

        self._cached_scanned_parts = scanned
        total = len(scanned)
        missing_r3 = [p for p in scanned if p.needs_r3]
        missing_plm = [p for p in scanned if p.needs_plm]
        model_name = self.get_current_model_name()

        if missing_r3 and not missing_plm:
            msg = f"Tìm thấy <b>{total}</b> mã trong {model_name} (Có <b>{len(missing_r3)}</b> mã đã có PLM nhưng <b>chưa có BOM R3</b>):"
        elif missing_plm and not missing_r3:
            msg = f"Tìm thấy <b>{total}</b> mã trong {model_name} (Có <b>{len(missing_plm)}</b> mã đã có R3 nhưng <b>chưa có BOM PLM</b>):"
        elif missing_r3 and missing_plm:
            msg = f"Tìm thấy <b>{total}</b> mã ({len(missing_r3)} thiếu R3, {len(missing_plm)} thiếu PLM):"
        else:
            msg = f"Tìm thấy <b>{total}</b> mã BOM trong thư mục {model_name}:"

        self.lbl_banner_msg.setText(msg)
        self.btn_banner_missing_r3.setVisible(len(missing_r3) > 0)
        if len(missing_r3) > 0:
            self.btn_banner_missing_r3.setText(f"⚡ Nạp {len(missing_r3)} mã thiếu R3")

        self.btn_banner_missing_plm.setVisible(len(missing_plm) > 0)
        if len(missing_plm) > 0:
            self.btn_banner_missing_plm.setText(f"⚡ Nạp {len(missing_plm)} mã thiếu PLM")

        self.btn_banner_all.setText(f"⚡ Nạp tất cả ({total})")
        self.banner_frame.setVisible(True)

    def _on_banner_load_missing_r3(self) -> None:
        missing_r3 = [p.part_number for p in self._cached_scanned_parts if p.needs_r3]
        if missing_r3:
            self.txt_parts.setPlainText("\n".join(missing_r3))

    def _on_banner_load_missing_plm(self) -> None:
        missing_plm = [p.part_number for p in self._cached_scanned_parts if p.needs_plm]
        if missing_plm:
            self.txt_parts.setPlainText("\n".join(missing_plm))

    def _on_banner_load_all(self) -> None:
        all_parts = [p.part_number for p in self._cached_scanned_parts]
        if all_parts:
            self.txt_parts.setPlainText("\n".join(all_parts))

    def _dismiss_smart_suggestion(self) -> None:
        self._suggestion_dismissed = True
        self.banner_frame.setVisible(False)

    def _on_text_changed(self) -> None:
        """Update counter, auto-detect model from machine dictionary, and sync date table."""
        self.current_items = parse_bom_items(self.txt_parts.toPlainText())
        count = len(self.current_items)

        if count == 0:
            self.lbl_count.setText("Đã nhận diện: <b>0</b> mã BOM")
            self.lbl_count.setStyleSheet("color: #64748b; font-size: 12px;")
            self.lbl_model_badge.setVisible(False)
            self._detected_model_name = None
            self._detected_models_map = {}
            self._update_smart_suggestion()
        else:
            if hasattr(self, "banner_frame"):
                self.banner_frame.setVisible(False)
            self.lbl_count.setText(f"Đã nhận diện: <b>{count}</b> mã BOM hợp lệ")
            self.lbl_count.setStyleSheet("color: #16a34a; font-size: 12px; font-weight: bold;")

            # Auto-detect machine model using MachineDictService across all items
            detected_models_map: Dict[str, str] = {}
            model_counts: Dict[str, int] = {}
            detected_infos: Dict[str, Any] = {}

            for item in self.current_items:
                info = self._machine_dict.extract_and_lookup_material(item.part_number)
                if info and info.machine_name:
                    m_name = info.machine_name
                    detected_models_map[item.part_number] = m_name
                    model_counts[m_name] = model_counts.get(m_name, 0) + 1
                    if m_name not in detected_infos:
                        detected_infos[m_name] = info

            self._detected_models_map = detected_models_map
            distinct_models = list(model_counts.keys())

            if len(distinct_models) == 1:
                model_name = distinct_models[0]
                self._detected_model_name = model_name
                info = detected_infos[model_name]
                matched_codes = ", ".join(info.machine_codes) if info.machine_codes else ""
                badge_text = f"✓ Nhận diện từ file_loaimay: <b>{model_name}</b>"
                if count > 1:
                    badge_text += f" (Tất cả {count} mã)"
                if matched_codes:
                    badge_text += f" (Mã: {matched_codes})"

                self.lbl_model_badge.setText(badge_text)
                self.lbl_model_badge.setStyleSheet(
                    "color: #065f46; background-color: #d1fae5; padding: 3px 8px; "
                    "border-radius: 4px; border: 1px solid #a7f3d0; font-size: 11px;"
                )
                self.lbl_model_badge.setVisible(True)

                # If combo is on Auto, update backup preview and destination according to detected model
                if self.combo_model.currentIndex() == 0:
                    self._on_model_selection_changed(0)

            elif len(distinct_models) > 1:
                self._detected_model_name = distinct_models[0]
                parts_breakdown = ", ".join([f"<b>{m}</b> ({c} mã)" for m, c in model_counts.items()])
                badge_text = f"✓ Nhận diện đa dòng máy: {parts_breakdown}"
                self.lbl_model_badge.setText(badge_text)
                self.lbl_model_badge.setStyleSheet(
                    "color: #1e40af; background-color: #dbeafe; padding: 3px 8px; "
                    "border-radius: 4px; border: 1px solid #bfdbfe; font-size: 11px;"
                )
                self.lbl_model_badge.setVisible(True)

                if self.combo_model.currentIndex() == 0:
                    self._update_backup_preview()
            else:
                self._detected_model_name = None
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

    def _on_standardize_toggled(self, checked: bool) -> None:
        """Enable or disable BolocBom pruning option when standardization is toggled."""
        if hasattr(self, "chk_apply_bolocbom"):
            self.chk_apply_bolocbom.setEnabled(checked)
            if not checked:
                self.chk_apply_bolocbom.setChecked(False)
            else:
                self.chk_apply_bolocbom.setChecked(True)

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
        self._detected_model_name = None
        self._detected_models_map = {}
        self.combo_model.setCurrentIndex(0)

    def _browse_dest_dir(self) -> None:
        folder = QFileDialog.getExistingDirectory(
            self,
            "Chọn thư mục lưu BOM",
            self.edit_dest.text().strip() or str(self.dest_dir),
        )
        if folder:
            self.edit_dest.setText(folder)
            self.dest_dir = Path(folder)
            save_persisted_download_dir(
                self.get_current_model_name(),
                self.dest_dir,
                config_path=self._config_path,
            )

    def _on_dest_editing_finished(self) -> None:
        txt = self.edit_dest.text().strip()
        if txt:
            p = Path(txt)
            self.dest_dir = p
            try:
                if p.exists() and p.is_dir():
                    save_persisted_download_dir(
                        self.get_current_model_name(),
                        p,
                        config_path=self._config_path,
                    )
            except Exception:
                pass

    def _set_ui_busy(self, busy: bool) -> None:
        self.btn_both.setEnabled(not busy)
        self.btn_plm_only.setEnabled(not busy)
        self.btn_sap_only.setEnabled(not busy)
        self.btn_browse.setEnabled(not busy)
        self.btn_paste.setEnabled(not busy)
        self.btn_clear.setEnabled(not busy)
        self.chk_custom_dates.setEnabled(not busy)
        if hasattr(self, "chk_standardize"):
            self.chk_standardize.setEnabled(not busy)
        if hasattr(self, "chk_apply_bolocbom"):
            self.chk_apply_bolocbom.setEnabled(not busy and self.chk_standardize.isChecked())

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
            save_persisted_download_dir(
                self.get_current_model_name(),
                self.dest_dir,
                config_path=self._config_path,
            )
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

        # Collect part -> model mapping
        if self.combo_model.currentIndex() > 0:
            selected_model = self.combo_model.currentText().strip()
            part_model_map = {item.part_number: selected_model for item in self.current_items}
        else:
            part_model_map = dict(getattr(self, "_detected_models_map", {}))
            for item in self.current_items:
                if item.part_number not in part_model_map:
                    info = self._machine_dict.extract_and_lookup_material(item.part_number)
                    if info and info.machine_name:
                        part_model_map[item.part_number] = info.machine_name

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
            part_model_map=part_model_map,
            apply_bolocbom=self.chk_apply_bolocbom.isChecked() if hasattr(self, "chk_apply_bolocbom") else True,
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
