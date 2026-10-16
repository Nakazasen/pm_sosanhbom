"""
SAP R3 CS12 BOM Multilevel Execution & Automated Spreadsheet Export.
Handles /nCS12 navigation, parameter binding, status bar fail-closed error detection,
SAPLSPO5:0150 dialog interaction, and file routing.
"""

import logging
import os
import shutil
import time
from datetime import date, datetime
from pathlib import Path
from typing import Any, Callable, List, Optional, Union

from src.automation.sap.models import (
    CS12Params,
    ExportResult,
    SAPCS12Error,
)

logger = logging.getLogger(__name__)

RPC_HRESULTS_U32 = {
    0x800706BE,  # RPC_S_CALL_FAILED (-2147023170)
    0x800706BA,  # RPC_S_SERVER_UNAVAILABLE (-2147023174)
    0x80010108,  # RPC_E_DISCONNECTED (-2147417848)
    0x800706BF,  # RPC_S_CALL_FAILED_DNE (-2147023169)
    0x80010105,  # RPC_E_SERVERFAULT (-2147417851)
    0x80010001,  # RPC_E_CALL_REJECTED (-2147418111)
    0x80080006,  # CO_E_SERVER_STOPPING (-2146959354)
    0x8001011F,  # RPC_E_SERVER_DIED (-2147417825)
    0x80010100,  # RPC_E_SYS_CALL_FAILED (-2147418112)
    0x8001010E,  # RPC_E_WRONG_THREAD (-2147417842)
    0x800706D9,  # EPT_S_NOT_REGISTERED (-2147023143)
    0x800706B5,  # RPC_S_UNKNOWN_IF (-2147023179)
}

# Bidirectional signed and unsigned representation set
RPC_HRESULTS = {
    code if code < 0x80000000 else code - 0x100000000
    for code in RPC_HRESULTS_U32
} | RPC_HRESULTS_U32

RPC_ERROR_KEYWORDS = (
    "remote procedure call failed",
    "rpc server is unavailable",
    "the rpc server is unavailable",
    "rpc_s_call_failed",
    "rpc_s_server_unavailable",
    "the object invoked has disconnected",
    "disconnected from its clients",
    "-2147023170",
    "-2147023174",
    "-2147417848",
    "0x800706be",
    "0x800706ba",
    "0x80010108",
    "server threw an exception",
    "call was rejected by callee",
    "session closed",
    "session is closed",
    "session disconnected",
    "session was disconnected",
    "session terminated",
    "session was terminated",
    "connection to partner",
    "connection to partner broken",
    "connection was terminated",
    "connection terminated",
    "connection closed",
    "connection lost",
    "connection broken",
    "connection reset",
    "network connection lost",
    "wsaeconnreset",
    "server died",
    "not connected to sap",
)


class CS12Service:
    """Automates SAP Transaction CS12 (Multilevel BOM Display & Export)."""

    def __init__(
        self,
        session: Any = None,
        connection_manager: Optional[Any] = None,
        status_callback: Optional[Callable[[str], None]] = None,
        on_session_refreshed: Optional[Callable[[Any], None]] = None,
    ):
        """
        Initialize CS12Service.

        Args:
            session: Active, authenticated SAP GUI session object.
            connection_manager: Optional SAPConnectionManager providing active session and reconnect capability.
            status_callback: Optional callable for streaming human-readable log/status messages to caller.
            on_session_refreshed: Optional callable invoked whenever a new session is established.
        """
        self.connection_manager = connection_manager
        self.status_callback = status_callback
        self.on_session_refreshed = on_session_refreshed

        if session is not None:
            self.session = session
        elif connection_manager is not None:
            if hasattr(connection_manager, "get_session"):
                self.session = connection_manager.get_session()
            elif hasattr(connection_manager, "session"):
                self.session = connection_manager.session
            else:
                self.session = connection_manager
        else:
            self.session = None

    @classmethod
    def _is_rpc_or_session_error(cls, ex: Optional[BaseException]) -> bool:
        """Detect whether an exception is caused by COM RPC failure or session disconnection."""
        if ex is None:
            return False

        curr: Optional[BaseException] = ex
        visited = set()
        while curr is not None and id(curr) not in visited:
            visited.add(id(curr))

            # 1. Check hresult / winerror attributes
            for attr in ("hresult", "winerror"):
                val = getattr(curr, attr, None)
                if isinstance(val, int) and ((val & 0xFFFFFFFF) in RPC_HRESULTS_U32 or val in RPC_HRESULTS):
                    return True

            # 2. Check args tuple (common in pywintypes.com_error: (hr, msg, details, ...))
            if hasattr(curr, "args") and curr.args:
                for arg in curr.args:
                    if isinstance(arg, int) and ((arg & 0xFFFFFFFF) in RPC_HRESULTS_U32 or arg in RPC_HRESULTS):
                        return True

            # 3. Check message text
            msg = str(curr).lower()
            if any(kw in msg for kw in RPC_ERROR_KEYWORDS):
                return True

            curr = getattr(curr, "__cause__", None) or getattr(curr, "__context__", None)

        return False

    def is_session_alive(self) -> bool:
        """Check if current SAP session is responsive and authenticated."""
        if self.session is None:
            return False
        try:
            okcd = self.session.findById("wnd[0]/tbar[0]/okcd")
            return okcd is not None
        except Exception:
            return False

    def reacquire_session(self) -> bool:
        """Re-acquire a fresh SAP session via connection_manager if available."""
        if self.connection_manager is None:
            return False
        try:
            logger.info("Re-acquiring fresh SAP session via connection_manager...")
            new_session = None
            if hasattr(self.connection_manager, "reconnect"):
                new_session = self.connection_manager.reconnect()
            elif hasattr(self.connection_manager, "get_session"):
                if hasattr(self.connection_manager, "disconnect"):
                    try:
                        self.connection_manager.disconnect()
                    except Exception:
                        pass
                if hasattr(self.connection_manager, "session"):
                    self.connection_manager.session = None
                new_session = self.connection_manager.get_session()
            elif hasattr(self.connection_manager, "get_or_create_session"):
                if hasattr(self.connection_manager, "disconnect"):
                    try:
                        self.connection_manager.disconnect()
                    except Exception:
                        pass
                new_session = self.connection_manager.get_or_create_session()

            if new_session is not None:
                self.session = new_session
                if self.on_session_refreshed is not None:
                    try:
                        self.on_session_refreshed(self.session)
                    except Exception as cb_err:
                        logger.debug("on_session_refreshed callback error: %s", cb_err)
                return True
            self.session = None
            return False
        except Exception as ex:
            logger.error("Failed to re-acquire SAP session via connection_manager: %s", ex)
            self.session = None
            return False

    def ensure_active_session(self) -> Any:
        """Ensure session is alive; reconnects via connection_manager if unresponsive."""
        if not self.is_session_alive():
            if self.connection_manager is not None:
                logger.warning("Phiên SAP hiện tại không phản hồi hoặc đã mất kết nối. Đang tự động kết nối lại...")
                self.reacquire_session()
        return self.session

    def execute_cs12(
        self,
        material: str,
        plant: str = "2200",
        bom_usage: str = "pp01",
        alternative: str = "01",
        valid_date: Optional[Union[date, str]] = None,
        destination_dir: Optional[Union[str, Path]] = None,
    ) -> ExportResult:
        """Execute CS12 query for material BOM with direct parameters."""
        params = CS12Params(
            material=material,
            plant=plant,
            bom_usage=bom_usage,
            alternative=alternative,
            valid_date=valid_date,
            destination_dir=destination_dir,
        )
        try:
            return self.execute_cs12_and_export(params)
        except SAPCS12Error:
            raise
        except Exception as ex:
            raise SAPCS12Error(f"CS12 execution error: {ex}") from ex

    def wait_ready(self, timeout_sec: float = 30.0) -> None:
        """Wait until session is not busy."""
        start_time = time.time()
        while getattr(self.session, "Busy", False) is True:
            time.sleep(0.1)
            if time.time() - start_time > timeout_sec:
                raise SAPCS12Error(f"SAP GUI session busy timeout after {timeout_sec}s.")

    def navigate_to_cs12(self) -> None:
        """Safely navigate to /nCS12 from any screen state."""
        logger.debug("Navigating to /nCS12...")
        try:
            okcd = self.session.findById("wnd[0]/tbar[0]/okcd")
            okcd.Text = "/nCS12"
            wnd0 = self.session.findById("wnd[0]")
            wnd0.sendVKey(0)  # Enter
            self.wait_ready()
        except Exception as ex:
            raise SAPCS12Error(f"Failed to navigate to /nCS12: {ex}") from ex

    def reset_navigation(self) -> None:
        """Reset SAP transaction screen to top-level menu via /n."""
        logger.debug("Resetting navigation via /n...")
        try:
            if self.session is None:
                return
            okcd = self.session.findById("wnd[0]/tbar[0]/okcd")
            okcd.Text = "/n"
            wnd0 = self.session.findById("wnd[0]")
            wnd0.sendVKey(0)
            self.wait_ready()
        except Exception as ex:
            logger.debug(f"Reset navigation encountered non-fatal error: {ex}")

    def logoff(self) -> bool:
        """Exit and logoff from SAP via /nex fast logoff."""
        logger.info("Logging off SAP via /nex...")
        try:
            if self.session is None:
                return False
            okcd = self.session.findById("wnd[0]/tbar[0]/okcd")
            okcd.Text = "/nex"
            wnd0 = self.session.findById("wnd[0]")
            wnd0.sendVKey(0)
            return True
        except Exception as ex:
            logger.debug(f"Failed logging off SAP via /nex: {ex}")
            return False

    def execute_cs12_and_export(
        self,
        params: CS12Params,
        allow_retry: bool = True,
    ) -> ExportResult:
        """
        Execute CS12 query for a single material BOM and export to spreadsheet.

        Args:
            params: Parameters including material, plant, bom_usage, alternative, valid_date.
            allow_retry: Whether to automatically reconnect and retry once on COM RPC failure.

        Returns:
            ExportResult detailing success status, output path, or status bar error.
        """
        start_time = time.time()
        material = params.material
        destination_dir = params.destination_dir or Path.cwd()
        destination_dir = Path(destination_dir)
        destination_dir.mkdir(parents=True, exist_ok=True)

        expected_file = destination_dir / params.expected_filename
        clean_dir_str = str(destination_dir.resolve())
        if not clean_dir_str.endswith("\\"):
            clean_dir_str += "\\"

        logger.info(
            f"Executing CS12 for Material='{material}', Plant='{params.plant}', "
            f"Usage='{params.bom_usage}', Alt='{params.alternative}', Date='{params.formatted_date_sap}'"
        )

        if self.session is None:
            if allow_retry and self.connection_manager is not None:
                self.reacquire_session()
            if self.session is None:
                return ExportResult(
                    success=False,
                    material=material,
                    error_message="No active SAP GUI session available.",
                    status_code="E",
                    execution_time_sec=time.time() - start_time,
                )

        try:
            # 1. Navigation to initial CS12 screen
            self.navigate_to_cs12()

            # 2. Parameter binding on wnd[0]/usr
            matnr_field = self.session.findById("wnd[0]/usr/ctxtRC29L-MATNR")
            matnr_field.Text = material

            werks_field = self.session.findById("wnd[0]/usr/ctxtRC29L-WERKS")
            werks_field.Text = params.plant

            stlal_field = self.session.findById("wnd[0]/usr/txtRC29L-STLAL")
            stlal_field.Text = params.alternative

            capid_field = self.session.findById("wnd[0]/usr/ctxtRC29L-CAPID")
            capid_field.Text = params.bom_usage

            datuv_field = self.session.findById("wnd[0]/usr/ctxtRC29L-DATUV")
            datuv_field.Text = params.formatted_date_sap

            # 3. Execute query (btn[8] / F8)
            btn_exec = self.session.findById("wnd[0]/tbar[1]/btn[8]")
            btn_exec.press()
            self.wait_ready()

            # 4. Status Bar Fail-Closed Inspection
            sbar = self.session.findById("wnd[0]/sbar")
            msg_type = getattr(sbar, "MessageType", "")
            msg_text = getattr(sbar, "Text", "")

            if msg_type in ("E", "A"):
                logger.error(f"SAP CS12 fail-closed gate triggered for '{material}': [{msg_type}] {msg_text}")
                self.reset_navigation()
                return ExportResult(
                    success=False,
                    material=material,
                    error_message=f"SAP Error [{msg_type}]: {msg_text}",
                    status_code=msg_type,
                    execution_time_sec=time.time() - start_time,
                )

            # 5. Trigger Spreadsheet Export (btn[45])
            try:
                btn_export = self.session.findById("wnd[0]/tbar[1]/btn[45]")
                btn_export.press()
                self.wait_ready()
            except Exception as ex:
                if self._is_rpc_or_session_error(ex):
                    raise
                self.reset_navigation()
                return ExportResult(
                    success=False,
                    material=material,
                    error_message=f"Failed to trigger export btn[45]: {ex}",
                    status_code="E",
                    execution_time_sec=time.time() - start_time,
                )

            # 6. Select Spreadsheet radio button in SAPLSPO5:0150 dialog
            try:
                radio_spreadsheet = self.session.findById(
                    "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]"
                )
                if radio_spreadsheet is not None and hasattr(radio_spreadsheet, "Select"):
                    radio_spreadsheet.Select()

                # Confirm selection dialog (btn[0])
                btn_ok_dialog = self.session.findById("wnd[1]/tbar[0]/btn[0]")
                btn_ok_dialog.press()
                self.wait_ready()
            except Exception as ex:
                if self._is_rpc_or_session_error(ex):
                    raise
                self.reset_navigation()
                return ExportResult(
                    success=False,
                    material=material,
                    error_message=f"Failed during SAPLSPO5:0150 format selection: {ex}",
                    status_code="E",
                    execution_time_sec=time.time() - start_time,
                )

            # 7. File Path & Filename Input in wnd[1]
            try:
                path_field = self.session.findById("wnd[1]/usr/ctxtDY_PATH")
                path_field.Text = clean_dir_str

                file_field = self.session.findById("wnd[1]/usr/ctxtDY_FILENAME")
                file_field.Text = params.expected_filename

                btn_save = self.session.findById("wnd[1]/tbar[0]/btn[0]")
                btn_save.press()
                self.wait_ready()
            except Exception as ex:
                if self._is_rpc_or_session_error(ex):
                    raise
                self.reset_navigation()
                return ExportResult(
                    success=False,
                    material=material,
                    error_message=f"Failed to submit export path and filename: {ex}",
                    status_code="E",
                    execution_time_sec=time.time() - start_time,
                )

            # 8. Handle Overwrite Confirmation popup (wnd[2] / btnSPOP-OPTION1)
            try:
                children = getattr(self.session, "Children", None)
                children_count = getattr(children, "Count", 0) if children is not None else 0
                has_multi_wnd = isinstance(children_count, int) and children_count > 1
                if has_multi_wnd or hasattr(self.session, "findById"):
                    try:
                        btn_replace = self.session.findById("wnd[2]/usr/btnSPOP-OPTION1")
                        if btn_replace is not None and hasattr(btn_replace, "press"):
                            logger.info("File already exists on disk; confirming overwrite in SAP popup.")
                            btn_replace.press()
                            self.wait_ready()
                    except Exception as inner_ex:
                        if self._is_rpc_or_session_error(inner_ex):
                            raise
            except Exception as ex:
                if self._is_rpc_or_session_error(ex):
                    raise
                # Overwrite popup did not appear
                pass

            # 9. Return to main screen (btn[3] / F3)
            try:
                btn_back = self.session.findById("wnd[0]/tbar[0]/btn[3]")
                btn_back.press()
                self.wait_ready()
            except Exception as ex:
                if self._is_rpc_or_session_error(ex):
                    # Check if file was already successfully written to disk during this execution
                    if (
                        expected_file.exists()
                        and expected_file.stat().st_size > 0
                        and expected_file.stat().st_mtime >= start_time
                    ):
                        logger.warning(
                            f"SAP session disconnected during post-export F3 return, but BOM file was successfully written: {expected_file}"
                        )
                        self.session = None
                        return ExportResult(
                            success=True,
                            material=material,
                            file_path=expected_file,
                            status_code="S",
                            execution_time_sec=time.time() - start_time,
                        )
                    raise
                pass

            # 10. Verify exported file exists and is non-empty
            if expected_file.exists() and expected_file.stat().st_size > 0:
                logger.info(f"Successfully exported CS12 BOM to: {expected_file}")
                return ExportResult(
                    success=True,
                    material=material,
                    file_path=expected_file,
                    status_code="S",
                    execution_time_sec=time.time() - start_time,
                )
            else:
                return ExportResult(
                    success=False,
                    material=material,
                    error_message=f"Export finished but target file not found or empty at: {expected_file}",
                    status_code="E",
                    execution_time_sec=time.time() - start_time,
                )

        except Exception as ex:
            is_rpc = self._is_rpc_or_session_error(ex)
            if allow_retry and self.connection_manager is not None and is_rpc:
                warn_msg = (
                    f"[-] [SAP R3] Phát hiện mất kết nối COM/RPC SAP khi xử lý mã '{material}': {ex}. "
                    "Nguyên nhân có thể do tài khoản bị đăng nhập ở nơi khác (Multi-Logon kick), "
                    "phiên làm việc bị ngắt hoặc mạng gián đoạn. Đang tự động kết nối lại và thử tải lại..."
                )
                logger.warning(warn_msg)
                if self.status_callback is not None:
                    try:
                        self.status_callback(warn_msg)
                    except Exception:
                        pass

                reconnected = self.reacquire_session()
                if reconnected:
                    info_msg = f"[+] [SAP R3] Tái lập kết nối SAP GUI thành công! Đang thử tải lại mã BOM '{material}'..."
                    logger.info(info_msg)
                    if self.status_callback is not None:
                        try:
                            self.status_callback(info_msg)
                        except Exception:
                            pass
                    return self.execute_cs12_and_export(params, allow_retry=False)
                else:
                    err_msg = f"[-] [SAP R3] Không thể tái kết nối SAP GUI để thử lại mã '{material}'."
                    logger.error(err_msg)
                    if self.status_callback is not None:
                        try:
                            self.status_callback(err_msg)
                        except Exception:
                            pass

            logger.exception(f"Unexpected exception during CS12 export for '{material}': {ex}")
            if not is_rpc:
                self.reset_navigation()
            else:
                self.session = None

            error_msg_detail = (
                f"Mất kết nối COM/RPC SAP: {ex}. Đã thử kết nối lại nhưng không thành công."
                if is_rpc and allow_retry and self.connection_manager is not None
                else str(ex)
            )

            return ExportResult(
                success=False,
                material=material,
                error_message=error_msg_detail,
                status_code="E",
                execution_time_sec=time.time() - start_time,
            )

    def download_bom(
        self,
        material_code: str,
        valid_date: date,
        output_dir: Union[str, Path],
        plant: str = "2200",
        bom_usage: str = "pp01",
        alternative: str = "01",
    ) -> Path:
        """
        Application Layer Interface Contract:
        download_bom(material_code: str, valid_date: datetime.date, output_dir: Path) -> Path

        Returns:
            Path to downloaded R3_*.xls file.

        Raises:
            SAPCS12Error: If export fails or file not generated.
        """
        output_path = Path(output_dir)
        params = CS12Params(
            material=material_code,
            plant=plant,
            bom_usage=bom_usage,
            alternative=alternative,
            valid_date=valid_date,
            destination_dir=output_path,
        )
        result = self.execute_cs12_and_export(params)
        if not result.success or not result.file_path:
            raise SAPCS12Error(
                f"Failed to download BOM for material '{material_code}': {result.error_message}"
            )
        return result.file_path

    def batch_download(
        self,
        materials: List[str],
        valid_date: date,
        base_destination_dir: Union[str, Path],
        plant: str = "2200",
        bom_usage: str = "pp01",
        alternative: str = "01",
        archive_existing: bool = True,
    ) -> List[ExportResult]:
        """
        Unifies machine model code processing across model families (e.g. '110*' ma1 and 'T10*' maT).
        Routes each material to its own dedicated subfolder:
            base_destination_dir / <material> / R3_<material>_<dd>_<mm>_<yyyy>.xls

        Optionally archives existing R3 files into:
            base_destination_dir / <material> / capnhat / old /
        """
        base_dir = Path(base_destination_dir)
        results = []

        for mat in materials:
            mat_clean = mat.strip()
            if not mat_clean:
                continue

            target_folder = base_dir / mat_clean
            target_folder.mkdir(parents=True, exist_ok=True)

            # Archive existing older files if present
            if archive_existing:
                self._archive_existing_files(target_folder, mat_clean)

            params = CS12Params(
                material=mat_clean,
                plant=plant,
                bom_usage=bom_usage,
                alternative=alternative,
                valid_date=valid_date,
                destination_dir=target_folder,
            )

            res = self.execute_cs12_and_export(params)
            results.append(res)

        return results

    @staticmethod
    def _archive_existing_files(folder: Path, material: str) -> None:
        """Move existing R3_*.xls files to a subfolder capnhat/old/ to prevent silent overwriting."""
        archive_dir = folder / "capnhat" / "old"
        pattern = f"R3_{material}_*.xls"
        existing_files = list(folder.glob(pattern))

        if existing_files:
            archive_dir.mkdir(parents=True, exist_ok=True)
            for file in existing_files:
                timestamp_tag = datetime.fromtimestamp(file.stat().st_mtime).strftime("%Y%m%d_%H%M%S")
                archived_name = f"{file.stem}_{timestamp_tag}{file.suffix}"
                target = archive_dir / archived_name
                try:
                    shutil.move(str(file), str(target))
                    logger.debug(f"Archived previous export {file.name} -> {target}")
                except Exception as ex:
                    logger.warning(f"Failed to archive existing file {file}: {ex}")


# Alias matching PROJECT.md interface contract specification
SAPCS12Client = CS12Service
