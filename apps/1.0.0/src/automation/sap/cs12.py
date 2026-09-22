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
from typing import Any, List, Optional, Union

from src.automation.sap.models import (
    CS12Params,
    ExportResult,
    SAPCS12Error,
)

logger = logging.getLogger(__name__)


class CS12Service:
    """Automates SAP Transaction CS12 (Multilevel BOM Display & Export)."""

    def __init__(self, session: Any = None, connection_manager: Optional[Any] = None):
        """
        Initialize CS12Service.

        Args:
            session: Active, authenticated SAP GUI session object.
            connection_manager: Optional SAPConnectionManager providing active session.
        """
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
            okcd = self.session.findById("wnd[0]/tbar[0]/okcd")
            okcd.Text = "/nex"
            wnd0 = self.session.findById("wnd[0]")
            wnd0.sendVKey(0)
            return True
        except Exception as ex:
            logger.debug(f"Failed logging off SAP via /nex: {ex}")
            return False

    def execute_cs12_and_export(self, params: CS12Params) -> ExportResult:
        """
        Execute CS12 query for a single material BOM and export to spreadsheet.

        Args:
            params: Parameters including material, plant, bom_usage, alternative, valid_date.

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
                children_count = getattr(getattr(self.session, "Children", None), "Count", 0)
                if children_count > 1 or hasattr(self.session, "findById"):
                    btn_replace = self.session.findById("wnd[2]/usr/btnSPOP-OPTION1")
                    if btn_replace is not None:
                        logger.info("File already exists on disk; confirming overwrite in SAP popup.")
                        btn_replace.press()
                        self.wait_ready()
            except Exception:
                # Overwrite popup did not appear
                pass

            # 9. Return to main screen (btn[3] / F3)
            try:
                btn_back = self.session.findById("wnd[0]/tbar[0]/btn[3]")
                btn_back.press()
                self.wait_ready()
            except Exception:
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
            logger.exception(f"Unexpected exception during CS12 export for '{material}': {ex}")
            self.reset_navigation()
            return ExportResult(
                success=False,
                material=material,
                error_message=str(ex),
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
