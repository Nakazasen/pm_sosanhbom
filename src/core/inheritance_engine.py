"""Explanation Inheritance and Multi-Version Archiving Engine.

Authoritative implementation for Feature F13/F14/F15/F16:
- Implements legacy ham_match_index_mix algorithm (from capnhat_PLM_R3.bas lines 315-388):
  * Matches Item ID (Col C) between new BOM and historical PLM sheet.
  * Preserves Explanation (Col O / giai_thich).
  * Preserves Responsible Person (Col P / phu_trach).
  * Preserves Manager Check Status (Col Q / quan_ly_check).
  * Leaves new / unmatched parts cleanly empty ("") without 'NaN'.
- Multi-version sheet archiving:
  * Creates PLM_old if not present.
  * Increments to PLM_old_1, PLM_old_2... if PLM_old already exists.
  * Preserves historical sheets without overwriting past versions.
- Raw file archival:
  * Moves superseded PLM and R3 exports to 'capnhat\\old\\'.
- Formula maintenance & Pivot Table cache refresh:
  * Populates Col R (=IF(C{r}="","",C{r})) and Col S (=IF(E{r}="","",E{r})).
  * Refreshes Pivot Tables via win32com with graceful headless fallback.
"""

from __future__ import annotations

import copy
from dataclasses import dataclass, field
import datetime
import logging
from pathlib import Path
import shutil
from typing import Any

import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
import pandas as pd

logger = logging.getLogger(__name__)


@dataclass
class ExplanationRecord:
    """Stored explanation and verification metadata for an engineering part."""

    part_code: str
    explanation: str = ""
    person_in_charge: str = ""
    manager_check: str = ""


@dataclass
class InheritanceSummary:
    """Result summary of the explanation inheritance migration process."""

    total_new_parts: int
    inherited_count: int
    new_unannotated_count: int
    archived_sheet_name: str
    archived_files: list[Path] = field(default_factory=list)
    success: bool = True
    message: str = ""


class InheritanceEngine:
    """Carries over human annotations from historical BOM sheets to updated exports."""

    @staticmethod
    def match_index_mix(
        df_new: pd.DataFrame,
        df_old: pd.DataFrame,
        key_column: str = "item_id",
    ) -> pd.DataFrame:
        """Pure-Python vectorized implementation of legacy ham_match_index_mix.

        Matches parts by Item ID / Part Code (Col C) and migrates:
        - Explanation (Col O -> giai_thich)
        - Person in Charge (Col P -> phu_trach)
        - Manager Check (Col Q -> quan_ly_check)

        Args:
            df_new: DataFrame of newly ingested BOM items.
            df_old: DataFrame of historical BOM items with existing annotations.
            key_column: Column name containing part code / item id in df_new.

        Returns:
            pd.DataFrame: df_new enriched with inherited annotations in Col O, P, Q.
        """
        df_res = df_new.copy()
        if df_res.empty:
            if "giai_thich" not in df_res.columns:
                df_res["giai_thich"] = pd.Series(dtype="object")
            if "phu_trach" not in df_res.columns:
                df_res["phu_trach"] = pd.Series(dtype="object")
            if "quan_ly_check" not in df_res.columns:
                df_res["quan_ly_check"] = pd.Series(dtype="object")
            return df_res

        # Resolve key column in df_new
        clean_key = "_clean_key_match"
        if key_column in df_res.columns:
            actual_new_key = key_column
        else:
            actual_new_key = next(
                (
                    c
                    for c in df_res.columns
                    if any(
                        k in str(c).lower()
                        for k in ["item_id", "part_code", "part code", "mã", "code", "item id"]
                    )
                ),
                df_res.columns[0],
            )

        # Build list of normalized search keys directly (avoid mutating/dropping columns)
        search_keys = [str(x).strip().upper() for x in df_res[actual_new_key]]

        # If df_old is empty, populate empty strings for all rows
        if df_old is None or df_old.empty:
            df_res["giai_thich"] = [""] * len(df_res)
            df_res["phu_trach"] = [""] * len(df_res)
            df_res["quan_ly_check"] = [""] * len(df_res)
            return df_res

        # Resolve columns in df_old
        old_key_col = next(
            (
                c
                for c in df_old.columns
                if any(
                    k in str(c).lower()
                    for k in [
                        "item_id",
                        "part_code",
                        "part code",
                        "mã",
                        "code",
                        "item id",
                        "col c",
                    ]
                )
            ),
            df_old.columns[0],
        )

        old_exp_col = next(
            (
                c
                for c in df_old.columns
                if any(
                    k in str(c).lower()
                    for k in [
                        "giai_thich",
                        "giải thích",
                        "explanation",
                        "note",
                        "gt",
                        "col o",
                    ]
                )
            ),
            None,
        )

        old_pic_col = next(
            (
                c
                for c in df_old.columns
                if any(
                    k in str(c).lower()
                    for k in [
                        "phu_trach",
                        "phụ trách",
                        "person",
                        "author",
                        "tennpt",
                        "col p",
                    ]
                )
            ),
            None,
        )

        old_mgr_col = next(
            (
                c
                for c in df_old.columns
                if any(
                    k in str(c).lower()
                    for k in [
                        "quan_ly_check",
                        "quản lý",
                        "manager",
                        "check_ql",
                        "quanly",
                        "col q",
                    ]
                )
            ),
            None,
        )

        # Build lookup table of first occurrences (matching MATCH(..., 0))
        lookup: dict[str, tuple[str, str, str]] = {}
        for _, row in df_old.iterrows():
            raw_k = row[old_key_col]
            if pd.isna(raw_k):
                continue
            k = str(raw_k).strip().upper()
            if not k or k in ("NAN", "NONE"):
                continue

            # First match wins (standard MATCH behavior)
            if k in lookup:
                continue

            exp = ""
            if old_exp_col and pd.notna(row[old_exp_col]):
                s_val = str(row[old_exp_col]).strip()
                if s_val.lower() not in ("nan", "none"):
                    exp = s_val

            pic = ""
            if old_pic_col and pd.notna(row[old_pic_col]):
                s_val = str(row[old_pic_col]).strip()
                if s_val.lower() not in ("nan", "none"):
                    pic = s_val

            mgr = ""
            if old_mgr_col and pd.notna(row[old_mgr_col]):
                s_val = str(row[old_mgr_col]).strip()
                if s_val.lower() not in ("nan", "none"):
                    mgr = s_val

            lookup[k] = (exp, pic, mgr)

        explanations: list[str] = []
        persons: list[str] = []
        mgr_checks: list[str] = []

        for k in search_keys:
            if k in lookup:
                rec = lookup[k]
                explanations.append(rec[0])
                persons.append(rec[1])
                mgr_checks.append(rec[2])
            else:
                explanations.append("")
                persons.append("")
                mgr_checks.append("")

        df_res["giai_thich"] = explanations
        df_res["phu_trach"] = persons
        df_res["quan_ly_check"] = mgr_checks
        return df_res

    def determine_next_archive_sheet_name(
        self, sheet_names: list[str], base_name: str = "PLM_old"
    ) -> str:
        """Determine sequential archival sheet name: PLM_old, PLM_old_1, PLM_old_2..."""
        if base_name not in sheet_names:
            return base_name

        idx = 1
        while f"{base_name}_{idx}" in sheet_names:
            idx += 1
        return f"{base_name}_{idx}"

    def extract_existing_plm_records(self, ws: Worksheet) -> pd.DataFrame:
        """Read historical annotations and part codes from existing Sheet 'PLM'.

        Reads:
        - Col C (3): Item ID / Part Code
        - Col O (15): Explanation (giai_thich)
        - Col P (16): Person in charge (phu_trach)
        - Col Q (17): Manager check (quan_ly_check)
        """
        records: list[dict[str, Any]] = []
        max_row = ws.max_row or 1
        for r in range(2, max_row + 1):
            val_c = ws.cell(row=r, column=3).value
            if val_c is not None:
                s_val = str(val_c).strip()
                if s_val and s_val.lower() not in ("nan", "none"):
                    exp_val = ws.cell(row=r, column=15).value or ""
                    pic_val = ws.cell(row=r, column=16).value or ""
                    mgr_val = ws.cell(row=r, column=17).value or ""
                    records.append(
                        {
                            "item_id": s_val,
                            "giai_thich": str(exp_val).strip()
                            if str(exp_val).lower() not in ("nan", "none")
                            else "",
                            "phu_trach": str(pic_val).strip()
                            if str(pic_val).lower() not in ("nan", "none")
                            else "",
                            "quan_ly_check": str(mgr_val).strip()
                            if str(mgr_val).lower() not in ("nan", "none")
                            else "",
                        }
                    )
        return pd.DataFrame(records)

    def archive_superseded_file(
        self,
        file_to_archive: str | Path,
        archive_dir: str | Path,
    ) -> Path:
        """Safely move a superseded file into the archive directory."""
        src = Path(file_to_archive)
        if not src.exists():
            raise FileNotFoundError(f"Source file not found: {src}")

        dest_dir = Path(archive_dir)
        dest_dir.mkdir(parents=True, exist_ok=True)

        dest = dest_dir / src.name
        if dest.exists() and dest.resolve() != src.resolve():
            # Generate unique timestamped filename if collision
            stem = src.stem
            ext = src.suffix
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            dest = dest_dir / f"{stem}_{timestamp}{ext}"

        shutil.move(str(src), str(dest))
        return dest

    def update_workbook_with_inheritance(
        self,
        workbook_path: str | Path,
        new_plm_data: pd.DataFrame,
        archive_dir: str | Path | None = None,
        source_raw_file: str | Path | None = None,
    ) -> InheritanceSummary:
        """Perform full multi-version update on form_ssbom workbook.

        Steps:
        1. Open workbook, access Sheet 'PLM'.
        2. Extract existing records (Col C, O, P, Q).
        3. Determine archival sheet name: 'PLM_old' or 'PLM_old_<N>'.
        4. Copy current Sheet 'PLM' to archival sheet preserving cells and formulas.
        5. Run match_index_mix to migrate human annotations.
        6. Overwrite Sheet 'PLM' rows 2+ with new data + inherited annotations.
        7. Populate formula network for Col R (=IF(C{r}="","",C{r})) and Col S (=IF(E{r}="","",E{r})).
        8. Clear any old rows beyond new data length to avoid ghost data.
        9. Save workbook.
        10. Move superseded raw files to capnhat\\old\\ if requested.
        """
        wb_path = Path(workbook_path)
        if not wb_path.exists():
            raise FileNotFoundError(f"Workbook not found: {wb_path}")

        wb = openpyxl.load_workbook(wb_path)
        if "PLM" not in wb.sheetnames:
            wb.close()
            raise ValueError(f"Sheet 'PLM' does not exist in target workbook: {wb_path}")

        ws_plm = wb["PLM"]

        # 1. Read existing PLM data
        df_old = self.extract_existing_plm_records(ws_plm)

        # 2. Determine archive sheet name
        archive_sheet_name = self.determine_next_archive_sheet_name(wb.sheetnames, "PLM_old")

        # 3. Create duplicate of current Sheet 'PLM'
        try:
            ws_archive = wb.copy_worksheet(ws_plm)
            ws_archive.title = archive_sheet_name
        except Exception as copy_err:
            logger.warning("copy_worksheet error (%s), using cell-by-cell copy", copy_err)
            ws_archive = wb.create_sheet(title=archive_sheet_name)
            for row in ws_plm.iter_rows(values_only=False):
                for cell in row:
                    dest_c = ws_archive.cell(row=cell.row, column=cell.column, value=cell.value)
                    if cell.has_style:
                        dest_c.font = copy.copy(cell.font)
                        dest_c.border = copy.copy(cell.border)
                        dest_c.fill = copy.copy(cell.fill)
                        dest_c.number_format = cell.number_format
                        dest_c.alignment = copy.copy(cell.alignment)

        # 4. Migrate annotations via match_index_mix
        df_inherited = self.match_index_mix(new_plm_data, df_old, key_column="item_id")

        # 5. Populate Sheet 'PLM'
        row_start = 2
        old_max_row = max(ws_plm.max_row or 1, 100)

        # Clear existing data rows in A2:S (matching VBA Range("A2:M5000").ClearContents & Range("O2:Q5000").ClearContents)
        for r in range(row_start, old_max_row + 1):
            for c in range(1, 20):
                ws_plm.cell(row=r, column=c).value = None

        inherited_count = 0
        new_unannotated_count = 0

        for idx, row_data in df_inherited.iterrows():
            r = row_start + idx

            # Col A: Level
            ws_plm.cell(row=r, column=1).value = row_data.get("level", 1)
            # Col B: Item Type
            ws_plm.cell(row=r, column=2).value = row_data.get("item_type", "Part")
            # Col C: Item Id
            item_id_val = str(row_data.get("item_id", "")).strip()
            ws_plm.cell(row=r, column=3).value = item_id_val
            # Col D: Has Children
            ws_plm.cell(row=r, column=4).value = str(row_data.get("has_children", "False"))
            # Col E: Quantity
            try:
                qty_val = float(row_data.get("quantity", 1.0))
            except (ValueError, TypeError):
                qty_val = 1.0
            ws_plm.cell(row=r, column=5).value = qty_val
            # Col F: 1st Parts
            ws_plm.cell(row=r, column=6).value = row_data.get("first_parts", "")
            # Col G: 2nd BOM Flag
            ws_plm.cell(row=r, column=7).value = row_data.get("second_bom_flag", "")
            # Col H: Occurrence Effectivities
            ws_plm.cell(row=r, column=8).value = row_data.get(
                "effectivity", row_data.get("occurrence_effectivities", "")
            )
            # Col I: Project List
            ws_plm.cell(row=r, column=9).value = row_data.get(
                "project_list", row_data.get("item_revision_projects_list", "")
            )
            # Col J: Item Name
            ws_plm.cell(row=r, column=10).value = row_data.get("item_name", "")
            # Col K: Notice No
            ws_plm.cell(row=r, column=11).value = row_data.get("notice_no", "")
            # Col L: Revision
            ws_plm.cell(row=r, column=12).value = row_data.get("revision", "")
            # Col M: Item Rev Status
            ws_plm.cell(row=r, column=13).value = row_data.get("item_rev_status", "")

            # Col O: Explanation
            exp_val = str(row_data.get("giai_thich", "")).strip()
            ws_plm.cell(row=r, column=15).value = exp_val if exp_val else None

            # Col P: Person in Charge
            pic_val = str(row_data.get("phu_trach", "")).strip()
            ws_plm.cell(row=r, column=16).value = pic_val if pic_val else None

            # Col Q: Manager Check
            mgr_val = str(row_data.get("quan_ly_check", "")).strip()
            ws_plm.cell(row=r, column=17).value = mgr_val if mgr_val else None

            # Col R & S Formulas
            ws_plm.cell(row=r, column=18).value = f'=IF(C{r}="","",C{r})'
            ws_plm.cell(row=r, column=19).value = f'=IF(E{r}="","",E{r})'

            if exp_val or pic_val or mgr_val:
                inherited_count += 1
            else:
                new_unannotated_count += 1

        wb.save(wb_path)
        wb.close()

        # 6. File Archiving to capnhat\old\
        archived_files: list[Path] = []
        target_archive_dir = Path(archive_dir) if archive_dir else (wb_path.parent / "capnhat" / "old")
        target_archive_dir.mkdir(parents=True, exist_ok=True)

        if source_raw_file:
            src_f = Path(source_raw_file)
            if src_f.exists() and src_f.resolve() != wb_path.resolve():
                archived_path = self.archive_superseded_file(src_f, target_archive_dir)
                archived_files.append(archived_path)

        return InheritanceSummary(
            total_new_parts=len(df_inherited),
            inherited_count=inherited_count,
            new_unannotated_count=new_unannotated_count,
            archived_sheet_name=archive_sheet_name,
            archived_files=archived_files,
            success=True,
            message=f"Migrated {inherited_count} annotations. Archived to '{archive_sheet_name}'.",
        )

    def refresh_workbook_pivots(self, workbook_path: str | Path) -> bool:
        """Trigger Pivot Cache refresh via win32com or Excel automation.

        Safely handles environments without Windows Excel GUI (e.g. headless/CI).
        """
        wb_resolved = Path(workbook_path).resolve()
        if not wb_resolved.exists():
            logger.error("Workbook does not exist: %s", wb_resolved)
            return False

        pythoncom = None
        excel = None
        co_initialized = False
        try:
            try:
                import pythoncom

                pythoncom.CoInitialize()
                co_initialized = True
            except Exception:
                pass

            import win32com.client

            excel = win32com.client.Dispatch("Excel.Application")
            excel.DisplayAlerts = False
            excel.Visible = False

            wb = excel.Workbooks.Open(str(wb_resolved))
            wb.RefreshAll()
            wb.Save()
            wb.Close()
            excel.Quit()
            excel = None
            logger.info("Successfully refreshed Pivot Tables via win32com for %s", wb_resolved)
            return True
        except Exception as exc:
            logger.info(
                "win32com Pivot Table refresh skipped/unavailable in current environment: %s", exc
            )
            return False
        finally:
            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass
            if co_initialized and pythoncom is not None:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass


# Convenience module-level aliases
def match_index_mix(
    df_new: pd.DataFrame,
    df_old: pd.DataFrame,
    key_column: str = "item_id",
) -> pd.DataFrame:
    """Convenience functional interface for ham_match_index_mix."""
    return InheritanceEngine.match_index_mix(df_new, df_old, key_column=key_column)


def update_workbook_with_inheritance(
    workbook_path: str | Path,
    new_plm_data: pd.DataFrame,
    archive_dir: str | Path | None = None,
    source_raw_file: str | Path | None = None,
) -> InheritanceSummary:
    """Convenience functional interface for workbook inheritance update."""
    engine = InheritanceEngine()
    return engine.update_workbook_with_inheritance(
        workbook_path=workbook_path,
        new_plm_data=new_plm_data,
        archive_dir=archive_dir,
        source_raw_file=source_raw_file,
    )


def refresh_workbook_pivots(workbook_path: str | Path) -> bool:
    """Convenience functional interface for Pivot Table refresh."""
    engine = InheritanceEngine()
    return engine.refresh_workbook_pivots(workbook_path)
