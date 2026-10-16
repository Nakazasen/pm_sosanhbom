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
import os
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


_VBA_SUFFIXES = (".xlsm", ".xltm")
_MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3


def _col_letter(n: int) -> str:
    from openpyxl.utils import get_column_letter

    return get_column_letter(n)


def backup_workbook(wb_path: Path, backup_dir: Path) -> Path:
    """Copy workbook to backup_dir before any write so the user can always recover."""
    backup_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = backup_dir / f"{wb_path.stem}_backup_{stamp}{wb_path.suffix}"
    shutil.copy2(wb_path, dest)
    logger.info("Backed up workbook to %s", dest)
    return dest


def excel_com_update_sheet(
    wb_path: Path,
    sheet_name: str,
    archive_name: str,
    rows: list[list[Any]],
    n_cols: int,
    headers: list[str] | None = None,
) -> bool:
    """Archive + rewrite a sheet through real Excel (preserves VBA, buttons, shapes, pivots).

    Returns False when Excel COM is unavailable so the caller can fall back to openpyxl.
    Raises on genuine Excel errors (e.g. file locked) to avoid silently corrupting data.
    """
    if os.environ.get("SSBOM_DISABLE_EXCEL_COM") == "1":
        return False
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        return False

    pythoncom.CoInitialize()
    excel = None
    wb = None
    try:
        try:
            # DispatchEx: private instance, never attaches to / quits the user's open Excel.
            excel = win32com.client.DispatchEx("Excel.Application")
        except Exception as exc:
            logger.info("Excel COM unavailable (%s); falling back to openpyxl", exc)
            return False

        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        try:
            excel.AutomationSecurity = _MSO_AUTOMATION_SECURITY_FORCE_DISABLE
        except Exception:
            pass

        try:
            wb = excel.Workbooks.Open(str(wb_path.resolve()), UpdateLinks=0, ReadOnly=False)
        except Exception as exc:
            logger.warning("Excel cannot open %s (%s); falling back to openpyxl", wb_path.name, exc)
            return False
        if wb.ReadOnly:
            raise PermissionError(f"File đang được mở bởi người khác hoặc chỉ đọc: {wb_path.name}")

        def sheet_names() -> list[str]:
            return [wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)]

        names = sheet_names()
        if sheet_name in names:
            ws = wb.Worksheets(sheet_name)
        else:
            ws = wb.Worksheets.Add(pythoncom.Empty, wb.Worksheets(wb.Worksheets.Count))
            ws.Name = sheet_name
            names = sheet_names()

        # Worksheet.Copy(Before, After) — positional: late-bound COM drops the After= keyword
        # and would copy into a brand-new workbook instead.
        ws.Copy(pythoncom.Empty, wb.Worksheets(wb.Worksheets.Count))
        added = [n for n in sheet_names() if n not in names]
        if len(added) != 1:
            raise RuntimeError(f"Không tạo được sheet lưu trữ cho '{sheet_name}' (added={added})")
        wb.Worksheets(added[0]).Name = archive_name

        if headers and not str(ws.Range("A1").Value or "").strip():
            ws.Range(ws.Cells(1, 1), ws.Cells(1, len(headers))).Value = [headers]

        # Active AutoFilter: ClearContents / Value= only touch VISIBLE rows -> scrambled data.
        had_filter = bool(ws.FilterMode)
        if had_filter:
            ws.ShowAllData()

        last_col = _col_letter(n_cols)
        used = ws.UsedRange
        last_row = max(used.Row + used.Rows.Count - 1, len(rows) + 1, 2)
        ws.Range(f"A2:{last_col}{last_row}").ClearContents()

        if rows:
            data = [tuple("" if v is None else v for v in r) for r in rows]
            ws.Range(f"A2:{last_col}{len(rows) + 1}").Value = data

        if had_filter:
            try:
                ws.AutoFilter.ApplyFilter()
            except Exception as exc:
                logger.warning("Could not re-apply AutoFilter on '%s': %s", sheet_name, exc)

        try:
            wb.RefreshAll()
            excel.CalculateUntilAsyncQueriesDone()
        except Exception as exc:
            logger.warning("Pivot refresh failed (data still saved): %s", exc)

        wb.Save()
        logger.info("Excel COM updated sheet '%s' in %s", sheet_name, wb_path)
        return True
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def load_workbook_preserving(wb_path: Path) -> openpyxl.Workbook:
    """openpyxl fallback: keep VBA so a .xlsm stays a valid macro-enabled package."""
    return openpyxl.load_workbook(wb_path, keep_vba=wb_path.suffix.lower() in _VBA_SUFFIXES)


class InheritanceEngine:
    """Carries over human annotations from historical BOM sheets to updated exports."""

    R3_HEADERS = [
        "Level",
        "Component (Material)",
        "Object Description",
        "Component Quantity",
        "Unit of Measure",
        "Revision Level (RevLev)",
        "Valid From",
        "Valid To",
    ]

    @staticmethod
    def _build_plm_rows(df_inherited: pd.DataFrame) -> tuple[list[list[Any]], int, int]:
        """Build PLM rows A..S (19 cols). Returns (rows, inherited_count, new_unannotated_count)."""
        rows: list[list[Any]] = []
        inherited = unannotated = 0
        for i, (_, rd) in enumerate(df_inherited.iterrows()):
            r = i + 2
            try:
                qty_val = float(rd.get("quantity", 1.0))
            except (ValueError, TypeError):
                qty_val = 1.0
            exp_val = str(rd.get("giai_thich", "")).strip()
            pic_val = str(rd.get("phu_trach", "")).strip()
            mgr_val = str(rd.get("quan_ly_check", "")).strip()
            rows.append([
                rd.get("level", 1),
                rd.get("item_type", "Part"),
                str(rd.get("item_id", "")).strip(),
                str(rd.get("has_children", "False")),
                qty_val,
                rd.get("first_parts", ""),
                rd.get("second_bom_flag", ""),
                rd.get("effectivity", rd.get("occurrence_effectivities", "")),
                rd.get("project_list", rd.get("item_revision_projects_list", "")),
                rd.get("item_name", ""),
                rd.get("notice_no", ""),
                rd.get("revision", ""),
                rd.get("item_rev_status", ""),
                None,  # Col N untouched by VBA
                exp_val or None,
                pic_val or None,
                mgr_val or None,
                f'=IF(C{r}="","",C{r})',
                f'=IF(E{r}="","",E{r})',
            ])
            if exp_val or pic_val or mgr_val:
                inherited += 1
            else:
                unannotated += 1
        return [[InheritanceEngine._clean(v) for v in row] for row in rows], inherited, unannotated

    @staticmethod
    def _clean(v: Any) -> Any:
        """Convert pandas/numpy scalars & NaN into plain Python values safe for COM/openpyxl."""
        if v is None:
            return None
        try:
            if pd.isna(v):
                return None
        except (TypeError, ValueError):
            pass
        if hasattr(v, "item"):
            return v.item()
        return v

    @staticmethod
    def _build_r3_rows(df_new: pd.DataFrame) -> list[list[Any]]:
        """Build R3 rows A..U (21 cols): data in A..H, pivot helper formulas in T/U."""

        def text(rd: Any, *keys: str, default: str = "") -> str:
            for k in keys:
                if k in rd and not pd.isna(rd[k]):
                    return str(rd[k]).strip()
            return default

        rows: list[list[Any]] = []
        for i, (_, rd) in enumerate(df_new.reset_index(drop=True).iterrows()):
            r = i + 2
            lvl = rd.get("level", rd.get("Level", 1))
            try:
                lvl_val = int(lvl) if pd.notna(lvl) else 1
            except (ValueError, TypeError):
                lvl_val = 1
            qty = rd.get("quantity", rd.get("Component Quantity", 0.0))
            try:
                qty_val = float(qty) if pd.notna(qty) else 0.0
            except (ValueError, TypeError):
                qty_val = 0.0
            row: list[Any] = [None] * 21
            row[0] = lvl_val
            row[1] = text(rd, "material", "part_code", "Component")
            row[2] = text(rd, "description", "part_name", "Object Description")
            row[3] = qty_val
            row[4] = text(rd, "uom", "Unit of Measure") or "PC"
            row[5] = text(rd, "revision", "rev_r3", "RevLev")
            row[6] = text(rd, "valid_from", "Valid From")
            row[7] = text(rd, "valid_to", "Valid To")
            row[19] = f'=IF(B{r}="","",B{r})'
            row[20] = f'=IF(D{r}="","",D{r})'
            rows.append(row)
        return rows

    @staticmethod
    def _write_rows_openpyxl(
        wb_path: Path,
        sheet_name: str,
        archive_name: str,
        rows: list[list[Any]],
        n_cols: int,
        headers: list[str] | None = None,
    ) -> None:
        """Fallback when Excel is unavailable. keep_vba keeps .xlsm valid (form buttons may be lost)."""
        logger.warning(
            "Excel COM unavailable: writing '%s' via openpyxl (form controls/shapes may not be preserved)",
            sheet_name,
        )
        wb = load_workbook_preserving(wb_path)
        try:
            ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.create_sheet(title=sheet_name)
            try:
                wb.copy_worksheet(ws).title = archive_name
            except Exception as copy_err:
                logger.warning("copy_worksheet error (%s), using value copy", copy_err)
                ws_archive = wb.create_sheet(title=archive_name)
                for row in ws.iter_rows():
                    for cell in row:
                        ws_archive.cell(row=cell.row, column=cell.column, value=cell.value)

            if headers and not str(ws.cell(row=1, column=1).value or "").strip():
                for c, h in enumerate(headers, start=1):
                    ws.cell(row=1, column=c, value=h)

            last_row = max(ws.max_row or 1, len(rows) + 1, 100)
            for r in range(2, last_row + 1):
                for c in range(1, n_cols + 1):
                    ws.cell(row=r, column=c).value = None
            for i, row in enumerate(rows):
                for c, v in enumerate(row, start=1):
                    if v is not None:
                        ws.cell(row=i + 2, column=c).value = v
            wb.save(wb_path)
        finally:
            wb.close()

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
        new_plm_data: pd.DataFrame | str | Path | Any,
        archive_dir: str | Path | None = None,
        source_raw_file: str | Path | None = None,
    ) -> InheritanceSummary:
        """Perform full multi-version update on form_ssbom workbook.

        Steps:
        1. Open workbook, access Sheet 'PLM'.
        2. Extract existing records (Col C, O, P, Q).
        3. Determine archival sheet name: 'PLM_old' or 'PLM_old_<N>'.
        4. Copy current Sheet 'PLM' to archival sheet preserving cells and formulas.
        5. Normalize new_plm_data (parse file or DataFrame).
        6. Run match_index_mix to migrate human annotations.
        7. Overwrite Sheet 'PLM' rows 2+ with new data + inherited annotations.
        8. Populate formula network for Col R (=IF(C{r}="","",C{r})) and Col S (=IF(E{r}="","",E{r})).
        9. Clear any old rows beyond new data length to avoid ghost data.
        10. Save workbook.
        11. Move superseded raw files to capnhat\\old\\ if requested.
        12. Refresh Pivot Tables via win32com (with graceful headless fallback).
        """
        wb_path = Path(workbook_path)
        if not wb_path.exists():
            raise FileNotFoundError(f"Workbook not found: {wb_path}")

        # Normalize new PLM data
        if isinstance(new_plm_data, (str, Path)):
            if source_raw_file is None:
                source_raw_file = Path(new_plm_data)
            try:
                from src.core.tree_parser import parse_plm_excel
                tree = parse_plm_excel(new_plm_data)
                df_new_plm = tree.to_dataframe()
            except Exception as parse_err:
                logger.warning("parse_plm_excel failed (%s), falling back to pd.read_excel", parse_err)
                df_new_plm = pd.read_excel(new_plm_data)
        elif hasattr(new_plm_data, "to_dataframe"):
            df_new_plm = new_plm_data.to_dataframe()
        elif isinstance(new_plm_data, pd.DataFrame):
            df_new_plm = new_plm_data.copy()
        elif isinstance(new_plm_data, list):
            df_new_plm = pd.DataFrame(new_plm_data)
        else:
            df_new_plm = pd.DataFrame()

        # Read-only pass: openpyxl is used only to READ (never saved) on the primary path.
        wb = openpyxl.load_workbook(wb_path)
        try:
            if "PLM" not in wb.sheetnames:
                raise ValueError(f"Sheet 'PLM' does not exist in target workbook: {wb_path}")
            df_old = self.extract_existing_plm_records(wb["PLM"])
            archive_sheet_name = self.determine_next_archive_sheet_name(wb.sheetnames, "PLM_old")
        finally:
            wb.close()

        df_inherited = self.match_index_mix(df_new_plm, df_old, key_column="item_id")

        rows, inherited_count, new_unannotated_count = self._build_plm_rows(df_inherited)

        target_archive_dir = Path(archive_dir) if archive_dir else (wb_path.parent / "capnhat" / "old")
        backup_workbook(wb_path, target_archive_dir)

        written_by_excel = excel_com_update_sheet(
            wb_path, "PLM", archive_sheet_name, rows, n_cols=19
        )
        if not written_by_excel:
            self._write_rows_openpyxl(wb_path, "PLM", archive_sheet_name, rows, n_cols=19)

        # File Archiving to capnhat\old\
        archived_files: list[Path] = []
        target_archive_dir.mkdir(parents=True, exist_ok=True)

        if source_raw_file:
            src_f = Path(source_raw_file)
            if src_f.exists() and src_f.resolve() != wb_path.resolve():
                archived_path = self.archive_superseded_file(src_f, target_archive_dir)
                archived_files.append(archived_path)

        if not written_by_excel:
            self.refresh_workbook_pivots(wb_path)

        return InheritanceSummary(
            total_new_parts=len(df_inherited),
            inherited_count=inherited_count,
            new_unannotated_count=new_unannotated_count,
            archived_sheet_name=archive_sheet_name,
            archived_files=archived_files,
            success=True,
            message=f"Migrated {inherited_count} annotations. Archived to '{archive_sheet_name}'.",
        )

    def update_r3_in_workbook(
        self,
        workbook_path: str | Path,
        new_r3_data: pd.DataFrame | str | Path | list[dict[str, Any]] | Any,
        archive_dir: str | Path | None = None,
        source_raw_file: str | Path | None = None,
    ) -> InheritanceSummary:
        """Perform full multi-version update on Sheet 'R3' of form_ssbom workbook.

        Steps:
        1. Open workbook, access or create Sheet 'R3'.
        2. Determine archival sheet name: 'R3_old' or 'R3_old_<N>'.
        3. Copy current Sheet 'R3' to archival sheet preserving cells and formulas.
        4. Normalize new_r3_data (parse file via ResilientR3Parser or DataFrame).
        5. Clear Sheet 'R3' data rows (rows 2+) and write new R3 records.
        6. Set helper columns T (Material) and U (Quantity) for PivotTable compatibility.
        7. Save workbook.
        8. Move superseded raw files to capnhat\\old\\ if requested.
        9. Refresh Pivot Tables via win32com (with graceful headless fallback).
        """
        wb_path = Path(workbook_path)
        if not wb_path.exists():
            raise FileNotFoundError(f"Workbook not found: {wb_path}")

        # Resolve raw file and DataFrame
        if isinstance(new_r3_data, (str, Path)):
            if source_raw_file is None:
                source_raw_file = Path(new_r3_data)
            try:
                from src.automation.sap.parser import ResilientR3Parser
                r3_parser = ResilientR3Parser()
                df_new = r3_parser.parse(new_r3_data)
            except Exception as e:
                logger.warning("ResilientR3Parser failed (%s), falling back to pd.read_excel/pd.read_csv", e)
                try:
                    df_new = pd.read_excel(new_r3_data)
                except Exception:
                    df_new = pd.read_csv(new_r3_data)
        elif isinstance(new_r3_data, pd.DataFrame):
            df_new = new_r3_data.copy()
        elif isinstance(new_r3_data, list):
            df_new = pd.DataFrame(new_r3_data)
        else:
            df_new = pd.DataFrame()

        wb = openpyxl.load_workbook(wb_path, read_only=True)
        try:
            archive_sheet_name = self.determine_next_archive_sheet_name(wb.sheetnames, "R3_old")
        finally:
            wb.close()

        rows = self._build_r3_rows(df_new)

        target_archive_dir = Path(archive_dir) if archive_dir else (wb_path.parent / "capnhat" / "old")
        backup_workbook(wb_path, target_archive_dir)

        written_by_excel = excel_com_update_sheet(
            wb_path, "R3", archive_sheet_name, rows, n_cols=21, headers=self.R3_HEADERS
        )
        if not written_by_excel:
            self._write_rows_openpyxl(
                wb_path, "R3", archive_sheet_name, rows, n_cols=21, headers=self.R3_HEADERS
            )

        # Archive superseded raw file to capnhat\old\
        archived_files: list[Path] = []
        target_archive_dir.mkdir(parents=True, exist_ok=True)

        if source_raw_file:
            src_f = Path(source_raw_file)
            if src_f.exists() and src_f.resolve() != wb_path.resolve():
                archived_path = self.archive_superseded_file(src_f, target_archive_dir)
                archived_files.append(archived_path)

        if not written_by_excel:
            self.refresh_workbook_pivots(wb_path)

        return InheritanceSummary(
            total_new_parts=len(df_new),
            inherited_count=len(df_new),
            new_unannotated_count=0,
            archived_sheet_name=archive_sheet_name,
            archived_files=archived_files,
            success=True,
            message=f"Đã cập nhật Sheet 'R3' với {len(df_new)} linh kiện. Lưu trữ bản cũ thành '{archive_sheet_name}'.",
        )

    def refresh_workbook_pivots(self, workbook_path: str | Path) -> bool:
        """Trigger Pivot Cache refresh via win32com or Excel automation.

        Safely handles environments without Windows Excel GUI (e.g. headless/CI).
        """
        wb_resolved = Path(workbook_path).resolve()
        if not wb_resolved.exists():
            logger.error("Workbook does not exist: %s", wb_resolved)
            return False
        if os.environ.get("SSBOM_DISABLE_EXCEL_COM") == "1":
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

            excel = win32com.client.DispatchEx("Excel.Application")
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
    new_plm_data: pd.DataFrame | str | Path | Any,
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


def update_plm_in_workbook(
    workbook_path: str | Path,
    new_plm_data: pd.DataFrame | str | Path | Any,
    archive_dir: str | Path | None = None,
    source_raw_file: str | Path | None = None,
) -> InheritanceSummary:
    """Alias for update_workbook_with_inheritance."""
    engine = InheritanceEngine()
    return engine.update_workbook_with_inheritance(
        workbook_path=workbook_path,
        new_plm_data=new_plm_data,
        archive_dir=archive_dir,
        source_raw_file=source_raw_file,
    )


def update_r3_in_workbook(
    workbook_path: str | Path,
    new_r3_data: pd.DataFrame | str | Path | list[dict[str, Any]] | Any,
    archive_dir: str | Path | None = None,
    source_raw_file: str | Path | None = None,
) -> InheritanceSummary:
    """Convenience functional interface for Sheet 'R3' update."""
    engine = InheritanceEngine()
    return engine.update_r3_in_workbook(
        workbook_path=workbook_path,
        new_r3_data=new_r3_data,
        archive_dir=archive_dir,
        source_raw_file=source_raw_file,
    )


def refresh_workbook_pivots(workbook_path: str | Path) -> bool:
    """Convenience functional interface for Pivot Table refresh."""
    engine = InheritanceEngine()
    return engine.refresh_workbook_pivots(workbook_path)


def filter_discrepancies_in_workbook(workbook_path: str | Path) -> dict[str, int]:
    """Apply discrepancy filter on form_ssbom workbook matching legacy locdl_ss.

    Hides 'OK' rows and leaves only 'NG' or '#N/A' discrepancy rows visible:
    - Sheet 'CTTT': Filter Field / Col (Status) != 'OK'
    - Sheet 'PLM': Filter Field / Col (Reverse Check) != 'OK'
    - Sheet 'CTTT_Total': Filter Field / Col (Status) != 'OK'
    """
    wb_path = Path(workbook_path)
    if not wb_path.exists():
        raise FileNotFoundError(f"Workbook not found: {wb_path}")

    wb = load_workbook_preserving(wb_path)
    counts = {"CTTT": 0, "PLM": 0, "CTTT_Total": 0}

    # 1. Sheet CTTT
    if "CTTT" in wb.sheetnames:
        ws = wb["CTTT"]
        status_col = 14
        for c in range(1, min(ws.max_column + 1, 25)):
            val = str(ws.cell(row=1, column=c).value or "").strip().upper()
            if any(k in val for k in ["CHECK", "TRẠNG THÁI", "STATUS"]):
                status_col = c
                break
        for r in range(2, ws.max_row + 1):
            val = str(ws.cell(row=r, column=status_col).value or "").strip().upper()
            if val == "OK":
                ws.row_dimensions[r].hidden = True
            elif val in ("NG", "#N/A", "DIFF") or "NG" in val or "#N/A" in val:
                ws.row_dimensions[r].hidden = False
                counts["CTTT"] += 1

    # 2. Sheet PLM
    if "PLM" in wb.sheetnames:
        ws = wb["PLM"]
        status_col = 14
        for c in range(1, min(ws.max_column + 1, 20)):
            val = str(ws.cell(row=1, column=c).value or "").strip().upper()
            if any(k in val for k in ["CHECK", "REVERSE", "THIẾU"]):
                status_col = c
                break
        for r in range(2, ws.max_row + 1):
            val = str(ws.cell(row=r, column=status_col).value or "").strip().upper()
            if val == "OK":
                ws.row_dimensions[r].hidden = True
            elif "#N/A" in val or "NG" in val or "THIẾU" in val:
                ws.row_dimensions[r].hidden = False
                counts["PLM"] += 1

    # 3. Sheet CTTT_Total
    if "CTTT_Total" in wb.sheetnames:
        ws = wb["CTTT_Total"]
        status_col = 6
        for c in range(1, min(ws.max_column + 1, 15)):
            val = str(ws.cell(row=1, column=c).value or "").strip().upper()
            if any(k in val for k in ["TRẠNG THÁI", "STATUS", "CHECK"]):
                status_col = c
                break
        for r in range(2, ws.max_row + 1):
            val = str(ws.cell(row=r, column=status_col).value or "").strip().upper()
            if val == "OK":
                ws.row_dimensions[r].hidden = True
            elif val in ("NG", "#N/A") or "NG" in val or "#N/A" in val:
                ws.row_dimensions[r].hidden = False
                counts["CTTT_Total"] += 1

    wb.save(wb_path)
    wb.close()
    return counts


def unhide_all_rows_in_workbook(workbook_path: str | Path) -> None:
    """Unhide all rows across all sheets in workbook."""
    wb_path = Path(workbook_path)
    if not wb_path.exists():
        return
    wb = load_workbook_preserving(wb_path)
    for ws in wb.worksheets:
        for r in range(1, ws.max_row + 1):
            ws.row_dimensions[r].hidden = False
    wb.save(wb_path)
    wb.close()
