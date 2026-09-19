"""Feature F10: MSI & Fix Serial Decision Engine.

Evaluates barcode-managed serialized assemblies (Image Unit, Fuser, Laser, Outer Case,
ISU, Hontai) against engineering PLM BOM and FIX_SERIAL_DLTOOL master data.

Implements the authoritative 9-branch decision table from legacy msi.bas (lines 120-379):
- Branch 1: Unit code does NOT exist in PLM BOM -> 'NG' (Col B, K = Red)
- Branch 2: Normalize blank service comment to '-'
- Branch 3: Unit code missing from PLM and code mismatch -> 'NG' (Col B, D, K = Red)
- Branch 4: Unit in PLM, but missing from Fix Serial Master Tool -> 'NG' (Col B, K = Red)
- Branch 5: Both match exactly (Code == Master Code, Service == Master Service) -> 'OK' (Col B, D, E, K = Green)
- Branch 6: Code matches, neither has service (Member == '-', Master == '') -> 'OK' (Col B, D, E, K = Green)
- Branch 7: Code matches, Member == '-', but Master != '' (CTTT missed required service note)
  -> 'OK' with service warning (Col B, D, K = Green, Col E = Red)
- Branch 8: In PLM, but 3-char MSI code mismatch -> 'NG' (Col D, K = Red)
- Branch 9: In PLM, Code matches, but Service comment mismatch -> 'NG' (Col E, K = Red)
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable

import pandas as pd


@dataclass
class MSIEvaluationResult:
    """Evaluation result for a single serialized MSI line item."""

    branch: int
    status: str  # "OK" or "NG"
    service_warning: bool
    reason: str
    member_code: str
    master_code: str
    member_service: str
    master_service: str
    in_plm: bool
    highlight_green: list[str] = field(default_factory=list)
    highlight_red: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        """Convert to dict matching test expectations."""
        return {
            "status": self.status,
            "service_warning": self.service_warning,
            "reason": self.reason,
            "branch": self.branch,
            "member_code": self.member_code,
            "master_code": self.master_code,
            "member_service": self.member_service,
            "master_service": self.master_service,
            "highlight_green": self.highlight_green,
            "highlight_red": self.highlight_red,
        }


def _clean_str(val: Any) -> str:
    """Clean and strip string, returning empty string for None/NaN."""
    if val is None or pd.isna(val):
        return ""
    s = str(val).strip()
    return "" if s.lower() == "nan" else s


def evaluate_msi_branch(
    in_plm: bool,
    member_code: str,
    master_code: str,
    member_service: str = "",
    master_service: str = "",
) -> dict[str, Any]:
    """Evaluate a single MSI row against the authoritative 9-branch decision table.

    Args:
        in_plm: Whether unit code exists in PLM BOM.
        member_code: 3-character fixed code entered by member (Col D).
        master_code: 3-character fixed code looked up from master tool (Col N).
        member_service: Service comment entered by member (Col E).
        master_service: Service comment looked up from master tool (Col O).

    Returns:
        dict with status ('OK'/'NG'), service_warning (bool), reason (str),
        branch (int), highlight_green (list), highlight_red (list).
    """
    m_code = _clean_str(member_code)
    mst_code = _clean_str(master_code)

    # Branch 2: Normalize empty service field to '-'
    raw_member_srv = _clean_str(member_service)
    norm_member_service = "-" if (raw_member_srv == "" or raw_member_srv == "-") else raw_member_srv
    norm_master_service = _clean_str(master_service)

    # Case: Not in PLM BOM
    if not in_plm:
        # Branch 3: Missing from PLM and code mismatch
        if m_code != mst_code:
            return {
                "branch": 3,
                "status": "NG",
                "service_warning": False,
                "reason": "Missing from PLM and code mismatch",
                "highlight_green": [],
                "highlight_red": ["B", "D", "K"],
                "member_service_normalized": norm_member_service,
            }
        # Branch 1: Unit code does not exist in PLM
        return {
            "branch": 1,
            "status": "NG",
            "service_warning": False,
            "reason": "Unit code not in PLM",
            "highlight_green": [],
            "highlight_red": ["B", "K"],
            "member_service_normalized": norm_member_service,
        }

    # Case: In PLM BOM
    # Branch 4: Missing from Fix Serial Master Tool
    if not mst_code:
        return {
            "branch": 4,
            "status": "NG",
            "service_warning": False,
            "reason": "Missing from Fix Serial Tool",
            "highlight_green": [],
            "highlight_red": ["B", "K"],
            "member_service_normalized": norm_member_service,
        }

    # Code matches
    if m_code == mst_code:
        # Branch 5: Both code and service match exactly
        if norm_member_service == norm_master_service:
            return {
                "branch": 5,
                "status": "OK",
                "service_warning": False,
                "reason": "Match",
                "highlight_green": ["B", "D", "E", "K"],
                "highlight_red": [],
                "member_service_normalized": norm_member_service,
            }

        # Branch 6: Both have no service note (member has '-', master has '')
        if norm_member_service == "-" and norm_master_service == "":
            return {
                "branch": 6,
                "status": "OK",
                "service_warning": False,
                "reason": "Match without service",
                "highlight_green": ["B", "D", "E", "K"],
                "highlight_red": [],
                "member_service_normalized": norm_member_service,
            }

        # Branch 7: Code matches, but CTTT has '-' while master specifies service note
        if norm_member_service == "-" and norm_master_service != "":
            return {
                "branch": 7,
                "status": "OK",
                "service_warning": True,
                "reason": "Service note required",
                "highlight_green": ["B", "D", "K"],
                "highlight_red": ["E"],
                "member_service_normalized": norm_member_service,
            }

        # Branch 9: Code matches, but service comment does not match and member != '-'
        return {
            "branch": 9,
            "status": "NG",
            "service_warning": False,
            "reason": "Service mismatch",
            "highlight_green": ["B", "D"],
            "highlight_red": ["E", "K"],
            "member_service_normalized": norm_member_service,
        }

    # Branch 8: In PLM, but 3-character code mismatch
    green_cols = ["B"]
    if norm_member_service == "-" and norm_master_service == "":
        green_cols.append("E")

    return {
        "branch": 8,
        "status": "NG",
        "service_warning": False,
        "reason": "3-char code mismatch",
        "highlight_green": green_cols,
        "highlight_red": ["D", "K"],
        "member_service_normalized": norm_member_service,
    }


class FixSerialMaster:
    """Master dictionary lookup for sub-units and machine codes from FIX_SERIAL_DLTOOL.

    Supports loading from Excel file (`FIX_SERIAL_DLTOOL_VER010.xls` or `.xlsx`)
    or direct initialization with python dictionaries for headless operation.
    """

    def __init__(self) -> None:
        # Key: prefix 9 chars of unit code -> (msi_code, service_comment)
        self.subunit_map: dict[str, tuple[str, str]] = {}
        # Key: prefix 10 chars of machine code -> (msi_code, service_comment)
        self.machine_map: dict[str, tuple[str, str]] = {}

    def add_subunit(self, unit_code: str, msi_code: str, service: str = "") -> None:
        """Add a subunit lookup entry. Indexes both full key and 9-character prefix."""
        clean_full = _clean_str(unit_code).upper()
        if clean_full:
            val = (_clean_str(msi_code), _clean_str(service))
            self.subunit_map[clean_full] = val
            prefix9 = clean_full[:9]
            if prefix9 and prefix9 not in self.subunit_map:
                self.subunit_map[prefix9] = val

    def add_machine(self, machine_code: str, msi_code: str, service: str = "") -> None:
        """Add a machine (Hontai) lookup entry. Indexes both full key and 10-character prefix."""
        clean_full = _clean_str(machine_code).upper()
        if clean_full:
            val = (_clean_str(msi_code), _clean_str(service))
            self.machine_map[clean_full] = val
            prefix10 = clean_full[:10]
            if prefix10 and prefix10 not in self.machine_map:
                self.machine_map[prefix10] = val

    def lookup_subunit(self, unit_code: str) -> tuple[str, str]:
        """Look up 3-character MSI code and SERVICE comment for a sub-unit.

        Matches by prefix of 9 characters or fuzzy substring.
        Returns: (msi_code, service_comment), defaulting to ("", "") if not found.
        """
        clean_full = _clean_str(unit_code).upper()
        if not clean_full:
            return ("", "")

        prefix9 = clean_full[:9]
        if prefix9 in self.subunit_map:
            return self.subunit_map[prefix9]
        if clean_full in self.subunit_map:
            return self.subunit_map[clean_full]

        # Fuzzy substring scan across registered keys - require at least 3 characters
        if len(clean_full) >= 3:
            for k, val in self.subunit_map.items():
                if (len(prefix9) >= 3 and prefix9 in k) or (len(k) >= 3 and k in clean_full):
                    return val

        return ("", "")

    def lookup_machine(self, machine_code: str) -> tuple[str, str]:
        """Look up 3-character MSI code and SERVICE comment for Hontai (main body).

        Matches by prefix of 10 characters or fuzzy substring.
        Returns: (msi_code, service_comment), defaulting to ("", "") if not found.
        """
        clean_full = _clean_str(machine_code).upper()
        if not clean_full:
            return ("", "")

        prefix10 = clean_full[:10]
        if prefix10 in self.machine_map:
            return self.machine_map[prefix10]
        if clean_full in self.machine_map:
            return self.machine_map[clean_full]

        if len(clean_full) >= 3:
            for k, val in self.machine_map.items():
                if (len(prefix10) >= 3 and prefix10 in k) or (len(k) >= 3 and k in clean_full):
                    return val

        return ("", "")


    def load_from_file(self, file_path: str | Path) -> bool:
        """Load lookup tables from FIX_SERIAL_DLTOOL Excel workbook (.xls or .xlsx).

        Reads:
        - Sheet 'UNIT': Col A (Unit Code), Col F (MSI code), Col H (SERVICE)
        - Sheet 'MACHINE': Col A (Machine Code), Col F (MSI code), Col H (SERVICE)
        """
        p = Path(file_path)
        if not p.exists():
            return False

        try:
            excel_data = pd.read_excel(p, sheet_name=None, header=None)
        except Exception:
            return False

        # Parse Sheet UNIT
        unit_sheet_name = next((s for s in excel_data if s.upper() == "UNIT"), None)
        if unit_sheet_name:
            df_unit = excel_data[unit_sheet_name]
            for _, row in df_unit.iterrows():
                # Col 0 (Col A): Unit Code
                # Col 5 (Col F): MSI Code
                # Col 7 (Col H): Service
                if len(row) > 0 and pd.notna(row.iloc[0]):
                    code = str(row.iloc[0]).strip()
                    msi = str(row.iloc[5]).strip() if len(row) > 5 and pd.notna(row.iloc[5]) else ""
                    srv = str(row.iloc[7]).strip() if len(row) > 7 and pd.notna(row.iloc[7]) else ""
                    self.add_subunit(code, msi, srv)

        # Parse Sheet MACHINE
        machine_sheet_name = next((s for s in excel_data if s.upper() == "MACHINE"), None)
        if machine_sheet_name:
            df_mach = excel_data[machine_sheet_name]
            for _, row in df_mach.iterrows():
                if len(row) > 0 and pd.notna(row.iloc[0]):
                    code = str(row.iloc[0]).strip()
                    msi = str(row.iloc[5]).strip() if len(row) > 5 and pd.notna(row.iloc[5]) else ""
                    srv = str(row.iloc[7]).strip() if len(row) > 7 and pd.notna(row.iloc[7]) else ""
                    self.add_machine(code, msi, srv)

        return True


class MSIEngine:
    """MSI & Fix Serial evaluation engine executing 9-branch decision logic."""

    def __init__(self, master: FixSerialMaster | None = None) -> None:
        self.master = master or FixSerialMaster()

    def evaluate_branch(
        self,
        in_plm: bool,
        member_code: str,
        master_code: str,
        member_service: str = "",
        master_service: str = "",
    ) -> dict[str, Any]:
        """Direct 9-branch evaluation method."""
        return evaluate_msi_branch(
            in_plm=in_plm,
            member_code=member_code,
            master_code=master_code,
            member_service=member_service,
            master_service=master_service,
        )

    def evaluate(
        self,
        in_plm: bool,
        member_code: str,
        master_code: str,
        member_service: str = "",
        master_service: str = "",
    ) -> dict[str, Any]:
        """Alias for evaluate_branch."""
        return self.evaluate_branch(
            in_plm=in_plm,
            member_code=member_code,
            master_code=master_code,
            member_service=member_service,
            master_service=master_service,
        )

    def evaluate_row(
        self,
        unit_code: str,
        member_code: str,
        member_service: str,
        plm_part_codes: Iterable[str],
        is_hontai: bool = False,
    ) -> MSIEvaluationResult:
        """Evaluate a single unit code row."""
        clean_unit = _clean_str(unit_code).upper()

        # Check if in PLM
        plm_set = {str(c).strip().upper() for c in plm_part_codes if pd.notna(c) and str(c).strip()}
        # In PLM if exact match or substring in PLM set (matching VBA Criteria1:='*' & code & '*')
        in_plm = False
        if clean_unit:
            if clean_unit in plm_set:
                in_plm = True
            else:
                in_plm = any(p and (clean_unit in p or p in clean_unit) for p in plm_set)


        # Lookup in master
        if is_hontai:
            master_code, master_service = self.master.lookup_machine(clean_unit)
        else:
            master_code, master_service = self.master.lookup_subunit(clean_unit)

        eval_res = evaluate_msi_branch(
            in_plm=in_plm,
            member_code=member_code,
            master_code=master_code,
            member_service=member_service,
            master_service=master_service,
        )

        return MSIEvaluationResult(
            branch=eval_res["branch"],
            status=eval_res["status"],
            service_warning=eval_res["service_warning"],
            reason=eval_res["reason"],
            member_code=member_code,
            master_code=master_code,
            member_service=eval_res["member_service_normalized"],
            master_service=master_service,
            in_plm=in_plm,
            highlight_green=eval_res["highlight_green"],
            highlight_red=eval_res["highlight_red"],
        )

    def evaluate_table(
        self,
        msi_data: pd.DataFrame | list[dict[str, Any]],
        plm_part_codes: Iterable[str],
        hontai_default_code: str = "",
    ) -> pd.DataFrame:
        """Evaluate entire Sheet MSI_7980_7990 table.

        Args:
            msi_data: Input MSI DataFrame or list of row dicts.
            plm_part_codes: Collection of valid PLM part codes.
            hontai_default_code: Fallback machine code for Row 36 if blank (from PLM!C2).

        Returns:
            pd.DataFrame: Enriched comparison table with KẾT QUẢ, GHI CHÚ, and highlight flags.
        """
        if isinstance(msi_data, pd.DataFrame):
            df = msi_data.copy()
        elif isinstance(msi_data, list):
            df = pd.DataFrame(msi_data)
        else:
            df = pd.DataFrame()

        standard_cols = [
            "MÃ LK BARCODE",
            "MÃ UNIT",
            "TÊN UNIT",
            "3 KÝ TỰ MSI",
            "SERVICE",
            "ABS",
            "TRANG CTTT",
            "ĐIỆN ÁP",
            "LOẠI LABEL",
            "TÊN PHỤ TRÁCH",
            "KẾT QUẢ",
            "GHI CHÚ",
            "MÃ TRONG PLM",
            "MSI TRONG MASTER",
            "SERVICE TRONG MASTER",
            "BRANCH",
            "SERVICE_WARNING",
        ]

        if df.empty:
            return pd.DataFrame(columns=standard_cols)

        # Resolve column names
        code_col = next(
            (c for c in ["MÃ UNIT", "unit_code", "Mã UNIT", "part_code", "item_id"] if c in df.columns),
            df.columns[0]
        )
        msi_col = next((c for c in ["3 KÝ TỰ MSI", "msi_code", "3 ký tự MSI", "msi", "code"] if c in df.columns), None)
        srv_col = next((c for c in ["SERVICE", "SEVICE", "service", "service_comment", "comment"] if c in df.columns), None)
        barcode_col = next((c for c in ["MÃ LK BARCODE", "barcode", "barcode_code"] if c in df.columns), None)
        name_col = next((c for c in ["TÊN UNIT", "unit_name", "Tên UNIT", "item_name"] if c in df.columns), None)
        pic_col = next((c for c in ["TÊN PHỤ TRÁCH", "phu_trach", "person_in_charge"] if c in df.columns), None)

        rows_out = []
        num_rows = len(df)

        for idx, row in df.iterrows():
            # Check if this row is Hontai / Main body
            is_hontai = False
            if "is_hontai" in row and pd.notna(row["is_hontai"]) and bool(row["is_hontai"]):
                is_hontai = True
            elif idx == 34 or idx == 35 or idx == (num_rows - 1):
                unit_name_str = str(row[name_col]).lower() if name_col and name_col in row and pd.notna(row[name_col]) else ""
                if "hontai" in unit_name_str or "machine" in unit_name_str or "main" in unit_name_str:
                    is_hontai = True

            raw_unit = str(row[code_col]).strip() if code_col in row and pd.notna(row[code_col]) else ""
            if is_hontai and not raw_unit and hontai_default_code:
                raw_unit = hontai_default_code.strip()

            raw_msi = str(row[msi_col]).strip() if msi_col and msi_col in row and pd.notna(row[msi_col]) else ""
            raw_srv = str(row[srv_col]).strip() if srv_col and srv_col in row and pd.notna(row[srv_col]) else ""

            res = self.evaluate_row(
                unit_code=raw_unit,
                member_code=raw_msi,
                member_service=raw_srv,
                plm_part_codes=plm_part_codes,
                is_hontai=is_hontai,
            )

            barcode_val = ""
            if barcode_col and barcode_col in row and pd.notna(row[barcode_col]):
                barcode_val = str(row[barcode_col]).strip()

            row_dict = {
                "MÃ LK BARCODE": barcode_val,
                "MÃ UNIT": raw_unit,
                "TÊN UNIT": str(row[name_col]).strip() if name_col and name_col in row and pd.notna(row[name_col]) else "",
                "3 KÝ TỰ MSI": raw_msi,
                "SERVICE": res.member_service,
                "ABS": str(row.get("ABS", "")),
                "TRANG CTTT": str(row.get("TRANG CTTT", "")),
                "ĐIỆN ÁP": str(row.get("ĐIỆN ÁP", "")),
                "LOẠI LABEL": str(row.get("LOẠI LABEL", "")),
                "TÊN PHỤ TRÁCH": str(row[pic_col]).strip() if pic_col and pic_col in row and pd.notna(row[pic_col]) else "",
                "KẾT QUẢ": res.status,
                "GHI CHÚ": res.reason,
                "MÃ TRONG PLM": raw_unit if res.in_plm else "",
                "MSI TRONG MASTER": res.master_code,
                "SERVICE TRONG MASTER": res.master_service,
                "BRANCH": res.branch,
                "SERVICE_WARNING": res.service_warning,
                # Programmatic aliases
                "status": res.status,
                "reason": res.reason,
                "branch": res.branch,
                "service_warning": res.service_warning,
            }
            rows_out.append(row_dict)

        return pd.DataFrame(rows_out)
