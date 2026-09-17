"""Core Cross-Reconciliation Engine for Kyocera BOM Comparison.

Modern Python modernization of legacy Excel VBA comparison macros:
- F6: Line-by-line three-way reconciliation (CTTT vs PLM vs R3).
- F7: Missing parts detection (reverse lookup on PLM identifying omitted parts in CTTT).
- F8: Cross-station total aggregation (CTTT_Total) across sub-units vs R3 grand totals.
- F9: Annotation migration engine preserving member explanations across revisions.
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass
from typing import Any

import pandas as pd

from src.core.models import BOMTree


@dataclass
class ReconciliationResult:
    """Consolidated three-way reconciliation result interface contract.

    Attributes:
        cttt_rows: Line-by-line reconciliation DataFrame (Sheet CTTT).
        plm_missing_rows: Reverse-lookup missing parts DataFrame (Sheet PLM).
        cttt_totals: Cross-station aggregated totals DataFrame (Sheet CTTT_Total).
        msi_results: MSI serialized component evaluation DataFrame (Sheet MSI_7980_7990).
        overall_status: Overall machine check status ('OK' or 'NG').
    """

    cttt_rows: pd.DataFrame
    plm_missing_rows: pd.DataFrame
    cttt_totals: pd.DataFrame
    msi_results: pd.DataFrame
    overall_status: str  # "OK" or "NG"


def _clean_part_code(val: Any) -> str:
    """Normalize part code by converting to string, stripping whitespace, and uppercasing."""
    if val is None or pd.isna(val):
        return ""
    return str(val).strip().upper()


def _to_float(val: Any) -> float:
    """Safely convert value to float, defaulting to 0.0 if invalid or NaN."""
    if val is None or pd.isna(val):
        return 0.0
    if isinstance(val, str):
        val = val.strip().replace(",", ".")
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0


def _normalize_rev(val: Any) -> str:
    """Normalize revision string, stripping whitespace, prefixes, and leading zeros."""
    if val is None or pd.isna(val):
        return ""
    if isinstance(val, (int, float)):
        if isinstance(val, float) and not math.isfinite(val):
            return ""
        if isinstance(val, float) and val.is_integer():
            val = int(val)
        s = str(val).strip()
    else:
        s = str(val).strip()
    if not s or s.lower() == "nan":
        return ""
    s = s.rstrip(";").strip()
    s = re.sub(r"^(?:rev\.?|\/)\s*", "", s, flags=re.IGNORECASE).strip()
    if s.isdigit():
        try:
            s = str(int(s))
        except ValueError:
            pass
    return s.upper()


def reconcile_single_row(
    cttt_qty: float,
    plm_qty: float,
    r3_qty: float,
    plm_rev: str,
    r3_rev: str,
    tolerance: float = 1e-6,
) -> dict[str, str]:
    """Reconcile a single line item between CTTT, PLM, and R3.

    Authoritative reference specification from form_ssbom.xlsm Sheet CTTT:
    - Col H: Compare (PLM Qty) = IF(G3=E3, 'OK', 'NG')
    - Col L: Compare (R3 Qty) = IF(K3=E3, 'OK', 'NG')
    - Col N: Compare (Rev) = IF(M3=I3, 'OK', 'NG')
    - Col R: Check = IF(OR(G3=0, K3=0, N3='NG', H3='NG', L3='NG'), 'NG', 'OK')

    Args:
        cttt_qty: Quantity specified in CTTT work instruction.
        plm_qty: Quantity found in PLM engineering BOM.
        r3_qty: Quantity found in SAP R3 ERP BOM.
        plm_rev: Engineering revision in PLM.
        r3_rev: Engineering revision in SAP R3.
        tolerance: Float comparison tolerance.

    Returns:
        dict with comp_plm_qty, comp_r3_qty, comp_rev, overall_check, and status.
    """
    if math.isnan(cttt_qty) or math.isnan(plm_qty) or math.isnan(r3_qty):
        raise ValueError("Quantity cannot be NaN")
    if math.isinf(cttt_qty) or math.isinf(plm_qty) or math.isinf(r3_qty):
        raise ValueError("Quantity cannot be Infinity")

    comp_plm_qty = "OK" if abs(cttt_qty - plm_qty) < tolerance else "NG"
    comp_r3_qty = "OK" if abs(cttt_qty - r3_qty) < tolerance else "NG"

    norm_plm_rev = _normalize_rev(plm_rev)
    norm_r3_rev = _normalize_rev(r3_rev)
    comp_rev = "OK" if norm_plm_rev == norm_r3_rev else "NG"

    is_ng = (
        plm_qty == 0
        or r3_qty == 0
        or comp_rev == "NG"
        or comp_plm_qty == "NG"
        or comp_r3_qty == "NG"
    )
    overall_check = "NG" if is_ng else "OK"

    return {
        "comp_plm_qty": comp_plm_qty,
        "comp_r3_qty": comp_r3_qty,
        "comp_rev": comp_rev,
        "cttt_vs_plm_check": comp_plm_qty,
        "cttt_vs_r3_check": comp_r3_qty,
        "overall_check": overall_check,
        "status": overall_check,
    }


def _extract_plm_summary(plm_data: Any) -> pd.DataFrame:
    """Summarize PLM data by part code, summing quantities and taking revision.

    Supports BOMTree, pd.DataFrame, or list[dict].
    Returns DataFrame with columns: ['part_code', 'plm_qty', 'plm_rev'].
    """
    if isinstance(plm_data, BOMTree):
        df_plm = plm_data.to_dataframe()
    elif isinstance(plm_data, pd.DataFrame):
        df_plm = plm_data.copy()
    elif isinstance(plm_data, list):
        df_plm = pd.DataFrame(plm_data)
    else:
        return pd.DataFrame(columns=["part_code", "plm_qty", "plm_rev"])

    if df_plm.empty:
        return pd.DataFrame(columns=["part_code", "plm_qty", "plm_rev"])

    # Resolve part code column
    code_col = None
    for candidate in ["item_id", "MÃ LINH KIỆN", "part_code", "Part Code", "Part Number"]:
        if candidate in df_plm.columns:
            code_col = candidate
            break
    if code_col is None:
        code_col = df_plm.columns[0]

    # Resolve quantity column
    qty_col = None
    for candidate in ["quantity", "SỐ LƯỢNG", "qty", "plm_qty", "Quantity"]:
        if candidate in df_plm.columns:
            qty_col = candidate
            break

    # Resolve revision column
    rev_col = None
    for candidate in ["revision", "Rev PLM", "rev", "item_rev", "Revision"]:
        if candidate in df_plm.columns:
            rev_col = candidate
            break

    df_plm["_norm_code"] = df_plm[code_col].apply(_clean_part_code)
    df_plm = df_plm[df_plm["_norm_code"] != ""]

    if qty_col:
        df_plm["_clean_qty"] = df_plm[qty_col].apply(_to_float)
    else:
        df_plm["_clean_qty"] = 1.0

    if rev_col:
        df_plm["_clean_rev"] = df_plm[rev_col].apply(_normalize_rev)
    else:
        df_plm["_clean_rev"] = ""

    # Group by normalized code
    summary = df_plm.groupby("_norm_code", as_index=False).agg(
        plm_qty=("_clean_qty", "sum"),
        plm_rev=("_clean_rev", "first"),
    )
    summary.rename(columns={"_norm_code": "part_code"}, inplace=True)
    return summary


def _extract_r3_summary(r3_data: Any) -> pd.DataFrame:
    """Summarize SAP R3 data by part code, summing quantities and taking revision.

    Supports pd.DataFrame or list[dict].
    Returns DataFrame with columns: ['part_code', 'r3_qty', 'r3_rev'].
    """
    if isinstance(r3_data, pd.DataFrame):
        df_r3 = r3_data.copy()
    elif isinstance(r3_data, list):
        df_r3 = pd.DataFrame(r3_data)
    else:
        return pd.DataFrame(columns=["part_code", "r3_qty", "r3_rev"])

    if df_r3.empty:
        return pd.DataFrame(columns=["part_code", "r3_qty", "r3_rev"])

    # Resolve part code column
    code_col = None
    for candidate in ["part_code", "MÃ LINH KIỆN", "item_id", "material", "Material", "Part Code"]:
        if candidate in df_r3.columns:
            code_col = candidate
            break
    if code_col is None:
        code_col = df_r3.columns[0]

    # Resolve quantity column
    qty_col = None
    for candidate in ["r3_total_qty", "quantity", "SỐ LƯỢNG", "qty", "r3_qty", "Quantity"]:
        if candidate in df_r3.columns:
            qty_col = candidate
            break

    # Resolve revision column
    rev_col = None
    for candidate in ["revision", "rev_r3", "Rev R3", "rev", "Revision"]:
        if candidate in df_r3.columns:
            rev_col = candidate
            break

    df_r3["_norm_code"] = df_r3[code_col].apply(_clean_part_code)
    df_r3 = df_r3[df_r3["_norm_code"] != ""]

    if qty_col:
        df_r3["_clean_qty"] = df_r3[qty_col].apply(_to_float)
    else:
        df_r3["_clean_qty"] = 1.0

    if rev_col:
        df_r3["_clean_rev"] = df_r3[rev_col].apply(_normalize_rev)
    else:
        df_r3["_clean_rev"] = ""

    summary = df_r3.groupby("_norm_code", as_index=False).agg(
        r3_qty=("_clean_qty", "sum"),
        r3_rev=("_clean_rev", "first"),
    )
    summary.rename(columns={"_norm_code": "part_code"}, inplace=True)
    return summary


def reconcile_three_way(
    cttt_data: pd.DataFrame | list[dict[str, Any]],
    plm_data: BOMTree | pd.DataFrame | list[dict[str, Any]],
    r3_data: pd.DataFrame | list[dict[str, Any]],
    tolerance: float = 1e-6,
) -> pd.DataFrame:
    """Feature F6: Line-by-line three-way reconciliation (CTTT vs PLM vs R3).

    Generates standardized comparison rows corresponding to Sheet CTTT in form_ssbom.xlsm.
    Evaluates quantity comparisons, revision comparisons, and overall row check.

    Args:
        cttt_data: Member work instruction data (CTTT).
        plm_data: Engineering BOM data (BOMTree or DataFrame).
        r3_data: SAP CS12 ERP BOM data.
        tolerance: Float comparison tolerance.

    Returns:
        pd.DataFrame containing standardized comparison matrix.
    """
    if isinstance(cttt_data, pd.DataFrame):
        df_cttt = cttt_data.copy()
    elif isinstance(cttt_data, list):
        df_cttt = pd.DataFrame(cttt_data)
    else:
        df_cttt = pd.DataFrame()

    standard_cols = [
        "SUB",
        "TRANG CTTT",
        "MÃ LINH KIỆN",
        "TÊN LINH KIỆN",
        "SỐ LƯỢNG",
        "PHỤ TRÁCH",
        "Q.ty (PLM)",
        "Compare (PLM Qty)",
        "Rev PLM",
        "Qty (R3)",
        "Compare (R3 Qty)",
        "Rev R3",
        "Compare (Rev)",
        "Giải thích",
        "Check",
    ]

    if df_cttt.empty:
        return pd.DataFrame(columns=standard_cols)

    # Resolve CTTT columns
    sub_col = next((c for c in ["SUB", "unit_name", "sub", "Unit"] if c in df_cttt.columns), None)
    page_col = next((c for c in ["TRANG CTTT", "page", "page_no", "Trang CTTT"] if c in df_cttt.columns), None)
    code_col = next((c for c in ["MÃ LINH KIỆN", "part_code", "item_id", "Part Code"] if c in df_cttt.columns), None)
    name_col = next((c for c in ["TÊN LINH KIỆN", "item_name", "part_name", "Part Name"] if c in df_cttt.columns), None)
    qty_col = next((c for c in ["SỐ LƯỢNG", "quantity", "qty", "cttt_qty"] if c in df_cttt.columns), None)
    pic_col = next((c for c in ["PHỤ TRÁCH", "phu_trach", "person_in_charge", "Person"] if c in df_cttt.columns), None)
    exp_col = next((c for c in ["Giải thích", "giai_thich", "explanation", "note"] if c in df_cttt.columns), None)

    if code_col is None:
        code_col = df_cttt.columns[0]

    # Build summaries for PLM and R3
    plm_summary = _extract_plm_summary(plm_data)
    r3_summary = _extract_r3_summary(r3_data)

    plm_map = plm_summary.set_index("part_code").to_dict(orient="index") if not plm_summary.empty else {}
    r3_map = r3_summary.set_index("part_code").to_dict(orient="index") if not r3_summary.empty else {}

    results = []
    for _, row in df_cttt.iterrows():
        raw_code = row[code_col] if code_col in row else ""
        norm_code = _clean_part_code(raw_code)

        cttt_qty = _to_float(row[qty_col]) if qty_col and qty_col in row else 0.0
        sub_val = str(row[sub_col]).strip() if sub_col and sub_col in row and pd.notna(row[sub_col]) else ""
        page_val = str(row[page_col]).strip() if page_col and page_col in row and pd.notna(row[page_col]) else ""
        name_val = str(row[name_col]).strip() if name_col and name_col in row and pd.notna(row[name_col]) else ""
        pic_val = str(row[pic_col]).strip() if pic_col and pic_col in row and pd.notna(row[pic_col]) else ""
        exp_val = str(row[exp_col]).strip() if exp_col and exp_col in row and pd.notna(row[exp_col]) else ""

        # Lookup in PLM
        plm_entry = plm_map.get(norm_code, {})
        plm_qty = plm_entry.get("plm_qty", 0.0)
        plm_rev = plm_entry.get("plm_rev", "")

        # Lookup in R3
        r3_entry = r3_map.get(norm_code, {})
        r3_qty = r3_entry.get("r3_qty", 0.0)
        r3_rev = r3_entry.get("r3_rev", "")

        rec = reconcile_single_row(
            cttt_qty=cttt_qty,
            plm_qty=plm_qty,
            r3_qty=r3_qty,
            plm_rev=plm_rev,
            r3_rev=r3_rev,
            tolerance=tolerance,
        )

        results.append({
            "SUB": sub_val,
            "TRANG CTTT": page_val,
            "MÃ LINH KIỆN": str(raw_code).strip(),
            "TÊN LINH KIỆN": name_val,
            "SỐ LƯỢNG": cttt_qty,
            "PHỤ TRÁCH": pic_val,
            "Q.ty (PLM)": plm_qty,
            "Compare (PLM Qty)": rec["comp_plm_qty"],
            "Rev PLM": plm_rev,
            "Qty (R3)": r3_qty,
            "Compare (R3 Qty)": rec["comp_r3_qty"],
            "Rev R3": r3_rev,
            "Compare (Rev)": rec["comp_rev"],
            "Giải thích": exp_val,
            "Check": rec["overall_check"],
            # Programmatic aliases for convenience
            "comp_plm_qty": rec["comp_plm_qty"],
            "comp_r3_qty": rec["comp_r3_qty"],
            "comp_rev": rec["comp_rev"],
            "overall_check": rec["overall_check"],
            "status": rec["overall_check"],
        })

    return pd.DataFrame(results)


def detect_missing_parts(
    plm_parts: BOMTree | pd.DataFrame | list[dict[str, Any]],
    cttt_parts: pd.DataFrame | list[dict[str, Any]],
    only_leaves: bool = False,
) -> pd.DataFrame:
    """Feature F7: Missing Parts Detection (Reverse verification on Sheet PLM).

    Identifies parts engineered in PLM that assembly line members omitted from CTTT.
    Equivalent to legacy macro `ws_Plm.Range("A1:S1").AutoFilter Field:=14, Criteria1:="#N/A"`.

    Args:
        plm_parts: PLM BOM data (BOMTree, DataFrame, or list of dicts).
        cttt_parts: CTTT work instruction DataFrame.
        only_leaves: When True, filters out parent assemblies that have children.

    Returns:
        pd.DataFrame containing all PLM rows that are absent in CTTT.
    """
    if isinstance(plm_parts, BOMTree):
        df_plm = plm_parts.to_dataframe()
    elif isinstance(plm_parts, pd.DataFrame):
        df_plm = plm_parts.copy()
    elif isinstance(plm_parts, list):
        df_plm = pd.DataFrame(plm_parts)
    else:
        df_plm = pd.DataFrame()

    if isinstance(cttt_parts, pd.DataFrame):
        df_cttt = cttt_parts.copy()
    elif isinstance(cttt_parts, list):
        df_cttt = pd.DataFrame(cttt_parts)
    else:
        df_cttt = pd.DataFrame()

    if df_plm.empty:
        return df_plm

    # Resolve PLM part code column
    plm_code_col = next((c for c in ["item_id", "MÃ LINH KIỆN", "part_code", "Part Code"] if c in df_plm.columns), None)
    if plm_code_col is None:
        plm_code_col = df_plm.columns[0]

    # Resolve CTTT part code column
    cttt_codes: set[str] = set()
    if not df_cttt.empty:
        cttt_code_col = next((c for c in ["MÃ LINH KIỆN", "part_code", "item_id", "Part Code"] if c in df_cttt.columns), None)
        if cttt_code_col is None:
            cttt_code_col = df_cttt.columns[0]
        cttt_codes = set(df_cttt[cttt_code_col].apply(_clean_part_code))
        cttt_codes.discard("")

    # Optionally filter to leaf parts only
    filtered_plm = df_plm
    if only_leaves and "has_children" in df_plm.columns:
        # has_children can be boolean or string ('True'/'False')
        is_leaf_mask = df_plm["has_children"].apply(
            lambda x: False if x in (True, 1, "True", "true", "1") else True
        )
        filtered_plm = df_plm[is_leaf_mask]

    # Missing check
    norm_plm_codes = filtered_plm[plm_code_col].apply(_clean_part_code)
    missing_mask = ~norm_plm_codes.isin(cttt_codes)

    return filtered_plm[missing_mask].copy()


def aggregate_cross_station(
    cttt_rows: pd.DataFrame | list[dict[str, Any]],
    r3_totals: pd.DataFrame | list[dict[str, Any]],
    tolerance: float = 1e-6,
) -> pd.DataFrame:
    """Feature F8: Cross-Station Total Aggregation (Sheet CTTT_Total).

    Sums shared components across multiple sub-units and reconciles them against
    the machine grand total from SAP R3.

    Args:
        cttt_rows: CTTT DataFrame with part codes and quantities per station.
        r3_totals: SAP R3 total quantity summary per part.
        tolerance: Float comparison tolerance.

    Returns:
        pd.DataFrame containing grouped CTTT_SUM_QTY, r3_total_qty, and STATUS ('OK'/'NG').
    """
    if isinstance(cttt_rows, pd.DataFrame):
        df_cttt = cttt_rows.copy()
    elif isinstance(cttt_rows, list):
        df_cttt = pd.DataFrame(cttt_rows)
    else:
        df_cttt = pd.DataFrame()

    if isinstance(r3_totals, pd.DataFrame):
        df_r3 = r3_totals.copy()
    elif isinstance(r3_totals, list):
        df_r3 = pd.DataFrame(r3_totals)
    else:
        df_r3 = pd.DataFrame()

    standard_cols = ["MÃ LINH KIỆN", "CTTT_SUM_QTY", "part_code", "r3_total_qty", "STATUS"]
    if df_cttt.empty:
        return pd.DataFrame(columns=standard_cols)

    # Resolve CTTT columns
    code_col = next((c for c in ["MÃ LINH KIỆN", "part_code", "item_id", "Part Code"] if c in df_cttt.columns), None)
    qty_col = next((c for c in ["SỐ LƯỢNG", "quantity", "qty", "cttt_qty"] if c in df_cttt.columns), None)
    if code_col is None:
        code_col = df_cttt.columns[0]

    df_cttt["_norm_code"] = df_cttt[code_col].apply(_clean_part_code)
    df_cttt = df_cttt[df_cttt["_norm_code"] != ""]

    if qty_col and qty_col in df_cttt.columns:
        df_cttt["_clean_qty"] = df_cttt[qty_col].apply(_to_float)
    else:
        df_cttt["_clean_qty"] = 1.0

    # Aggregate across sub-units
    grouped = df_cttt.groupby("_norm_code", as_index=False)["_clean_qty"].sum()
    grouped.rename(columns={"_norm_code": "MÃ LINH KIỆN", "_clean_qty": "CTTT_SUM_QTY"}, inplace=True)

    # Process R3 totals
    r3_summary = _extract_r3_summary(df_r3)
    r3_summary.rename(columns={"r3_qty": "r3_total_qty"}, inplace=True)

    merged = pd.merge(
        grouped,
        r3_summary[["part_code", "r3_total_qty"]],
        left_on="MÃ LINH KIỆN",
        right_on="part_code",
        how="left",
    )
    # If part not in R3, default part_code to MÃ LINH KIỆN and r3_total_qty to 0.0
    merged["part_code"] = merged["part_code"].fillna(merged["MÃ LINH KIỆN"])
    merged["r3_total_qty"] = pd.to_numeric(merged["r3_total_qty"], errors="coerce").fillna(0.0)

    merged["STATUS"] = merged.apply(
        lambda r: "OK" if abs(r["CTTT_SUM_QTY"] - r["r3_total_qty"]) < tolerance else "NG",
        axis=1,
    )
    # Convenience programmatic alias
    merged["status"] = merged["STATUS"]
    return merged


def migrate_annotations(
    df_new_plm: pd.DataFrame,
    df_old_plm: pd.DataFrame,
    key_column: str | None = None,
) -> pd.DataFrame:
    """Feature F9: Annotation Migration Engine (ham_match_index_mix).

    Preserves and migrates member explanations ('Giải thích'), person in charge
    ('Phụ trách'), and manager verification ('Quản lý check') from previous PLM
    sheets to new PLM exports. Newly introduced parts receive clean empty strings.

    Args:
        df_new_plm: Newly imported PLM DataFrame.
        df_old_plm: Previous revision PLM DataFrame containing user annotations.
        key_column: Optional explicit key column name for matching.

    Returns:
        pd.DataFrame: df_new_plm with annotations migrated.
    """
    if df_new_plm.empty:
        return df_new_plm.copy()
    if df_old_plm.empty:
        df_res = df_new_plm.copy()
        for col in ["giai_thich", "phu_trach", "quan_ly_check"]:
            if col not in df_res.columns:
                df_res[col] = ""
        return df_res

    # Determine key columns independently for new and old sheets
    new_key = key_column
    if new_key is None or new_key not in df_new_plm.columns:
        for candidate in ["part_code", "item_id", "MÃ LINH KIỆN", "Part Code"]:
            if candidate in df_new_plm.columns:
                new_key = candidate
                break
    if new_key is None:
        new_key = df_new_plm.columns[0]

    old_key = key_column
    if old_key is None or old_key not in df_old_plm.columns:
        for candidate in ["part_code", "item_id", "MÃ LINH KIỆN", "Part Code"]:
            if candidate in df_old_plm.columns:
                old_key = candidate
                break
    if old_key is None:
        old_key = df_old_plm.columns[0]

    # Resolve annotation columns in old sheet
    exp_col = next((c for c in ["giai_thich", "Giải thích", "explanation"] if c in df_old_plm.columns), None)
    pic_col = next((c for c in ["phu_trach", "Phụ trách", "person_in_charge"] if c in df_old_plm.columns), None)
    chk_col = next((c for c in ["quan_ly_check", "Quản lý check", "manager_check"] if c in df_old_plm.columns), None)

    old_copy = df_old_plm.copy()
    old_copy["_clean_key"] = old_copy[old_key].apply(_clean_part_code)
    old_copy = old_copy[old_copy["_clean_key"] != ""]

    old_copy["_gt"] = old_copy[exp_col].fillna("") if exp_col else ""
    old_copy["_pt"] = old_copy[pic_col].fillna("") if pic_col else ""
    old_copy["_ql"] = old_copy[chk_col].fillna("") if chk_col else ""

    # Deduplicate old annotations by key
    old_annotations = old_copy[["_clean_key", "_gt", "_pt", "_ql"]].drop_duplicates(subset=["_clean_key"])

    new_copy = df_new_plm.copy()
    new_copy["_clean_key"] = new_copy[new_key].apply(_clean_part_code)

    migrated = pd.merge(
        new_copy,
        old_annotations,
        on="_clean_key",
        how="left",
    )

    # Assign to standard columns
    migrated["giai_thich"] = migrated["_gt"].fillna("").astype(str).replace("nan", "")
    migrated["phu_trach"] = migrated["_pt"].fillna("").astype(str).replace("nan", "")
    migrated["quan_ly_check"] = migrated["_ql"].fillna("").astype(str).replace("nan", "")

    # Also assign to Vietnamese headers if present or standard
    migrated["Giải thích"] = migrated["giai_thich"]
    migrated["Phụ trách"] = migrated["phu_trach"]
    migrated["Quản lý check"] = migrated["quan_ly_check"]

    # Drop temporary join columns
    migrated.drop(columns=["_clean_key", "_gt", "_pt", "_ql"], inplace=True, errors="ignore")
    return migrated


class ReconciliationEngine:
    """Consolidated engine executing all BOM reconciliation workflows (F6..F9)."""

    def __init__(self, tolerance: float = 1e-6) -> None:
        self.tolerance = tolerance

    def reconcile_single_row(
        self,
        cttt_qty: float,
        plm_qty: float,
        r3_qty: float,
        plm_rev: str,
        r3_rev: str,
    ) -> dict[str, str]:
        """Reconcile a single line item."""
        return reconcile_single_row(
            cttt_qty=cttt_qty,
            plm_qty=plm_qty,
            r3_qty=r3_qty,
            plm_rev=plm_rev,
            r3_rev=r3_rev,
            tolerance=self.tolerance,
        )

    def reconcile_three_way(
        self,
        cttt_data: pd.DataFrame | list[dict[str, Any]],
        plm_data: BOMTree | pd.DataFrame | list[dict[str, Any]],
        r3_data: pd.DataFrame | list[dict[str, Any]],
    ) -> pd.DataFrame:
        """Execute line-by-line three-way reconciliation (F6)."""
        return reconcile_three_way(
            cttt_data=cttt_data,
            plm_data=plm_data,
            r3_data=r3_data,
            tolerance=self.tolerance,
        )

    def detect_missing_parts(
        self,
        plm_parts: BOMTree | pd.DataFrame | list[dict[str, Any]],
        cttt_parts: pd.DataFrame | list[dict[str, Any]],
        only_leaves: bool = False,
    ) -> pd.DataFrame:
        """Detect engineered parts omitted from CTTT (F7)."""
        return detect_missing_parts(
            plm_parts=plm_parts,
            cttt_parts=cttt_parts,
            only_leaves=only_leaves,
        )

    def aggregate_cross_station(
        self,
        cttt_rows: pd.DataFrame | list[dict[str, Any]],
        r3_totals: pd.DataFrame | list[dict[str, Any]],
    ) -> pd.DataFrame:
        """Aggregate cross-station components and reconcile against R3 grand totals (F8)."""
        return aggregate_cross_station(
            cttt_rows=cttt_rows,
            r3_totals=r3_totals,
            tolerance=self.tolerance,
        )

    def migrate_annotations(
        self,
        df_new_plm: pd.DataFrame,
        df_old_plm: pd.DataFrame,
        key_column: str | None = None,
    ) -> pd.DataFrame:
        """Migrate member explanations across revisions (F9)."""
        return migrate_annotations(
            df_new_plm=df_new_plm,
            df_old_plm=df_old_plm,
            key_column=key_column,
        )

    def run_full_reconciliation(
        self,
        cttt_data: pd.DataFrame | list[dict[str, Any]],
        plm_data: BOMTree | pd.DataFrame | list[dict[str, Any]],
        r3_data: pd.DataFrame | list[dict[str, Any]],
        msi_results: pd.DataFrame | None = None,
    ) -> ReconciliationResult:
        """Execute complete reconciliation pipeline producing ReconciliationResult.

        Evaluates:
        1. Line-by-line three-way reconciliation (cttt_rows).
        2. Reverse-lookup missing parts detection (plm_missing_rows).
        3. Cross-station total aggregation (cttt_totals).
        4. Overall machine check status ('OK' if zero 'NG' in all tables, else 'NG').
        """
        cttt_rows = self.reconcile_three_way(cttt_data, plm_data, r3_data)
        plm_missing_rows = self.detect_missing_parts(plm_data, cttt_data)
        cttt_totals = self.aggregate_cross_station(cttt_data, r3_data)

        if msi_results is None:
            msi_results = pd.DataFrame()

        # Determine overall status
        has_ng = False
        if not cttt_rows.empty and (cttt_rows["Check"] == "NG").any():
            has_ng = True
        if not plm_missing_rows.empty:
            has_ng = True
        if not cttt_totals.empty and (cttt_totals["STATUS"] == "NG").any():
            has_ng = True
        if not msi_results.empty:
            status_col = "KẾT QUẢ" if "KẾT QUẢ" in msi_results.columns else "status"
            if status_col in msi_results.columns and (msi_results[status_col] == "NG").any():
                has_ng = True

        overall_status = "NG" if has_ng else "OK"

        return ReconciliationResult(
            cttt_rows=cttt_rows,
            plm_missing_rows=plm_missing_rows,
            cttt_totals=cttt_totals,
            msi_results=msi_results,
            overall_status=overall_status,
        )


# Backward-compatible functional aliases
detect_missing_plm_parts = detect_missing_parts
aggregate_cross_station_totals = aggregate_cross_station
