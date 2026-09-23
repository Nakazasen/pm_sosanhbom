"""Robust PLM Excel BOM Tree Parser.

Parses Siemens Teamcenter Active Workspace (TC14) and legacy Excel BOM exports
(14-column and 13-column schemas) into a structured, in-memory BOMTree hierarchy.
"""

from __future__ import annotations

import logging
import re
from collections.abc import Sequence
from pathlib import Path
from typing import Any

import openpyxl
import pandas as pd

from src.core.models import BOMNode, BOMTree

logger = logging.getLogger(__name__)

# Standard column aliases for resilient header matching
COLUMN_ALIASES: dict[str, list[str]] = {
    "level": [
        "level", "lv", "bomlevel", "bom_level", "cap", "capdo", "cấp", "cấpđộ"
    ],
    "item_type": [
        "itemtype", "item_type", "type", "loai", "loạilinhkiện"
    ],
    "item_id": [
        "itemid", "item_id", "partcode", "part_code", "partnumber", "part_number",
        "partno", "part_no", "malinhkien", "mãlinhkiện", "malk"
    ],
    "has_children": [
        "haschildren", "has_children", "cocon", "cócon", "children", "expandable",
        "assemblyindicator", "assembly_indicator"
    ],
    "quantity": [
        "quantity", "qty", "soluong", "sốlượng", "count"
    ],
    "first_parts": [
        "1stparts", "firstparts", "1st_parts", "first_parts"
    ],
    "second_bom_flag": [
        "2ndbomflag", "secondbomflag", "2nd_bom_flag", "second_bom_flag"
    ],
    "effectivity": [
        "occurrenceeffectivities", "occurrence_effectivities", "effectivities",
        "effectivity", "elementeffectivities", "element_effectivities",
        "chuoihieuluc", "chuỗichiếtlực", "hieuluc", "hiệulực", "validity"
    ],
    "item_revision_projects_list": [
        "itemrevisionprojectslist", "itemrevisionprojectlist", "projectslist", "projects_list"
    ],
    "item_name": [
        "itemname", "item_name", "partname", "part_name", "partstext", "parts_text",
        "parttext", "tenlinhkien", "tênlinhkiện", "tenlk", "description", "mota"
    ],
    "notice_no": [
        "noticeno", "notice_no", "ecn", "ecnno", "ecn_no"
    ],
    "revision": [
        "revision", "rev", "revlev", "phienban", "phiênbản"
    ],
    "item_rev_status": [
        "itemrevstatus", "item_rev_status", "releasestatus", "release_status",
        "revstatus", "status", "trangthai", "trạngthái"
    ],
}


def _normalize_header(header: Any) -> str:
    """Normalize header string by lowercasing and removing non-alphanumeric characters."""
    if header is None:
        return ""
    text = str(header).strip().lower()
    return re.sub(r"[^\w]", "", text, flags=re.UNICODE)


def _parse_bool(val: Any) -> bool:
    """Parse boolean value from varied representations."""
    if isinstance(val, bool):
        return val
    if val is None:
        return False
    norm = str(val).strip().lower()
    return norm in ("true", "1", "yes", "y", "t", "x", "fixed assembly", "assembly")


def _parse_int_level(val: Any) -> int | None:
    """Parse level integer from various formats like '1', 1, 1.0, '.1', '...3'."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        try:
            return int(val)
        except (ValueError, OverflowError):
            return None
    text = str(val).strip()
    if not text:
        return None
    # Strip leading dots often present in BOM level formatting (e.g. '..2' -> 2)
    cleaned = text.lstrip(".").strip()
    try:
        return int(cleaned)
    except ValueError:
        # Match first integer sequence
        match = re.search(r"\b(\d+)\b", text)
        if match:
            return int(match.group(1))
        return None


def _parse_float_qty(val: Any, default: float = 1.0) -> float:
    """Parse component quantity float with fallback."""
    if val is None:
        return default
    if isinstance(val, (int, float)):
        return float(val)
    text = str(val).strip().replace(",", ".")
    try:
        return float(text)
    except ValueError:
        match = re.search(r"[-+]?\d*\.?\d+", text)
        if match:
            try:
                return float(match.group(0))
            except ValueError:
                return default
        return default


def _clean_str(val: Any) -> str:
    """Clean string values, returning empty string for None."""
    if val is None:
        return ""
    return str(val).strip()


class PLMTreeParser:
    """Robust parser for 14-column PLM BOM Excel files and tabular data."""

    def __init__(self, default_qty: float = 1.0) -> None:
        self.default_qty = default_qty

    def detect_column_mapping(self, headers: list[Any]) -> dict[str, int]:
        """Detect column positions from header names."""
        mapping: dict[str, int] = {}
        normalized_headers = [_normalize_header(h) for h in headers]

        # Check if this is a Teamcenter export with 'Parts Text' and/or 'Name'
        has_parts_text = any(nh in ("partstext", "parttext", "parts_text") for nh in normalized_headers)
        has_name = "name" in normalized_headers

        if has_parts_text:
            # In Teamcenter format, 'Parts Text' is always Item Name (description)
            for idx, nh in enumerate(normalized_headers):
                if nh in ("partstext", "parttext", "parts_text") and "item_name" not in mapping:
                    mapping["item_name"] = idx
                    break
            # In Teamcenter format, 'Name' is the Item ID (part code)
            if has_name and "item_id" not in mapping:
                mapping["item_id"] = normalized_headers.index("name")

        for col_name, aliases in COLUMN_ALIASES.items():
            if col_name in mapping:
                continue
            for idx, nh in enumerate(normalized_headers):
                if not nh:
                    continue
                if nh in aliases:
                    mapping[col_name] = idx
                    break

        # If 'item_name' not matched yet and 'name' header exists (generic format without Parts Text)
        if "item_name" not in mapping and not has_parts_text and has_name:
            mapping["item_name"] = normalized_headers.index("name")

        # Fallback to positional mapping only if standard TC14 / TC24 width (at least 13 cols)
        if ("level" not in mapping or "item_id" not in mapping) and len(headers) >= 13:
            mapping = self._get_positional_fallback(len(headers))

        return mapping

    def _get_positional_fallback(self, num_cols: int) -> dict[str, int]:
        """Fallback to positional indexing based on column count:
        
        20+ col TC2412 format:
          0: Level (string)
          1: Level (int)
          2: Item Type
          3: Item Id (Name)
          4: 1st Parts
          5: Quantity
          6: 2nd BOM Flag
          7: Item Name (Parts Text)
          8: Notice No
          9: Revision
          10: Item Rev Status (Release Status)
          14: Has Children (Assembly Indicator)
          16: Effectivity (Element Effectivities)

        14-col TC14 format:
          0: Row Index / Object
          1: Level
          2: Item Type
          3: Item Id
          4: Has Children
          5: Quantity
          6: 1st Parts
          7: 2nd BOM Flag
          8: Occurrence Effectivities
          9: Item Revision Projects List
          10: Item Name
          11: Notice No
          12: Revision
          13: Item Rev Status
          
        13-col format (no row index col):
          0: Level
          1: Item Type
          2: Item Id
          3: Has Children
          4: Quantity
          5: 1st Parts
          6: 2nd BOM Flag
          7: Occurrence Effectivities
          8: Item Revision Projects List
          9: Item Name
          10: Notice No
          11: Revision
          12: Item Rev Status
        """
        if num_cols >= 20:
            return {
                "level": 1,
                "item_type": 2,
                "item_id": 3,
                "first_parts": 4,
                "quantity": 5,
                "second_bom_flag": 6,
                "item_name": 7,
                "notice_no": 8,
                "revision": 9,
                "item_rev_status": 10,
                "has_children": 14,
                "effectivity": 16,
            }
        if num_cols >= 14:
            return {
                "level": 1,
                "item_type": 2,
                "item_id": 3,
                "has_children": 4,
                "quantity": 5,
                "first_parts": 6,
                "second_bom_flag": 7,
                "effectivity": 8,
                "item_revision_projects_list": 9,
                "item_name": 10,
                "notice_no": 11,
                "revision": 12,
                "item_rev_status": 13,
            }
        return {
            "level": 0,
            "item_type": 1,
            "item_id": 2,
            "has_children": 3,
            "quantity": 4,
            "first_parts": 5,
            "second_bom_flag": 6,
            "effectivity": 7,
            "item_revision_projects_list": 8,
            "item_name": 9,
            "notice_no": 10,
            "revision": 11,
            "item_rev_status": 12,
        }

    def parse_excel(
        self,
        file_path: str | Path,
        sheet_name: str | int | None = None
    ) -> BOMTree:
        """Parse BOM from an Excel file (.xlsx, .xlsm, .xls)."""
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"PLM Excel file not found: {path}")
        if path.stat().st_size == 0:
            raise ValueError(f"PLM Excel file is empty (0 bytes): {path}")

        # Use openpyxl with data_only=True to evaluate formulas if any
        wb = openpyxl.load_workbook(str(path), data_only=True, read_only=False)
        try:
            if sheet_name is not None:
                if isinstance(sheet_name, int):
                    ws = wb.worksheets[sheet_name]
                else:
                    ws = wb[sheet_name]
            else:
                ws = wb.active

            all_rows = list(ws.iter_rows(values_only=True))
        finally:
            wb.close()

        if not all_rows:
            return BOMTree(roots=[], source_file=str(path))

        # Scan first 10 rows to locate header row
        header_row_idx = 0
        best_score = -1
        detected_mapping: dict[str, int] = {}

        for r_idx in range(min(10, len(all_rows))):
            row_candidates = all_rows[r_idx]
            mapping = self.detect_column_mapping(row_candidates)
            score = sum(1 for k in ("level", "item_id", "has_children", "quantity", "effectivity", "item_name") if k in mapping)
            if score > best_score:
                best_score = score
                header_row_idx = r_idx
                detected_mapping = mapping

        data_rows = all_rows[header_row_idx + 1:]
        return self._build_tree_from_matrix(data_rows, detected_mapping, source_file=str(path), row_offset=header_row_idx + 2)

    def parse_file(
        self,
        file_path: str | Path,
        sheet_name: str | int | None = None,
    ) -> BOMTree:
        """Parse BOM from file with file path and directory validation."""
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {path}")
        if path.is_dir():
            raise IsADirectoryError(f"Target path is a directory: {path}")
        if path.stat().st_size == 0:
            raise ValueError(f"PLM Excel file is empty (0 bytes): {path}")
        return self.parse_excel(path, sheet_name=sheet_name)

    def parse_dataframe(self, df: pd.DataFrame, source_file: str | None = None) -> BOMTree:
        """Parse BOM hierarchy from a pandas DataFrame."""
        headers = list(df.columns)
        mapping = self.detect_column_mapping(headers)
        data_matrix = df.values.tolist()
        return self._build_tree_from_matrix(data_matrix, mapping, source_file=source_file, row_offset=2)

    def parse_records(self, records: list[dict[str, Any]], source_file: str | None = None) -> BOMTree:
        """Parse BOM hierarchy from a list of dictionary records."""
        if not records:
            return BOMTree(roots=[], source_file=source_file)
        df = pd.DataFrame(records)
        return self.parse_dataframe(df, source_file=source_file)

    def _build_tree_from_matrix(
        self,
        rows: Sequence[Sequence[Any]],
        mapping: dict[str, int],
        source_file: str | None = None,
        row_offset: int = 1
    ) -> BOMTree:
        """Construct multi-level BOM tree using a parent stack traversal."""
        roots: list[BOMNode] = []
        stack: list[BOMNode] = []

        lvl_col = mapping.get("level", 0)
        id_col = mapping.get("item_id", 2)
        has_child_col = mapping.get("has_children", 3)
        qty_col = mapping.get("quantity", 4)
        eff_col = mapping.get("effectivity", 7)
        name_col = mapping.get("item_name", 9)
        rev_col = mapping.get("revision", 11)
        type_col = mapping.get("item_type")
        notice_col = mapping.get("notice_no")
        first_col = mapping.get("first_parts")
        second_col = mapping.get("second_bom_flag")
        status_col = mapping.get("item_rev_status")

        for row_num, row in enumerate(rows, start=row_offset):
            # Check row bounds
            if not row or all(v is None or str(v).strip() == "" for v in row):
                continue

            lvl_val = row[lvl_col] if lvl_col < len(row) else None
            parsed_level = _parse_int_level(lvl_val)
            if parsed_level is None:
                continue

            item_id = _clean_str(row[id_col]) if id_col < len(row) else ""
            item_name = _clean_str(row[name_col]) if name_col < len(row) else ""
            has_children = _parse_bool(row[has_child_col]) if has_child_col < len(row) else False
            qty = _parse_float_qty(row[qty_col], default=self.default_qty) if qty_col < len(row) else self.default_qty
            eff = _clean_str(row[eff_col]) if eff_col < len(row) else ""
            rev = _clean_str(row[rev_col]) if rev_col < len(row) else ""

            node = BOMNode(
                level=parsed_level,
                item_id=item_id,
                item_name=item_name,
                has_children=has_children,
                quantity=qty,
                effectivity=eff,
                revision=rev,
                item_type=_clean_str(row[type_col]) if type_col is not None and type_col < len(row) else "",
                notice_no=_clean_str(row[notice_col]) if notice_col is not None and notice_col < len(row) else "",
                first_parts=_clean_str(row[first_col]) if first_col is not None and first_col < len(row) else "",
                second_bom_flag=_clean_str(row[second_col]) if second_col is not None and second_col < len(row) else "",
                item_rev_status=_clean_str(row[status_col]) if status_col is not None and status_col < len(row) else "",
                row_index=row_num,
                children=[],
            )

            # Stack-based hierarchy assembly:
            # Pop stack while stack top's level is greater than or equal to current node's level
            while stack and stack[-1].level >= node.level:
                stack.pop()

            if stack:
                stack[-1].add_child(node)
            else:
                roots.append(node)

            stack.append(node)

        return BOMTree(roots=roots, source_file=source_file)


def parse_plm_excel(
    file_path: str | Path,
    sheet_name: str | int | None = None
) -> BOMTree:
    """Convenience functional API to parse PLM Excel BOM file."""
    parser = PLMTreeParser()
    return parser.parse_excel(file_path, sheet_name=sheet_name)


# Alias for backward compatibility and test frameworks
BOMTreeParser = PLMTreeParser
