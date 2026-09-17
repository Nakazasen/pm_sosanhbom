"""
Resilient SAP R3 CS12 Spreadsheet Parser.
Dynamically resolves Component, Quantity, Revision (RevLev), and metadata columns
across diverse ALV layouts, HTML-in-XLS, TSV, and native Excel variants without
relying on fragile column shifting or deletions.
"""

import io
import logging
import os
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

import pandas as pd

from html.parser import HTMLParser

from src.automation.sap.models import (
    R3ComponentRow,
    SAPParseError,
)

logger = logging.getLogger(__name__)


class SimpleHTMLTableParser(HTMLParser):
    """Zero-dependency HTML table parser that preserves raw strings without type coercion."""

    def __init__(self):
        super().__init__()
        self.tables: List[List[List[str]]] = []
        self.current_table: List[List[str]] = []
        self.current_row: List[str] = []
        self.current_cell: List[str] = []
        self.in_cell = False

    def handle_starttag(self, tag: str, attrs: Any) -> None:
        tag_lower = tag.lower()
        if tag_lower in ("td", "th"):
            self.in_cell = True
            self.current_cell = []
        elif tag_lower == "tr":
            self.current_row = []
        elif tag_lower == "table":
            self.current_table = []

    def handle_endtag(self, tag: str) -> None:
        tag_lower = tag.lower()
        if tag_lower in ("td", "th"):
            self.in_cell = False
            self.current_row.append("".join(self.current_cell).strip())
        elif tag_lower == "tr":
            if self.current_row:
                self.current_table.append(self.current_row)
        elif tag_lower == "table":
            if self.current_table:
                self.tables.append(self.current_table)

    def handle_data(self, data: str) -> None:
        if self.in_cell:
            self.current_cell.append(data)


# Canonical Header Match Patterns (Regular expressions for case-insensitive matching)
PATTERNS_PART_CODE = [
    r"^PART\s*CODE",
    r"^COMPONENT",
    r"^MATERIAL",
    r"^ITEM\s*CODE",
    r"^MATNR",
    r"^M[ÃA]\s*LINH\s*KI[ỆE]N",
    r"^LINH\s*KI[ỆE]N",
    r"^PART\s*NO",
    r"^PART_NUMBER",
]

PATTERNS_QUANTITY = [
    r"^Q\.?TY",
    r"^QTY",
    r"^QUANTITY",
    r"^MENGE",
    r"^S[ỐO]\s*L[ƯU][ỢO]NG",
    r"^COMP\.?\s*QTY",
    r"^COMPONENT\s*QTY",
    r"^AMOUNT",
]

PATTERNS_REVISION = [
    r"^REVLEV",
    r"^REV(\.|\b)",
    r"^REVISION",
    r"^REV\s*R3",
    r"^REVISION\s*LEVEL",
    r"^ZREV",
]

PATTERNS_ITEM_NUM = [
    r"^ITEM",
    r"^ITEM\s*NO",
    r"^POSNR",
    r"^V[ỊI]\s*TR[ÍI]",
    r"^STT",
]

PATTERNS_LEVEL = [
    r"^LEVEL",
    r"^STUFE",
    r"^C[ẤA]P",
    r"^T[ẦA]NG",
]

PATTERNS_DESCRIPTION = [
    r"^DESCRIPTION",
    r"^OBJECT\s*DESCRIPTION",
    r"^MAKTX",
    r"^T[ÊE]N\s*LINH\s*KI[ỆE]N",
    r"^DI[ỄE]N\s*GI[ẢA]I",
]


class ResilientR3Parser:
    """Parser for SAP R3 CS12 BOM export files."""

    def __init__(self):
        pass

    def parse(self, filepath: Union[str, Path]) -> pd.DataFrame:
        """
        Parse an SAP R3 CS12 exported file into a normalized pandas DataFrame.

        Columns produced:
            ['part_code', 'quantity', 'rev_r3'] (always guaranteed)
            plus ['level', 'item_num', 'description'] if detected.
        """
        path = Path(filepath)
        if not path.exists():
            raise SAPParseError(f"R3 CS12 file does not exist: {path}")

        if path.stat().st_size == 0:
            raise SAPParseError(f"R3 CS12 file is empty (0 bytes): {path}")

        # 1. Load raw tabular data from file
        df_raw = self._read_file_raw(path)
        if df_raw.empty or len(df_raw) < 1:
            raise SAPParseError(f"No tabular data extracted from file: {path}")

        # 2. Extract and normalize
        return self._normalize_dataframe(df_raw)

    def parse_content(self, content: Union[str, bytes], format_hint: Optional[str] = None) -> pd.DataFrame:
        """Parse raw string or bytes in-memory (useful for unit testing and streaming)."""
        df_raw = self._read_content_raw(content, format_hint=format_hint)
        return self._normalize_dataframe(df_raw)

    def parse_to_records(self, filepath: Union[str, Path]) -> List[R3ComponentRow]:
        """Parse file and return a list of typed R3ComponentRow objects."""
        df = self.parse(filepath)
        records = []
        for _, row in df.iterrows():
            rec = R3ComponentRow(
                part_code=str(row.get("part_code", "")),
                quantity=float(row.get("quantity", 0.0)),
                rev_r3=str(row.get("rev_r3", "")),
                item_num=str(row["item_num"]) if "item_num" in row and pd.notna(row["item_num"]) else None,
                description=str(row["description"]) if "description" in row and pd.notna(row["description"]) else None,
                level=int(row["level"]) if "level" in row and pd.notna(row["level"]) else None,
            )
            records.append(rec)
        return records

    @classmethod
    def _parse_html_table_robust(cls, html_text: str) -> Optional[pd.DataFrame]:
        """Parse HTML table using SimpleHTMLTableParser preserving exact string content."""
        try:
            parser = SimpleHTMLTableParser()
            parser.feed(html_text)
            if parser.tables:
                rows = parser.tables[0]
                if rows:
                    max_cols = max(len(r) for r in rows)
                    padded = [r + [""] * (max_cols - len(r)) for r in rows]
                    return pd.DataFrame(padded)
        except Exception:
            pass
        return None

    def _read_file_raw(self, path: Path) -> pd.DataFrame:
        """Attempt reading across multiple formats supported by SAPLSPO5 exports."""
        # Method 0: Check for Binary Excel magic bytes (ZIP for .xlsx / OLE for .xls)
        try:
            with open(path, "rb") as f:
                magic = f.read(4)
            if magic.startswith(b"PK\x03\x04") or magic.startswith(b"\xd0\xcf\x11\xe0") or path.suffix.lower() in (".xlsx", ".xlsm", ".xlsb"):
                df = pd.read_excel(str(path), header=None, dtype=str)
                if not df.empty:
                    logger.debug("Successfully parsed R3 file as binary Excel spreadsheet.")
                    return df
        except Exception:
            pass

        # Method 1: HTML table (most frequent in SAP GUI Spreadsheet exports)
        for encoding in ("utf-8", "cp1252", "latin1", "utf-16"):
            try:
                with open(path, "r", encoding=encoding, errors="ignore") as f:
                    content = f.read()
                if "<table" in content.lower():
                    df = self._parse_html_table_robust(content)
                    if df is not None and not df.empty:
                        logger.debug("Successfully parsed R3 file as HTML table via SimpleHTMLTableParser.")
                        return df
            except Exception:
                continue

        try:
            tables = pd.read_html(str(path), header=None)
            if tables:
                logger.debug("Successfully parsed R3 file as HTML table via pd.read_html.")
                return tables[0]
        except Exception:
            pass

        # Method 2: Robust text line reading (TSV, CSV, or semicolon)
        for encoding in ("utf-8", "cp1252", "latin1", "utf-16"):
            try:
                with open(path, "r", encoding=encoding, errors="replace") as f:
                    text = f.read()
                # Check for excessive binary null bytes indicating non-text
                if text.count("\x00") > 10:
                    continue
                df = self._parse_text_lines_robust(text)
                if df is not None and not df.empty and df.shape[1] > 1:
                    logger.debug(f"Successfully parsed R3 file as delimited text with encoding={encoding}.")
                    return df
            except Exception:
                continue

        # Method 3: Standard Excel (xlrd, openpyxl, etc.)
        try:
            df = pd.read_excel(str(path), header=None, dtype=str)
            if not df.empty:
                logger.debug("Successfully parsed R3 file as Excel spreadsheet.")
                return df
        except Exception:
            pass

        raise SAPParseError(f"Unsupported file format or unreadable spreadsheet content in {path.name}")


    def _read_content_raw(self, content: Union[str, bytes], format_hint: Optional[str] = None) -> pd.DataFrame:
        """Parse in-memory string or bytes."""
        if isinstance(content, str):
            # Try HTML
            if "<table" in content.lower():
                df = self._parse_html_table_robust(content)
                if df is not None and not df.empty:
                    return df
                try:
                    tables = pd.read_html(io.StringIO(content), header=None)
                    if tables:
                        return tables[0]
                except Exception:
                    pass

            # Try robust line parsing
            df = self._parse_text_lines_robust(content)
            if df is not None and not df.empty:
                return df
        else:
            # Bytes
            # Try HTML
            for enc in ("utf-8", "cp1252", "latin1"):
                try:
                    text = content.decode(enc)
                    if "<table" in text.lower():
                        df = self._parse_html_table_robust(text)
                        if df is not None and not df.empty:
                            return df
                except Exception:
                    continue

            try:
                tables = pd.read_html(io.BytesIO(content), header=None)
                if tables:
                    return tables[0]
            except Exception:
                pass

            # Try Excel
            try:
                return pd.read_excel(io.BytesIO(content), header=None, dtype=str)
            except Exception:
                pass

            # Try robust line parsing across encodings
            for enc in ("utf-8", "cp1252", "latin1"):
                try:
                    text = content.decode(enc)
                    df = self._parse_text_lines_robust(text)
                    if df is not None and not df.empty:
                        return df
                except Exception:
                    continue

        raise SAPParseError("Unable to parse in-memory content across supported tabular formats.")


    @staticmethod
    def _parse_text_lines_robust(text: str) -> Optional[pd.DataFrame]:
        """Split text lines, detect separator, and create rectangular DataFrame with padding."""
        lines = [l for l in text.splitlines() if l.strip()]
        if not lines:
            return None

        tab_count = sum(l.count("\t") for l in lines)
        comma_count = sum(l.count(",") for l in lines)
        semi_count = sum(l.count(";") for l in lines)

        if tab_count >= comma_count and tab_count >= semi_count and tab_count > 0:
            sep = "\t"
        elif comma_count >= semi_count and comma_count > 0:
            sep = ","
        elif semi_count > 0:
            sep = ";"
        else:
            sep = None

        if sep is not None:
            rows = [line.split(sep) for line in lines]
        else:
            rows = [re.split(r"\s{2,}", line.strip()) for line in lines]

        max_cols = max(len(r) for r in rows) if rows else 0
        if max_cols < 2:
            # Try splitting by multiple whitespace as fallback
            rows = [re.split(r"\s+", line.strip()) for line in lines]
            max_cols = max(len(r) for r in rows) if rows else 0

        padded = [r + [""] * (max_cols - len(r)) for r in rows]
        df = pd.DataFrame(padded)
        return df

    def _normalize_dataframe(self, df_raw: pd.DataFrame) -> pd.DataFrame:
        """Locates header row, dynamically identifies target columns, cleans data."""
        # Prepend column names if they are text headers (e.g. from pd.read_html)
        has_text_headers = any(
            isinstance(c, str) and re.search(r"[A-Za-z\u00C0-\u1EF9]", str(c))
            for c in df_raw.columns
        )
        if has_text_headers:
            header_row = pd.DataFrame([list(df_raw.columns)], columns=df_raw.columns)
            df_raw = pd.concat([header_row, df_raw], ignore_index=True)

        df_raw.columns = range(df_raw.shape[1])
        header_row_idx, col_map = self._detect_headers_and_columns(df_raw)

        # Slice data starting immediately after the header row
        df_data = df_raw.iloc[header_row_idx + 1:].copy()


        # Extract mapped columns
        part_idx = col_map.get("part_code")
        qty_idx = col_map.get("quantity")
        rev_idx = col_map.get("rev_r3")
        level_idx = col_map.get("level")
        item_idx = col_map.get("item_num")
        desc_idx = col_map.get("description")

        if part_idx is None:
            raise SAPParseError("Failed to identify Part Code / Component column in R3 file.")

        part_series = df_data.iloc[:, part_idx].astype(str).str.strip()

        # Quantity handling: clean numeric values
        if qty_idx is not None and qty_idx < df_data.shape[1]:
            qty_raw = df_data.iloc[:, qty_idx].astype(str).str.strip()
            # Handle comma vs period decimal format if needed (e.g. "1.000,00" or "1,000")
            qty_series = self._clean_numeric_quantities(qty_raw)
        else:
            qty_series = pd.Series(0.0, index=df_data.index)

        # Revision handling
        if rev_idx is not None and rev_idx < df_data.shape[1]:
            rev_series = df_data.iloc[:, rev_idx].astype(str).str.strip()
            # Replace 'nan', 'None' with empty string
            rev_series = rev_series.replace({"nan": "", "None": "", "NAN": "", "NONE": ""})
        else:
            rev_series = pd.Series("", index=df_data.index)

        result_dict = {
            "part_code": part_series,
            "quantity": qty_series,
            "rev_r3": rev_series,
        }

        # Optional metadata columns
        if level_idx is not None and level_idx < df_data.shape[1]:
            lvl_raw = pd.to_numeric(df_data.iloc[:, level_idx].astype(str).str.extract(r"(\d+)", expand=False), errors="coerce")
            result_dict["level"] = lvl_raw

        if item_idx is not None and item_idx < df_data.shape[1]:
            result_dict["item_num"] = df_data.iloc[:, item_idx].astype(str).str.strip()

        if desc_idx is not None and desc_idx < df_data.shape[1]:
            result_dict["description"] = df_data.iloc[:, desc_idx].astype(str).str.strip()

        df_clean = pd.DataFrame(result_dict)

        # Filter out empty rows, header duplicates, and visual separators
        df_clean = df_clean[df_clean["part_code"].str.len() > 0]
        ignore_values = {
            "PART CODE", "COMPONENT", "MATERIAL", "MATNR", "NAN", "NONE",
            "ITEM CODE", "LINH KIEN", "MA LINH KIEN",
        }
        df_clean = df_clean[~df_clean["part_code"].str.upper().isin(ignore_values)]
        # Filter out rows that look like border lines (e.g. "-------" or "======")
        df_clean = df_clean[~df_clean["part_code"].str.contains(r"^[-=_*#\s]+$", regex=True)]

        # Reset index
        df_clean = df_clean.reset_index(drop=True)
        return df_clean

    def _detect_headers_and_columns(self, df_raw: pd.DataFrame) -> Tuple[int, Dict[str, int]]:
        """
        Scan rows up to row 30 to locate the header row and map target column indices.
        Returns: (header_row_index, col_index_mapping)
        """
        max_scan_rows = min(35, len(df_raw))
        best_row_idx = 10  # Legacy standard header row is index 10 (row 11)
        best_col_map: Dict[str, int] = {}
        max_score = -1

        for r in range(max_scan_rows):
            row_values = [str(v).strip() for v in df_raw.iloc[r].tolist()]
            col_map = self._match_columns_in_row(row_values)
            score = 0
            if "part_code" in col_map:
                score += 5
            if "quantity" in col_map:
                score += 3
            if "rev_r3" in col_map:
                score += 2
            if "level" in col_map:
                score += 1
            if "item_num" in col_map:
                score += 1
            if "description" in col_map:
                score += 1

            if score > max_score:
                max_score = score
                best_row_idx = r
                best_col_map = col_map

            # If we found at least part_code and quantity, this is a strong match
            if "part_code" in col_map and "quantity" in col_map and score >= 8:
                best_row_idx = r
                best_col_map = col_map
                break

        # If no explicit header matched, fallback to legacy fixed column heuristics
        if "part_code" not in best_col_map:
            # If explicit column headers were found (e.g. quantity, level, description) but part_code is missing
            if any(k in best_col_map for k in ("quantity", "level", "description", "item_num")):
                raise SAPParseError("Failed to identify Part Code / Component column in R3 file.")

            logger.warning(
                f"No explicit header matched in top {max_scan_rows} rows. "
                "Applying resilient fallback column heuristics."
            )
            best_col_map = self._apply_fallback_column_map(df_raw)

        return best_row_idx, best_col_map

    def _match_columns_in_row(self, row_values: List[str]) -> Dict[str, int]:
        """Examine a row's cells against regular expressions."""
        col_map: Dict[str, int] = {}

        for idx, val in enumerate(row_values):
            clean_val = val.strip().upper()
            if not clean_val or clean_val in ("NAN", "NONE"):
                continue

            if "part_code" not in col_map and self._matches_patterns(clean_val, PATTERNS_PART_CODE):
                col_map["part_code"] = idx
            elif "quantity" not in col_map and self._matches_patterns(clean_val, PATTERNS_QUANTITY):
                col_map["quantity"] = idx
            elif "rev_r3" not in col_map and self._matches_patterns(clean_val, PATTERNS_REVISION):
                col_map["rev_r3"] = idx
            elif "level" not in col_map and self._matches_patterns(clean_val, PATTERNS_LEVEL):
                col_map["level"] = idx
            elif "item_num" not in col_map and self._matches_patterns(clean_val, PATTERNS_ITEM_NUM):
                col_map["item_num"] = idx
            elif "description" not in col_map and self._matches_patterns(clean_val, PATTERNS_DESCRIPTION):
                col_map["description"] = idx

        return col_map

    @staticmethod
    def _matches_patterns(text: str, patterns: List[str]) -> bool:
        """Check if normalized text matches any regex in patterns."""
        for pat in patterns:
            if re.search(pat, text, flags=re.IGNORECASE):
                return True
        return False

    @staticmethod
    def _apply_fallback_column_map(df_raw: pd.DataFrame) -> Dict[str, int]:
        """
        Fallback column map based on legacy SAP CS12 ALV standard:
        Col E = index 4: Part Code
        Col G = index 6: RevLev / Revision (if present)
        Col J = index 9: Quantity (or Col I = index 8)
        """
        num_cols = df_raw.shape[1]
        col_map = {}
        if num_cols > 4:
            col_map["part_code"] = 4
        else:
            col_map["part_code"] = 0

        if num_cols > 9:
            col_map["quantity"] = 9
        elif num_cols > 8:
            col_map["quantity"] = 8
        elif num_cols > 1:
            col_map["quantity"] = 1

        if num_cols > 6:
            col_map["rev_r3"] = 6

        return col_map

    @staticmethod
    def _clean_numeric_quantities(series: pd.Series) -> pd.Series:
        """
        Parse diverse numeric representations in SAP exports safely:
        - "1,000.00" -> 1000.0
        - "1.000,00" -> 1000.0 (German notation)
        - "2" -> 2.0
        - " 1.5 " -> 1.5
        - "1-" -> -1.0 (SAP trailing minus)
        """
        def clean_val(val: Any) -> float:
            if pd.isna(val):
                return 0.0
            s = str(val).strip()
            if not s or s.upper() in ("NAN", "NONE"):
                return 0.0

            # SAP trailing minus sign (e.g. "100-")
            is_negative = False
            if s.endswith("-"):
                is_negative = True
                s = s[:-1].strip()
            elif s.startswith("-"):
                is_negative = True
                s = s[1:].strip()

            # Check European format: dots as thousand separators, comma as decimal (e.g. 1.250,50)
            if re.search(r"^\d{1,3}(\.\d{3})*,\d+$", s):
                s = s.replace(".", "").replace(",", ".")
            # Or standard format: commas as thousand separators (e.g. 1,250.50)
            elif re.search(r"^\d{1,3}(,\d{3})*\.\d+$", s):
                s = s.replace(",", "")
            # Or comma only as decimal (e.g. 1,5)
            elif "," in s and "." not in s:
                s = s.replace(",", ".")

            try:
                num = float(s)
                return -num if is_negative else num
            except ValueError:
                return 0.0

        return series.apply(clean_val)


# Module-level convenience function matching legacy / specification signatures
def parse_r3_cs12_file(filepath: Union[str, Path]) -> pd.DataFrame:
    """
    Convenience function to parse an SAP R3 CS12 exported file.
    Returns a DataFrame with columns: ['part_code', 'quantity', 'rev_r3'].
    """
    parser = ResilientR3Parser()
    return parser.parse(filepath)
