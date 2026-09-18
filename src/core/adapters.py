"""R5 Flexible Provider Adapter Architecture.

Decouples the core comparison engine and tree algorithms from concrete data sources
(Selenium web automation vs local Excel files, SAP GUI COM automation vs exported TSV/HTML/XLS).

Enables hot-swapping between:
- PLMProvider: TeamcenterSeleniumAdapter <-> ExcelPLMAdapter
- ERPProvider: SAPR3COMAdapter <-> ExcelR3Adapter
"""

from __future__ import annotations

import abc
import datetime
import logging
import shutil
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

import pandas as pd

from src.core.models import BOMTree
from src.core.tree_parser import PLMTreeParser
from src.automation.sap.parser import ResilientR3Parser

logger = logging.getLogger(__name__)


# =============================================================================
# 1. PLM Provider Interface & Implementations
# =============================================================================

class PLMProvider(abc.ABC):
    """Abstract Base Class for PLM (Product Lifecycle Management) data providers."""

    @abc.abstractmethod
    def fetch_bom(self, item_id: str, rev: Optional[str] = None) -> BOMTree:
        """Fetch and parse the full BOM hierarchy for an item.

        Args:
            item_id: Part code / Item ID (e.g. '110C103NL0', 'Virgo', etc.).
            rev: Revision identifier (e.g. '01', 'A', etc.).

        Returns:
            Structured in-memory BOMTree.
        """
        pass

    @abc.abstractmethod
    def export_excel(
        self,
        item_id: str,
        rev: Optional[str] = None,
        output_path: Optional[Union[str, Path]] = None,
    ) -> Path:
        """Export the BOM structure to an Excel file on disk.

        Args:
            item_id: Part code / Item ID.
            rev: Optional revision identifier.
            output_path: Target output path (file or directory).

        Returns:
            Path to the saved Excel file.
        """
        pass


from src.core.tc2412_bridge import TC2412CanonicalNormalizer, TC2412SheetPLMBridge

class TeamcenterSeleniumAdapter(PLMProvider):
    """PLM Adapter wrapping TC2412AutomationClient for live Active Workspace 2412 extraction."""

    def __init__(
        self,
        client: Optional[Any] = None,
        tree_parser: Optional[PLMTreeParser] = None,
    ) -> None:
        """Initialize adapter with TC2412AutomationClient instance."""
        if client is None:
            try:
                from src.automation.tc2412.client import TC2412AutomationClient
                self.client = TC2412AutomationClient()
            except (ImportError, Exception):
                from src.automation.tc14.client import TC14AutomationClient
                self.client = TC14AutomationClient()
        else:
            self.client = client
        self.parser = tree_parser or PLMTreeParser()

    def export_excel(
        self,
        item_id: str,
        rev: Optional[str] = None,
        output_path: Optional[Union[str, Path]] = None,
        progress_callback: Optional[Any] = None,
    ) -> Path:
        """Export BOM Excel via TC2412AutomationClient."""
        target_dir = (
            Path(output_path).parent
            if output_path and Path(output_path).suffix
            else (Path(output_path) if output_path else Path("./downloads"))
        )
        target_dir.mkdir(parents=True, exist_ok=True)

        try:
            exported_file = self.client.download_bom_full(
                part_number=item_id,
                output_dir=target_dir,
                part_rev=rev,
                progress_callback=progress_callback,
            )
        except TypeError:
            exported_file = self.client.download_bom_full(
                part_number=item_id,
                output_dir=target_dir,
                part_rev=rev,
            )
        if output_path and Path(output_path).suffix and Path(exported_file) != Path(output_path):
            shutil.copy2(exported_file, output_path)
            return Path(output_path)
        return Path(exported_file)

    def fetch_bom(self, item_id: str, rev: Optional[str] = None) -> BOMTree:
        """Export BOM from Active Workspace and parse into a BOMTree."""
        temp_excel = self.export_excel(item_id, rev=rev)
        return self.parser.parse_file(temp_excel)



class ExcelPLMAdapter(PLMProvider):
    """PLM Adapter for reading local PLM Excel files without network or browser dependencies."""

    def __init__(
        self,
        file_map: Optional[Dict[str, Union[str, Path]]] = None,
        default_file: Optional[Union[str, Path]] = None,
        search_dir: Optional[Union[str, Path]] = None,
        tree_parser: Optional[PLMTreeParser] = None,
    ) -> None:
        """Initialize ExcelPLMAdapter.

        Args:
            file_map: Dictionary mapping item_id (or item_id:rev) to specific file paths.
            default_file: Fallback file path if item_id not in file_map.
            search_dir: Directory to search for matching files by name.
            tree_parser: Custom or default PLMTreeParser instance.
        """
        self.file_map = {k: Path(v) for k, v in (file_map or {}).items()}
        self.default_file = Path(default_file) if default_file else None
        self.search_dir = Path(search_dir) if search_dir else None
        self.parser = tree_parser or PLMTreeParser()

    def _resolve_file(self, item_id: str, rev: Optional[str] = None) -> Path:
        """Resolve file path for given item_id and rev."""
        key_with_rev = f"{item_id}:{rev}" if rev else item_id
        if key_with_rev in self.file_map and self.file_map[key_with_rev].exists():
            return self.file_map[key_with_rev]
        if item_id in self.file_map and self.file_map[item_id].exists():
            return self.file_map[item_id]

        if self.search_dir and self.search_dir.exists():
            candidates = list(self.search_dir.glob(f"*{item_id}*.xls*"))
            if candidates:
                return candidates[0]

        if self.default_file and self.default_file.exists():
            return self.default_file

        raise FileNotFoundError(
            f"ExcelPLMAdapter: Could not locate PLM Excel file for item_id='{item_id}', rev='{rev}'."
        )

    def fetch_bom(self, item_id: str, rev: Optional[str] = None) -> BOMTree:
        """Read and parse local PLM Excel file into a BOMTree."""
        file_path = self._resolve_file(item_id, rev=rev)
        return self.parser.parse_file(file_path)

    def export_excel(
        self,
        item_id: str,
        rev: Optional[str] = None,
        output_path: Optional[Union[str, Path]] = None,
    ) -> Path:
        """Return path to existing Excel file, copying to output_path if requested."""
        src_path = self._resolve_file(item_id, rev=rev)
        if output_path is None:
            return src_path

        dest = Path(output_path)
        if dest.is_dir() or not dest.suffix:
            dest.mkdir(parents=True, exist_ok=True)
            dest = dest / src_path.name
        else:
            dest.parent.mkdir(parents=True, exist_ok=True)
        if dest.resolve() != src_path.resolve():
            shutil.copy2(src_path, dest)
        return dest


# =============================================================================
# 2. ERP Provider Interface & Implementations
# =============================================================================

class ERPProvider(abc.ABC):
    """Abstract Base Class for ERP (Enterprise Resource Planning) data providers."""

    @abc.abstractmethod
    def fetch_multilevel_bom(
        self,
        material: str,
        plant: str = "2200",
        valid_date: Optional[Union[datetime.date, str]] = None,
        bom_usage: str = "pp01",
        alternative: str = "01",
    ) -> pd.DataFrame:
        """Fetch and return multilevel BOM data as a normalized pandas DataFrame.

        DataFrame guaranteed columns:
            ['part_code', 'quantity', 'rev_r3']
        """
        pass

    @abc.abstractmethod
    def export_multilevel_bom_file(
        self,
        material: str,
        plant: str = "2200",
        valid_date: Optional[Union[datetime.date, str]] = None,
        destination_dir: Optional[Union[str, Path]] = None,
        bom_usage: str = "pp01",
        alternative: str = "01",
    ) -> Path:
        """Export CS12 multilevel BOM to a file on disk.

        Returns:
            Path to exported file.
        """
        pass


class SAPR3COMAdapter(ERPProvider):
    """ERP Adapter wrapping CS12Service via SAP GUI Scripting COM interface."""

    def __init__(
        self,
        cs12_service: Optional[Any] = None,
        parser: Optional[ResilientR3Parser] = None,
    ) -> None:
        """Initialize SAPR3COMAdapter."""
        if cs12_service is None:
            from src.automation.sap.cs12 import CS12Service
            self.cs12_service = CS12Service()
        else:
            self.cs12_service = cs12_service
        self.parser = parser or ResilientR3Parser()

    def export_multilevel_bom_file(
        self,
        material: str,
        plant: str = "2200",
        valid_date: Optional[Union[datetime.date, str]] = None,
        destination_dir: Optional[Union[str, Path]] = None,
        bom_usage: str = "pp01",
        alternative: str = "01",
    ) -> Path:
        """Execute CS12 transaction via SAP COM and export result file."""
        dest = Path(destination_dir or "./downloads")
        dest.mkdir(parents=True, exist_ok=True)

        date_val = valid_date
        if isinstance(date_val, str):
            date_val = datetime.date.fromisoformat(date_val)
        elif date_val is None:
            date_val = datetime.date.today()

        result = self.cs12_service.execute_cs12(
            material=material,
            plant=plant,
            valid_date=date_val,
            bom_usage=bom_usage,
            alternative=alternative,
            destination_dir=dest,
        )
        if not result.success or not result.file_path:
            raise RuntimeError(f"SAP CS12 execution failed: {result.error_message}")

        return Path(result.file_path)

    def fetch_multilevel_bom(
        self,
        material: str,
        plant: str = "2200",
        valid_date: Optional[Union[datetime.date, str]] = None,
        bom_usage: str = "pp01",
        alternative: str = "01",
    ) -> pd.DataFrame:
        """Export CS12 via COM and parse into a normalized DataFrame."""
        file_path = self.export_multilevel_bom_file(
            material=material,
            plant=plant,
            valid_date=valid_date,
            bom_usage=bom_usage,
            alternative=alternative,
        )
        return self.parser.parse(file_path)


class ExcelR3Adapter(ERPProvider):
    """ERP Adapter reading local SAP R3 CS12 export files without SAP GUI or COM."""

    def __init__(
        self,
        file_map: Optional[Dict[str, Union[str, Path]]] = None,
        default_file: Optional[Union[str, Path]] = None,
        search_dir: Optional[Union[str, Path]] = None,
        parser: Optional[ResilientR3Parser] = None,
    ) -> None:
        """Initialize ExcelR3Adapter."""
        self.file_map = {k: Path(v) for k, v in (file_map or {}).items()}
        self.default_file = Path(default_file) if default_file else None
        self.search_dir = Path(search_dir) if search_dir else None
        self.parser = parser or ResilientR3Parser()

    def _resolve_file(self, material: str) -> Path:
        """Resolve file path for given material."""
        if material in self.file_map and self.file_map[material].exists():
            return self.file_map[material]

        if self.search_dir and self.search_dir.exists():
            candidates = list(self.search_dir.glob(f"*{material}*.xls*"))
            if candidates:
                return candidates[0]

        if self.default_file and self.default_file.exists():
            return self.default_file

        raise FileNotFoundError(
            f"ExcelR3Adapter: Could not locate SAP R3 export file for material='{material}'."
        )

    def fetch_multilevel_bom(
        self,
        material: str,
        plant: str = "2200",
        valid_date: Optional[Union[datetime.date, str]] = None,
        bom_usage: str = "pp01",
        alternative: str = "01",
    ) -> pd.DataFrame:
        """Read and parse local SAP R3 file into a normalized DataFrame."""
        file_path = self._resolve_file(material)
        return self.parser.parse(file_path)

    def export_multilevel_bom_file(
        self,
        material: str,
        plant: str = "2200",
        valid_date: Optional[Union[datetime.date, str]] = None,
        destination_dir: Optional[Union[str, Path]] = None,
        bom_usage: str = "pp01",
        alternative: str = "01",
    ) -> Path:
        """Return path to existing R3 export file, copying if destination requested."""
        src_path = self._resolve_file(material)
        if destination_dir is None:
            return src_path

        dest = Path(destination_dir)
        if dest.is_dir() or not dest.suffix:
            dest.mkdir(parents=True, exist_ok=True)
            dest = dest / src_path.name
        else:
            dest.parent.mkdir(parents=True, exist_ok=True)
        if dest.resolve() != src_path.resolve():
            shutil.copy2(src_path, dest)
        return dest
