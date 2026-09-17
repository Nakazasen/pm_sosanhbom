"""Unit tests for R5 Flexible Provider Adapter Architecture.

Verifies:
1. Abstract base class contract enforcement (PLMProvider, ERPProvider).
2. ExcelPLMAdapter file resolution, BOMTree parsing, and export.
3. TeamcenterSeleniumAdapter wrapping TC14AutomationClient.
4. ExcelR3Adapter file resolution, DataFrame parsing, and export.
5. SAPR3COMAdapter wrapping CS12Service.
6. Complete Core Engine Independence: swapping providers produces 100% bit-accurate
   reconciliation results without touching ReconciliationEngine.
"""

from __future__ import annotations

import datetime
from pathlib import Path
from unittest.mock import MagicMock, patch

import pandas as pd
import pytest

from src.core.adapters import (
    ERPProvider,
    ExcelPLMAdapter,
    ExcelR3Adapter,
    PLMProvider,
    SAPR3COMAdapter,
    TeamcenterSeleniumAdapter,
)
from src.core.models import BOMNode, BOMTree
from src.core.reconciliation import ReconciliationEngine


class TestProviderContracts:
    """Validate ABC instantiation prevention and interface adherence."""

    def test_plm_provider_abc_cannot_be_instantiated(self):
        with pytest.raises(TypeError):
            PLMProvider()  # type: ignore[abstract]

    def test_erp_provider_abc_cannot_be_instantiated(self):
        with pytest.raises(TypeError):
            ERPProvider()  # type: ignore[abstract]


class TestExcelPLMAdapter:
    """Test offline ExcelPLMAdapter functionality."""

    def test_resolve_file_from_file_map(self, tmp_path: Path):
        sample_file = tmp_path / "sample_plm.xlsx"
        sample_file.write_text("dummy", encoding="utf-8")

        adapter = ExcelPLMAdapter(file_map={"110C103NL0": sample_file})
        resolved = adapter._resolve_file("110C103NL0")
        assert resolved == sample_file

    def test_resolve_file_from_search_dir(self, tmp_path: Path):
        sample_file = tmp_path / "BOM_110C103NL0_full.xlsx"
        sample_file.write_text("dummy", encoding="utf-8")

        adapter = ExcelPLMAdapter(search_dir=tmp_path)
        resolved = adapter._resolve_file("110C103NL0")
        assert resolved == sample_file

    def test_resolve_file_missing_raises_filenotfound(self):
        adapter = ExcelPLMAdapter()
        with pytest.raises(FileNotFoundError):
            adapter._resolve_file("NON_EXISTENT_PART")

    def test_fetch_bom_delegates_to_parser(self, tmp_path: Path):
        sample_file = tmp_path / "sample_plm.xlsx"
        sample_file.write_text("dummy", encoding="utf-8")

        mock_parser = MagicMock()
        expected_tree = BOMTree(roots=[BOMNode(level=1, item_id="PART_1", quantity=1.0)])
        mock_parser.parse_file.return_value = expected_tree

        adapter = ExcelPLMAdapter(default_file=sample_file, tree_parser=mock_parser)
        tree = adapter.fetch_bom("PART_1")
        assert tree == expected_tree
        mock_parser.parse_file.assert_called_once_with(sample_file)

    def test_export_excel_copies_to_destination(self, tmp_path: Path):
        sample_file = tmp_path / "src_plm.xlsx"
        sample_file.write_text("dummy_content", encoding="utf-8")
        out_file = tmp_path / "sub" / "dest_plm.xlsx"

        adapter = ExcelPLMAdapter(default_file=sample_file)
        result_path = adapter.export_excel("ANY_PART", output_path=out_file)
        assert result_path == out_file
        assert out_file.exists()
        assert out_file.read_text(encoding="utf-8") == "dummy_content"


class TestTeamcenterSeleniumAdapter:
    """Test TeamcenterSeleniumAdapter wrapping TC14AutomationClient."""

    def test_export_excel_calls_download_bom_full(self, tmp_path: Path):
        mock_client = MagicMock()
        exported_file = tmp_path / "exported.xlsx"
        exported_file.write_text("data", encoding="utf-8")
        mock_client.download_bom_full.return_value = exported_file

        adapter = TeamcenterSeleniumAdapter(client=mock_client)
        target = tmp_path / "target.xlsx"
        out = adapter.export_excel("110C103NL0", rev="A", output_path=target)

        assert out.exists()
        mock_client.download_bom_full.assert_called_once()

    def test_fetch_bom_parses_exported_file(self, tmp_path: Path):
        mock_client = MagicMock()
        temp_file = tmp_path / "temp.xlsx"
        temp_file.write_text("data", encoding="utf-8")
        mock_client.download_bom_full.return_value = temp_file

        mock_parser = MagicMock()
        expected_tree = BOMTree(roots=[BOMNode(level=1, item_id="TC_PART", quantity=2.0)])
        mock_parser.parse_file.return_value = expected_tree

        adapter = TeamcenterSeleniumAdapter(client=mock_client, tree_parser=mock_parser)
        tree = adapter.fetch_bom("TC_PART")
        assert tree == expected_tree
        mock_parser.parse_file.assert_called_once()


class TestExcelR3Adapter:
    """Test offline ExcelR3Adapter functionality."""

    def test_resolve_file_from_file_map(self, tmp_path: Path):
        sample_file = tmp_path / "r3_export.xls"
        sample_file.write_text("dummy", encoding="utf-8")

        adapter = ExcelR3Adapter(file_map={"MAT_01": sample_file})
        resolved = adapter._resolve_file("MAT_01")
        assert resolved == sample_file

    def test_resolve_file_missing_raises_filenotfound(self):
        adapter = ExcelR3Adapter()
        with pytest.raises(FileNotFoundError):
            adapter._resolve_file("NON_EXISTENT_MAT")

    def test_fetch_multilevel_bom_delegates_to_parser(self, tmp_path: Path):
        sample_r3 = tmp_path / "sample_r3.xls"
        sample_r3.write_text("dummy", encoding="utf-8")

        mock_parser = MagicMock()
        expected_df = pd.DataFrame([{"part_code": "PART_R3", "quantity": 3.0, "rev_r3": "01"}])
        mock_parser.parse.return_value = expected_df

        adapter = ExcelR3Adapter(default_file=sample_r3, parser=mock_parser)
        df = adapter.fetch_multilevel_bom("110C103NL0")
        assert df.equals(expected_df)
        mock_parser.parse.assert_called_once_with(sample_r3)

    def test_export_multilevel_bom_file(self, tmp_path: Path):
        sample_r3 = tmp_path / "sample_r3.xls"
        sample_r3.write_text("content", encoding="utf-8")
        out_dir = tmp_path / "dest_dir"

        adapter = ExcelR3Adapter(default_file=sample_r3)
        exported = adapter.export_multilevel_bom_file("MAT_01", destination_dir=out_dir)
        assert exported.exists()
        assert exported.name == "sample_r3.xls"


class TestSAPR3COMAdapter:
    """Test SAPR3COMAdapter wrapping CS12Service."""

    def test_fetch_multilevel_bom_executes_cs12_and_parses(self, tmp_path: Path):
        mock_cs12 = MagicMock()
        exported_file = tmp_path / "cs12_export.xls"
        exported_file.write_text("export data", encoding="utf-8")

        mock_result = MagicMock()
        mock_result.success = True
        mock_result.file_path = exported_file
        mock_cs12.execute_cs12.return_value = mock_result

        mock_parser = MagicMock()
        expected_df = pd.DataFrame([{"part_code": "MAT_01", "quantity": 10.0, "rev_r3": "00"}])
        mock_parser.parse.return_value = expected_df

        adapter = SAPR3COMAdapter(cs12_service=mock_cs12, parser=mock_parser)
        df = adapter.fetch_multilevel_bom("MAT_01")
        assert df.equals(expected_df)
        mock_cs12.execute_cs12.assert_called_once()
        mock_parser.parse.assert_called_once_with(exported_file)

    def test_export_multilevel_bom_failure_raises_runtime_error(self):
        mock_cs12 = MagicMock()
        mock_result = MagicMock()
        mock_result.success = False
        mock_result.error_message = "SAP Logon connection refused"
        mock_cs12.execute_cs12.return_value = mock_result

        adapter = SAPR3COMAdapter(cs12_service=mock_cs12)
        with pytest.raises(RuntimeError, match="SAP CS12 execution failed: SAP Logon connection refused"):
            adapter.export_multilevel_bom_file("FAILED_MAT")


class TestCoreEngineIndependenceViaAdapterSwapping:
    """Verify core ReconciliationEngine operates identically regardless of swapped provider."""

    def test_engine_independence_with_swapped_providers(self, tmp_path: Path):
        # 1. Setup sample datasets
        cttt_data = pd.DataFrame([
            {"SUB": "Unit-1", "TRANG CTTT": "1", "MÃ LINH KIỆN": "302FP02010", "TÊN LINH KIỆN": "COVER", "SỐ LƯỢNG": 2.0, "PHỤ TRÁCH": "User1"}
        ])
        plm_tree = BOMTree(roots=[
            BOMNode(level=1, item_id="302FP02010", item_name="COVER", quantity=2.0, revision="01", unit_name="Unit-1")
        ])
        r3_df = pd.DataFrame([
            {"part_code": "302FP02010", "quantity": 2.0, "rev_r3": "01"}
        ])

        # Provider Pair A: Offline Excel Adapters
        plm_file = tmp_path / "plm_test.xlsx"
        plm_file.write_text("dummy", encoding="utf-8")
        r3_file = tmp_path / "r3_test.xls"
        r3_file.write_text("dummy", encoding="utf-8")

        mock_plm_parser = MagicMock()
        mock_plm_parser.parse_file.return_value = plm_tree

        mock_r3_parser = MagicMock()
        mock_r3_parser.parse.return_value = r3_df

        excel_plm = ExcelPLMAdapter(default_file=plm_file, tree_parser=mock_plm_parser)
        excel_r3 = ExcelR3Adapter(default_file=r3_file, parser=mock_r3_parser)

        # Provider Pair B: Automated Selenium / COM Adapters (Mocked)
        mock_tc14_client = MagicMock()
        mock_tc14_client.download_bom_full.return_value = plm_file
        selenium_plm = TeamcenterSeleniumAdapter(client=mock_tc14_client, tree_parser=mock_plm_parser)

        mock_cs12_service = MagicMock()
        mock_cs12_res = MagicMock()
        mock_cs12_res.success = True
        mock_cs12_res.file_path = r3_file
        mock_cs12_service.execute_cs12.return_value = mock_cs12_res
        com_r3 = SAPR3COMAdapter(cs12_service=mock_cs12_service, parser=mock_r3_parser)

        # Execute Engine with Provider Pair A
        engine = ReconciliationEngine()
        tree_a = excel_plm.fetch_bom("VIRGO")
        df_r3_a = excel_r3.fetch_multilevel_bom("VIRGO")
        result_a = engine.run_full_reconciliation(
            cttt_data=cttt_data,
            plm_data=tree_a,
            r3_data=df_r3_a,
        )

        # Execute Engine with Provider Pair B
        tree_b = selenium_plm.fetch_bom("VIRGO")
        df_r3_b = com_r3.fetch_multilevel_bom("VIRGO")
        result_b = engine.run_full_reconciliation(
            cttt_data=cttt_data,
            plm_data=tree_b,
            r3_data=df_r3_b,
        )

        # Verify 100% bit-accurate identity across providers
        assert result_a.overall_status == result_b.overall_status == "OK"
        assert len(result_a.cttt_rows) == len(result_b.cttt_rows) == 1
        row_a = result_a.cttt_rows.iloc[0]
        row_b = result_b.cttt_rows.iloc[0]
        assert row_a["Compare (PLM Qty)"] == row_b["Compare (PLM Qty)"] == "OK"
        assert row_a["Compare (R3 Qty)"] == row_b["Compare (R3 Qty)"] == "OK"
        assert row_a["Compare (Rev)"] == row_b["Compare (Rev)"] == "OK"
        assert row_a["Check"] == row_b["Check"] == "OK"
