"""Tier 3 Combination Tests: MSI & Fix Serial Pipeline Crosscheck (F10 + Tree).

Verifies:
1. End-to-end integration between parsed PLM tree and MSI serialized component evaluation.
2. Branch 5 (Exact Match): Unit in PLM, 3-char MSI code matches, Service matches -> 'OK'.
3. Branch 1 & 3 (Missing from PLM): Serialized unit absent from engineering BOM -> 'NG'.
4. Branch 8 (Code Mismatch): Unit in PLM but 3-character prefix differs -> 'NG'.
5. Branch 7 (Service Note Warning): Code matches, CTTT has '-', master has required note -> 'OK' with service_warning.
6. Table evaluation: batch processing of Sheet MSI_7980_7990 against FixSerialMaster tool.
"""

from pathlib import Path
import pandas as pd
import pytest

from src.core.msi_engine import (
    FixSerialMaster,
    MSIEngine,
    evaluate_msi_branch,
)
from src.core.tree_parser import PLMTreeParser


class TestPipelineMSICrosscheck:
    """Combinatorial pipeline tests for MSI serialization and PLM cross-checking."""

    @pytest.fixture
    def configured_msi_engine(self) -> MSIEngine:
        """Create and populate an MSIEngine with known test master data."""
        master = FixSerialMaster()
        # Add Image Unit, Fuser, and Main machine codes
        master.add_subunit(unit_code="302FP93010", msi_code="1HN", service="SERVICE REQUIRED")
        master.add_subunit(unit_code="302FP93020", msi_code="2HN", service="")
        master.add_subunit(unit_code="302FP93030", msi_code="3HN", service="")
        master.add_machine(machine_code="110C103NL0", msi_code="NL0", service="")
        return MSIEngine(master=master)

    def test_msi_pipeline_branch5_exact_match(
        self,
        sample_plm_excel_file: Path,
        configured_msi_engine: MSIEngine,
    ):
        """Branch 5: Unit exists in PLM, 3-char code matches, service matches -> 'OK'."""
        plm_tree = PLMTreeParser().parse_excel(sample_plm_excel_file)
        plm_parts = set(n.item_id.strip().upper() for n in plm_tree.flatten())

        unit_code = "302FP93010"
        in_plm = unit_code.upper() in plm_parts
        assert in_plm is True

        # Lookup in master
        master_code, master_srv = configured_msi_engine.master.lookup_subunit(unit_code)
        assert master_code == "1HN"
        assert master_srv == "SERVICE REQUIRED"

        res = configured_msi_engine.evaluate_branch(
            in_plm=in_plm,
            member_code="1HN",
            master_code=master_code,
            member_service="SERVICE REQUIRED",
            master_service=master_srv,
        )

        assert res["status"] == "OK"
        assert res["branch"] == 5
        assert res["service_warning"] is False
        assert "K" in res["highlight_green"]

    def test_msi_pipeline_branch1_missing_from_plm(
        self,
        sample_plm_excel_file: Path,
        configured_msi_engine: MSIEngine,
    ):
        """Branch 1: Serialized unit code absent from PLM engineering BOM -> 'NG'."""
        plm_tree = PLMTreeParser().parse_excel(sample_plm_excel_file)
        plm_parts = set(n.item_id.strip().upper() for n in plm_tree.flatten())

        fake_unit = "302FAKE999"
        in_plm = fake_unit.upper() in plm_parts
        assert in_plm is False

        res = configured_msi_engine.evaluate_branch(
            in_plm=in_plm,
            member_code="999",
            master_code="999",
        )

        assert res["status"] == "NG"
        assert res["branch"] == 1
        assert "K" in res["highlight_red"]

    def test_msi_pipeline_branch8_3char_code_mismatch(
        self,
        sample_plm_excel_file: Path,
        configured_msi_engine: MSIEngine,
    ):
        """Branch 8: Unit in PLM, but member entered wrong 3-char code -> 'NG'."""
        plm_tree = PLMTreeParser().parse_excel(sample_plm_excel_file)
        plm_parts = set(n.item_id.strip().upper() for n in plm_tree.flatten())

        unit_code = "302FP93020"
        in_plm = unit_code.upper() in plm_parts
        assert in_plm is True

        master_code, _ = configured_msi_engine.master.lookup_subunit(unit_code)
        assert master_code == "2HN"

        # Member entered "WRONG_CODE"
        res = configured_msi_engine.evaluate_branch(
            in_plm=in_plm,
            member_code="WRG",
            master_code=master_code,
        )

        assert res["status"] == "NG"
        assert res["branch"] == 8
        assert "D" in res["highlight_red"]
        assert "K" in res["highlight_red"]

    def test_msi_pipeline_branch7_service_warning(
        self,
        sample_plm_excel_file: Path,
        configured_msi_engine: MSIEngine,
    ):
        """Branch 7: Code matches, member left service as '-', but master requires service -> 'OK' with warning."""
        plm_tree = PLMTreeParser().parse_excel(sample_plm_excel_file)
        unit_code = "302FP93010"
        in_plm = unit_code.upper() in set(n.item_id.strip().upper() for n in plm_tree.flatten())

        master_code, master_srv = configured_msi_engine.master.lookup_subunit(unit_code)
        assert master_srv == "SERVICE REQUIRED"

        res = configured_msi_engine.evaluate_branch(
            in_plm=in_plm,
            member_code="1HN",
            master_code=master_code,
            member_service="-",  # Member omitted note
            master_service=master_srv,
        )

        assert res["status"] == "OK"
        assert res["branch"] == 7
        assert res["service_warning"] is True
        assert "E" in res["highlight_red"]  # Service column flagged red

    def test_msi_batch_table_evaluation(
        self,
        sample_plm_excel_file: Path,
        sample_msi_df: pd.DataFrame,
        configured_msi_engine: MSIEngine,
    ):
        """Evaluate an entire MSI worksheet table against PLM tree."""
        plm_tree = PLMTreeParser().parse_excel(sample_plm_excel_file)

        evaluated_df = configured_msi_engine.evaluate_table(
            msi_data=sample_msi_df,
            plm_part_codes=[n.item_id for n in plm_tree.flatten()],
        )

        assert isinstance(evaluated_df, pd.DataFrame)
        assert len(evaluated_df) == len(sample_msi_df)
        assert "KẾT QUẢ" in evaluated_df.columns or "status" in evaluated_df.columns or "Ket_luan" in evaluated_df.columns
