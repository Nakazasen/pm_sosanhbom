"""Tier 3 Combination Tests: Multi-Source Reconciliation Pipeline (PLM + R3 + CTTT).

Verifies:
1. Three-way data flow joining parsed PLM, parsed R3 CS12, and CTTT instructions.
2. Accurate row-level checks (Compare PLM Qty, Compare R3 Qty, Compare Rev, Overall Check).
3. Identification of reverse-lookup missing parts (Sheet PLM) not assigned to any CTTT station.
4. Aggregation of cross-station grand totals (Sheet CTTT_Total) across sub-units vs R3 totals.
5. Fail-safe overall machine status determination ('OK' vs 'NG') reflecting component integrity.
"""

from pathlib import Path
import pandas as pd
import pytest

from src.automation.sap.parser import ResilientR3Parser
from src.core.reconciliation import (
    ReconciliationEngine,
    aggregate_cross_station,
    detect_missing_parts,
    reconcile_three_way,
)
from src.core.tree_parser import PLMTreeParser
from src.core.unit_resolver import UnitResolver


class TestPipelineReconciliation:
    """Combinatorial pipeline tests integrating PLM, R3, and CTTT datasets."""

    def test_full_reconciliation_flow(
        self,
        sample_plm_excel_file: Path,
        sample_cs12_file: Path,
        sample_cttt_df: pd.DataFrame,
    ):
        """Execute full three-way cross-reconciliation across realistic test datasets."""
        # 1. Parse and resolve PLM tree
        plm_tree = PLMTreeParser().parse_excel(sample_plm_excel_file)
        resolved_tree = UnitResolver().resolve_tree(plm_tree)

        # 2. Parse SAP R3 CS12 export
        r3_df = ResilientR3Parser().parse(sample_cs12_file)

        # 3. Execute three-way reconciliation
        engine = ReconciliationEngine()
        reconciled_df = engine.reconcile_three_way(
            cttt_data=sample_cttt_df,
            plm_data=resolved_tree,
            r3_data=r3_df,
        )

        assert len(reconciled_df) == len(sample_cttt_df)
        assert "Compare (PLM Qty)" in reconciled_df.columns
        assert "Compare (R3 Qty)" in reconciled_df.columns
        assert "Compare (Rev)" in reconciled_df.columns
        assert "Check" in reconciled_df.columns

    def test_reverse_lookup_missing_parts_pipeline(
        self,
        sample_plm_excel_file: Path,
        sample_cttt_df: pd.DataFrame,
    ):
        """Detect parts present in PLM engineering BOM but completely absent in CTTT."""
        plm_tree = PLMTreeParser().parse_excel(sample_plm_excel_file)

        missing_df = detect_missing_parts(
            plm_parts=plm_tree,
            cttt_parts=sample_cttt_df,
        )

        assert isinstance(missing_df, pd.DataFrame)
        # Any part in PLM not found in CTTT is flagged as missing
        cttt_parts = set(sample_cttt_df["MÃ LINH KIỆN"].str.strip().str.upper())
        plm_parts = set(n.item_id.strip().upper() for n in plm_tree.flatten())
        expected_missing = plm_parts - cttt_parts

        if expected_missing:
            assert len(missing_df) >= 1

    def test_cross_station_totals_aggregation_pipeline(
        self,
        sample_plm_excel_file: Path,
        sample_cs12_file: Path,
        sample_cttt_df: pd.DataFrame,
    ):
        """Aggregate duplicate parts across multiple stations and compare to R3 grand total."""
        plm_tree = PLMTreeParser().parse_excel(sample_plm_excel_file)
        r3_df = ResilientR3Parser().parse(sample_cs12_file)

        engine = ReconciliationEngine()
        reconciled_df = engine.reconcile_three_way(
            cttt_data=sample_cttt_df,
            plm_data=plm_tree,
            r3_data=r3_df,
        )

        totals_df = aggregate_cross_station(
            cttt_rows=reconciled_df,
            r3_totals=r3_df,
        )

        assert isinstance(totals_df, pd.DataFrame)
        assert "CTTT_SUM_QTY" in totals_df.columns
        assert "STATUS" in totals_df.columns
        assert len(totals_df) <= len(reconciled_df)

    def test_reconciliation_overall_status_evaluation(self):
        """Overall machine check is OK if all rows OK; NG if any single row NG."""
        engine = ReconciliationEngine()

        cttt_ok = pd.DataFrame([
            {"MÃ LINH KIỆN": "P1", "SỐ LƯỢNG": 2.0, "TÊN LINH KIỆN": "PART 1"},
        ])
        plm_ok = pd.DataFrame([
            {"part_code": "P1", "quantity": 2.0, "revision": "A"},
        ])
        r3_ok = pd.DataFrame([
            {"part_code": "P1", "quantity": 2.0, "revision": "A"},
        ])

        rec_ok = engine.reconcile_three_way(cttt_ok, plm_ok, r3_ok)
        assert all(rec_ok["Check"] == "OK")

        # Now introduce a mismatch
        r3_ng = pd.DataFrame([
            {"part_code": "P1", "quantity": 1.0, "revision": "A"},  # Qty mismatch (1.0 vs 2.0)
        ])
        rec_ng = engine.reconcile_three_way(cttt_ok, plm_ok, r3_ng)
        assert any(rec_ng["Check"] == "NG")
