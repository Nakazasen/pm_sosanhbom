"""Challenger 2 Empirical Challenge Test Suite: Reconciliation, Edge Cases & Fault Injection.

Empirically probes:
1. Floating-point rounding errors and precision in component quantities.
2. String revision variations (leading zeros '01' vs '1', whitespace, casing, suffix notations).
3. Adapter fault injection:
   - SAP COM sudden disconnection and recovery.
   - TC14 headless browser crash and session recovery.
   - Corrupted and unreadable Excel workbooks.
   - Cyclic BOM graphs and infinite recursion prevention across all tree engines.
"""

from __future__ import annotations

import datetime
import io
import math
import tempfile
import zipfile
from pathlib import Path
from unittest.mock import MagicMock, patch

import numpy as np
import pandas as pd
import pytest
from selenium.common.exceptions import WebDriverException

from src.automation.sap.connection import SAPConnectionManager
from src.automation.sap.cs12 import CS12Service
from src.automation.sap.models import SAPConnectionError, SAPCS12Error, SAPParseError
from src.automation.sap.parser import ResilientR3Parser
from src.automation.tc14.client import TC14AutomationClient, TC14Error
from src.automation.tc14.session import BrowserConfig, TC14SessionManager
from src.core.date_filter import DateFilter
from src.core.model_pruner import ModelPruner
from src.core.models import BOMNode, BOMTree, ModelRule
from src.core.reconciliation import (
    ReconciliationEngine,
    _clean_part_code,
    _normalize_rev,
    _to_float,
    aggregate_cross_station,
    reconcile_single_row,
    reconcile_three_way,
)
from src.core.tree_parser import PLMTreeParser
from src.core.unit_resolver import UnitResolver


# =============================================================================
# 1. Floating-point Precision and Component Quantity Probing
# =============================================================================
class TestQuantityPrecisionAndRounding:
    """Probes floating point arithmetic, tolerances, edge numbers, and representations."""

    def test_float_addition_imprecision_within_tolerance(self):
        """Probe 0.1 + 0.2 = 0.30000000000000004 vs 0.3 matches within 1e-6."""
        cttt_qty = 0.1 + 0.2
        plm_qty = 0.3
        r3_qty = 0.3
        res = reconcile_single_row(cttt_qty, plm_qty, r3_qty, "A", "A")
        assert res["comp_plm_qty"] == "OK"
        assert res["comp_r3_qty"] == "OK"
        assert res["overall_check"] == "OK"

    def test_repeated_floating_point_accumulation_cross_station(self):
        """Probe 100 sub-stations each with 0.01 quantity aggregating to 1.0."""
        # 100 * 0.01 in IEEE 754 often accumulates error: 1.0000000000000007
        rows = [
            {"MÃ LINH KIỆN": "PART_TINY", "SỐ LƯỢNG": 0.01, "SUB": f"STATION_{i}"}
            for i in range(100)
        ]
        r3_totals = pd.DataFrame([{"part_code": "PART_TINY", "quantity": 1.0}])

        agg = aggregate_cross_station(rows, r3_totals)
        assert len(agg) == 1
        assert agg.iloc[0]["STATUS"] == "OK"
        assert abs(agg.iloc[0]["CTTT_SUM_QTY"] - 1.0) < 1e-6

    def test_delta_at_and_beyond_tolerance_boundary(self):
        """Probe boundary around tolerance=1e-6 (9.99e-7 vs 1.001e-6)."""
        base = 10.0
        # Just within tolerance
        res_ok = reconcile_single_row(base, base + 9.9e-7, base - 9.9e-7, "A", "A", tolerance=1e-6)
        assert res_ok["comp_plm_qty"] == "OK"
        assert res_ok["comp_r3_qty"] == "OK"

        # Just exceeding tolerance
        res_ng = reconcile_single_row(base, base + 1.01e-6, base, "A", "A", tolerance=1e-6)
        assert res_ng["comp_plm_qty"] == "NG"
        assert res_ng["overall_check"] == "NG"

    def test_european_comma_decimal_handling_in_to_float(self):
        """Probe whether European/Vietnamese comma decimal '1,5' is converted correctly or coerced to 0.0."""
        # In manufacturing shop floors, operators frequently enter '1,5' or '0,5'
        val = _to_float("1,5")
        # Defect probe: Python's float("1,5") raises ValueError and _to_float returns 0.0!
        assert val == 1.5, f"Defect: '1,5' was converted to {val} instead of 1.5"

    def test_nan_and_inf_quantity_rejection(self):
        """Probe that NaN or Infinity in quantities fail-closed and raise ValueError."""
        with pytest.raises(ValueError, match="Quantity cannot be NaN"):
            reconcile_single_row(float("nan"), 1.0, 1.0, "A", "A")

        with pytest.raises(ValueError, match="Quantity cannot be Infinity"):
            reconcile_single_row(1.0, float("inf"), 1.0, "A", "A")

    def test_zero_quantity_behavior_when_all_sources_agree_zero(self):
        """Probe whether quantity=0 in all three sources (CTTT=0, PLM=0, R3=0) is marked OK or NG."""
        # form_ssbom formula: IF(OR(G3=0, K3=0, ...), "NG", "OK")
        res = reconcile_single_row(0.0, 0.0, 0.0, "A", "A")
        # By design per legacy formula, 0 quantity in BOM is flagged NG
        assert res["overall_check"] == "NG"

    def test_extreme_large_quantity_scale(self):
        """Probe astronomical quantities (e.g. bulk washers or raw bulk items)."""
        huge_qty = 10_000_000.0
        res = reconcile_single_row(huge_qty, huge_qty, huge_qty, "A", "A")
        assert res["overall_check"] == "OK"


# =============================================================================
# 2. String Revision Variations and Normalization Probing
# =============================================================================
class TestRevisionVariationsAndNormalization:
    """Probes revision formats: leading zeros ('01' vs '1'), casing, whitespace, and suffix notations."""

    def test_leading_zero_revision_matching(self):
        """Probe PLM rev '01' vs SAP R3 rev '1'."""
        # In Kyocera engineering, Teamcenter exports '01', while SAP R3 or Excel HTML import yields '1'
        res = reconcile_single_row(1.0, 1.0, 1.0, plm_rev="01", r3_rev="1")
        # Defect probe: _normalize_rev('01') == '01' != '1' == _normalize_rev('1') -> comp_rev='NG'!
        assert res["comp_rev"] == "OK", (
            f"Defect: Leading-zero revision mismatch: PLM '01' vs R3 '1' produced comp_rev='{res['comp_rev']}'"
        )

    def test_case_insensitivity_in_revision_matching(self):
        """Probe lowercase vs uppercase engineering revisions ('a' vs 'A')."""
        res = reconcile_single_row(1.0, 1.0, 1.0, plm_rev="a", r3_rev="A")
        # Defect probe: _normalize_rev('a') == 'a' != 'A' == _normalize_rev('A') -> comp_rev='NG'!
        assert res["comp_rev"] == "OK", (
            f"Defect: Case-sensitive revision mismatch: 'a' vs 'A' produced comp_rev='{res['comp_rev']}'"
        )

    def test_float_representation_of_revision_from_excel(self):
        """Probe numeric revision imported via Excel as 1.0 vs string '01' or '1'."""
        norm_float = _normalize_rev(1.0)
        # 1.0 becomes '1.0'
        res = reconcile_single_row(1.0, 1.0, 1.0, plm_rev="01", r3_rev=1.0)
        assert res["comp_rev"] == "OK", (
            f"Defect: Float-imported revision 1.0 vs '01' produced comp_rev='{res['comp_rev']}'"
        )

    def test_whitespace_and_newline_in_revisions(self):
        """Probe leading/trailing whitespace, tabs, and newlines in revision tags."""
        assert _normalize_rev("  B \t\n") == "B"
        res = reconcile_single_row(1.0, 1.0, 1.0, plm_rev="  B \t", r3_rev="B\n")
        assert res["comp_rev"] == "OK"

    def test_slash_and_prefix_revision_notations(self):
        """Probe revisions with prefix notations like 'REV A' vs 'A' or '/A' vs 'A'."""
        # Teamcenter often formats revision as '/A;' or 'Rev.01'
        res = reconcile_single_row(1.0, 1.0, 1.0, plm_rev="Rev.01", r3_rev="01")
        assert res["comp_rev"] == "OK", (
            f"Defect: Revision notation 'Rev.01' vs '01' failed to match: '{res['comp_rev']}'"
        )


# =============================================================================
# 3. Fault Injection: External Adapters (SAP COM, TC14, Excel, Cycles)
# =============================================================================
class TestAdapterFaultInjectionSAPCOM:
    """Simulates sudden COM disconnection, RPC server crashes, and recovery."""

    def test_sap_com_sudden_disconnection_state_detection(self):
        """Probe whether SAPConnectionManager recognizes a severed COM connection."""
        mock_session = MagicMock()
        mgr = SAPConnectionManager()
        mgr.session = mock_session
        mgr.connection = MagicMock()
        mgr._connected = True

        # Initially reported as connected
        assert mgr.is_connected() is True

        # Simulate sudden COM crash (RPC server unavailable / disconnected from clients)
        # Windows COM HRESULT -2147417848: The object invoked has disconnected from its clients.
        com_disconnect_exc = Exception((-2147417848, "The object invoked has disconnected from its clients.", None, None))
        mock_session.findById.side_effect = com_disconnect_exc

        # is_logged_in catches Exception and returns False
        assert mgr.is_logged_in() is False

        # DEFECT PROBE: is_connected() should return False if session is disconnected,
        # but current implementation only checks:
        # bool(self._connected and (self.session is not None or self.connection is not None))
        is_conn = mgr.is_connected()
        assert is_conn is False, (
            "Defect: SAPConnectionManager.is_connected() returned True despite disconnected COM session!"
        )

    def test_sap_com_stale_session_recovery_on_get_session(self):
        """Probe whether get_session() recovers automatically or returns a dead session pointer."""
        mock_dead_session = MagicMock()
        mock_dead_session.findById.side_effect = Exception("COM disconnected")

        mgr = SAPConnectionManager()
        mgr.session = mock_dead_session

        # DEFECT PROBE: mgr.get_session() does: `return self.session or self.get_or_create_session()`
        # When self.session is a dead pointer, it returns self.session without health-checking!
        with patch.object(mgr, "get_or_create_session") as mock_reconnect:
            mock_new_session = MagicMock()
            mock_reconnect.return_value = mock_new_session

            retrieved = mgr.get_session()
            assert retrieved == mock_new_session, (
                "Defect: SAPConnectionManager.get_session() returned dead session instead of reconnecting!"
            )

    def test_cs12_service_error_handling_on_sudden_com_drop(self):
        """Probe CS12Service fails closed and returns ExportResult(success=False) when COM drops during export."""
        mock_session = MagicMock()
        mock_session.findById.side_effect = Exception("RPC Server Unavailable 0x800706BA")

        svc = CS12Service(session=mock_session)
        with tempfile.TemporaryDirectory() as tmpdir:
            result = svc.execute_cs12(
                material="110C103NL0",
                destination_dir=tmpdir,
                valid_date=datetime.date(2026, 9, 17),
            )
            assert result.success is False
            assert "RPC Server Unavailable" in result.error_message


class TestAdapterFaultInjectionTC14BrowserCrash:
    """Simulates browser driver crash, process termination, and session recovery."""

    def test_tc14_driver_crash_detection_in_is_session_alive(self):
        """Probe whether is_session_alive() detects a crashed/killed WebDriver."""
        mock_driver = MagicMock()
        session_mgr = TC14SessionManager(driver=mock_driver)
        session_mgr.mark_authenticated("vn_pe03")
        assert session_mgr.is_session_alive() is True

        # Simulate browser crash (process killed, window closed, connection refused)
        mock_driver.current_url = PropertyMock_or_Method = MagicMock(
            side_effect=WebDriverException("Session deleted because of page crash / browser closed")
        )
        type(mock_driver).current_url = PropertyMock_or_Method

        # DEFECT PROBE: is_session_alive calls is_login_page_present(), which catches the
        # WebDriverException and returns False. Then is_session_alive returns self._authenticated (True)!
        alive = session_mgr.is_session_alive()
        assert alive is False, (
            "Defect: TC14SessionManager.is_session_alive() returned True when browser had crashed!"
        )

    def test_tc14_client_login_name_error_timeout(self):
        """Probe the NameError 'timeout' bug in TC14AutomationClient.login()."""
        mock_driver = MagicMock()
        session_mgr = TC14SessionManager(driver_factory=lambda: mock_driver)
        client = TC14AutomationClient(session_manager=session_mgr)

        # Defect probe: Line 157 in client.py references undefined 'timeout'
        with pytest.raises(Exception) as exc_info:
            client.login()

        assert "name 'timeout' is not defined" not in str(exc_info.value), (
            "Defect: TC14AutomationClient.login() crashed with NameError: name 'timeout' is not defined"
        )


class TestCorruptedAndUnreadableExcelSheets:
    """Probes handling of empty, corrupt, truncated, or non-Excel files."""

    def test_plm_parser_with_zero_byte_file(self):
        """Probe PLMTreeParser on 0-byte Excel file."""
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            f.write(b"")
            f_path = Path(f.name)

        try:
            parser = PLMTreeParser()
            # Openpyxl raises InvalidFileException; verify whether parser wraps or propagates cleanly
            with pytest.raises(Exception) as exc_info:
                parser.parse_file(f_path)
            # Should be an informative exception, not an unhandled crash
            assert exc_info.type in (ValueError, FileNotFoundError, Exception)
        finally:
            f_path.unlink(missing_ok=True)

    def test_plm_parser_with_corrupted_zip_bytes(self):
        """Probe PLMTreeParser on corrupted random bytes disguised as .xlsx."""
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            f.write(b"PK\x03\x04CORRUPTED_ZIP_RANDOM_GARBAGE_BYTES_1234567890")
            f_path = Path(f.name)

        try:
            parser = PLMTreeParser()
            with pytest.raises(Exception) as exc_info:
                parser.parse_file(f_path)
            assert exc_info.value is not None
        finally:
            f_path.unlink(missing_ok=True)

    def test_r3_parser_with_corrupted_file(self):
        """Probe ResilientR3Parser on unreadable binary content raises SAPParseError."""
        with tempfile.NamedTemporaryFile(suffix=".xls", delete=False) as f:
            f.write(b"\x00\x01\x02\x03\x04\x05\x06\x07GARBAGE_BINARY_STREAM")
            f_path = Path(f.name)

        try:
            parser = ResilientR3Parser()
            with pytest.raises(SAPParseError, match="Unsupported file format or unreadable"):
                parser.parse(f_path)
        finally:
            f_path.unlink(missing_ok=True)


class TestCyclicBOMGraphsAndRecursion:
    """Probes recursion defenses across BOMNode, UnitResolver, DateFilter, and ModelPruner."""

    def test_bom_node_flatten_and_clone_with_self_cycle(self):
        """Probe BOMNode.flatten() and clone() on self-referential graph."""
        node = BOMNode(level=1, item_id="LOOP_PART", item_name="LOOP_NAME")
        node.children.append(node)

        # flatten() has cycle protection
        flat = node.flatten()
        assert len(flat) == 1

        # clone() has cycle protection
        cloned = node.clone()
        assert cloned.item_id == "LOOP_PART"

    def test_date_filter_recursive_infinite_loop_on_cyclic_graph(self):
        """Probe whether DateFilter._filter_node_recursive handles cycles or blows the call stack."""
        node_a = BOMNode(level=1, item_id="CYC_A", effectivity="01-Jan-2024 UP")
        node_b = BOMNode(level=2, item_id="CYC_B", effectivity="01-Jan-2024 UP")
        node_a.children.append(node_b)
        node_b.children.append(node_a)  # Cyclic loop: A -> B -> A

        tree = BOMTree(roots=[node_a])
        filtrator = DateFilter()

        # DEFECT PROBE: DateFilter._filter_node_recursive has NO visited set!
        # It recurses infinitely until RecursionError!
        try:
            filtrator.filter_tree(tree, in_place=True)
        except RecursionError:
            pytest.fail("Defect: DateFilter.filter_tree() entered infinite recursion (RecursionError) on cyclic BOM graph!")

    def test_model_pruner_recursive_infinite_loop_on_cyclic_graph(self):
        """Probe whether ModelPruner._apply_rule_to_nodes handles cycles or blows the call stack."""
        node_a = BOMNode(level=1, item_id="PRUNE_A", item_name="NO_MATCH_A")
        node_b = BOMNode(level=2, item_id="PRUNE_B", item_name="NO_MATCH_B")
        node_a.children.append(node_b)
        node_b.children.append(node_a)  # Cyclic loop: A -> B -> A

        tree = BOMTree(roots=[node_a])
        pruner = ModelPruner()
        rule = ModelRule(item_name="SOME_OTHER_RULE", match_mode="Full_name")

        # DEFECT PROBE: ModelPruner._apply_rule_to_nodes has NO visited set!
        # When nodes do not match rule, it recurses infinitely into node.children!
        try:
            pruner.prune_tree(tree, rules=[rule], in_place=True)
        except RecursionError:
            pytest.fail("Defect: ModelPruner.prune_tree() entered infinite recursion (RecursionError) on cyclic BOM graph!")
