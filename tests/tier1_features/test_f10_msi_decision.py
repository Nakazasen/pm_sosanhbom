"""Feature F10: MSI & Fix Serial Decision Engine Isolation Tests.

Verifies the 9-branch decision table from legacy msi.bas (lines 120-379)
evaluating serialized assemblies, 3-character MSI fixed codes, and service comments:
- Branch 1: Unit code does not exist in PLM -> 'NG'
- Branch 2: Normalize blank service comment to '-'
- Branch 4: Unit in PLM but missing in Fix Serial Master Tool -> 'NG'
- Branch 5 & 6: Code and Service match -> 'OK'
- Branch 7: Code matches, but CTTT missed required Service note -> 'OK' with service warning
- Branch 8: 3-character MSI code mismatch -> 'NG'
- Branch 9: Service comment mismatch -> 'NG'
"""

from __future__ import annotations

import pytest

from src.core.msi_engine import MSIEngine, MSIEvaluationResult, evaluate_msi_branch


class TestF10MSIDecision:
    """Test suite for Feature F10: MSI & Fix Serial Decision Engine."""

    def test_f10_branch1_missing_in_plm_yields_ng(self) -> None:
        """Test 1: Branch 1 - Unit code missing in PLM yields 'NG'."""
        res = evaluate_msi_branch(
            in_plm=False,
            member_code="2NL",
            master_code="2NL",
            member_service="-",
            master_service="",
        )
        assert res["status"] == "NG"
        assert res["branch"] == 1
        assert "PLM" in res["reason"]

    def test_f10_branch4_missing_in_master_tool_yields_ng(self) -> None:
        """Test 2: Branch 4 - Unit in PLM but not found in Fix Serial master yields 'NG'."""
        res = evaluate_msi_branch(
            in_plm=True,
            member_code="2NL",
            master_code="",
            member_service="-",
            master_service="",
        )
        assert res["status"] == "NG"
        assert res["branch"] == 4
        assert "Fix Serial" in res["reason"]

    def test_f10_branch5_6_perfect_match_yields_ok(self) -> None:
        """Test 3: Branch 5 & 6 - Code matches and service matches yields 'OK'."""
        res1 = evaluate_msi_branch(
            in_plm=True,
            member_code="2NL",
            master_code="2NL",
            member_service="REPLACE ON JAM",
            master_service="REPLACE ON JAM",
        )
        assert res1["status"] == "OK"
        assert res1["branch"] == 5
        assert res1["service_warning"] is False

        res2 = evaluate_msi_branch(
            in_plm=True,
            member_code="1HN",
            master_code="1HN",
            member_service="",
            master_service="",
        )
        assert res2["status"] == "OK"
        assert res2["branch"] == 6
        assert res2["service_warning"] is False

    def test_f10_branch7_service_note_warning_yields_ok(self) -> None:
        """Test 4: Branch 7 - Code matches, but CTTT missed required Service note."""
        res = evaluate_msi_branch(
            in_plm=True,
            member_code="2NL",
            master_code="2NL",
            member_service="-",
            master_service="LUBRICATE EVERY 50K",
        )
        assert res["status"] == "OK"
        assert res["branch"] == 7
        assert res["service_warning"] is True

    def test_f10_branch8_9_code_or_service_mismatch_yields_ng(self) -> None:
        """Test 5: Branch 8 (code mismatch) and Branch 9 (service mismatch) yield 'NG'."""
        res8 = evaluate_msi_branch(
            in_plm=True,
            member_code="2NL",
            master_code="3NL",
            member_service="-",
            master_service="-",
        )
        assert res8["status"] == "NG"
        assert res8["branch"] == 8

        res9 = evaluate_msi_branch(
            in_plm=True,
            member_code="2NL",
            master_code="2NL",
            member_service="WRONG NOTE",
            master_service="CORRECT NOTE",
        )
        assert res9["status"] == "NG"
        assert res9["branch"] == 9
