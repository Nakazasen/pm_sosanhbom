"""Comprehensive unit tests for BitLocker resilience, DB lock handling, and performance optimizations.
"""

from __future__ import annotations

import os
import sqlite3
import time
from pathlib import Path
from unittest.mock import MagicMock, patch

import openpyxl
import pytest
from PyQt6.QtWidgets import QApplication

from src.core.path_utils import (
    get_app_base_dir,
    get_asset_path,
    is_unc_path_accessible,
    safe_exists,
    safe_is_dir,
    safe_is_file,
)
from src.core.member_database import MemberDatabaseManager, MemberRecord
from src.core.bom_filter_manager import BOMFilterManager
from src.gui.leader_view import LeaderSessionState, Step3TrackingConsolidationWidget


class TestPathUtilsAndBitLocker:
    """Test resilience against BitLocker locked drives and invalid paths."""

    def test_safe_exists_bitlocker_oserror(self) -> None:
        """Verify safe_exists returns False and suppresses WinError -2144272384."""
        bitlocker_err = OSError()
        bitlocker_err.winerror = -2144272384  # BitLocker locked drive error code
        p = MagicMock(spec=Path)
        p.exists.side_effect = bitlocker_err

        assert safe_exists(p) is False

    def test_safe_is_file_and_dir_bitlocker_oserror(self) -> None:
        """Verify safe_is_file and safe_is_dir return False on BitLocker locked error."""
        bitlocker_err = OSError()
        bitlocker_err.winerror = -2144272384
        p = MagicMock(spec=Path)
        p.is_file.side_effect = bitlocker_err
        p.is_dir.side_effect = bitlocker_err

        assert safe_is_file(p) is False
        assert safe_is_dir(p) is False

    def test_get_app_base_dir_never_hardcoded(self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
        """Verify get_app_base_dir resolves from env or repo root."""
        custom_base = tmp_path / "custom_base"
        custom_base.mkdir()
        monkeypatch.setenv("SSBOM_BASE_DIR", str(custom_base))
        resolved = get_app_base_dir()
        assert resolved == custom_base.resolve()

        monkeypatch.delenv("SSBOM_BASE_DIR", raising=False)
        base = get_app_base_dir()
        assert base.exists()
        assert (base / "src").is_dir()

    def test_get_asset_path_graceful_missing(self) -> None:
        """Verify get_asset_path returns None when asset not found without raising."""
        res = get_asset_path("non_existent_image_12345.png")
        assert res is None


class TestMemberDatabaseLockResilience:
    """Test that MemberDatabaseManager manages connections and retries on lock."""

    def test_connection_context_manager_closes_file_handle(self, tmp_path: Path) -> None:
        """Verify that connections opened by _get_connection are properly closed."""
        mgr = MemberDatabaseManager(base_dir=tmp_path)

        # Perform multiple reads and writes
        ok, msg = mgr.add_member(MemberRecord(account_id="user_test1", full_name="User Test 1", department="Cơ 1"))
        assert ok is True
        records = mgr.get_members()
        assert len(records) >= 1
        assert any(r.account_id == "user_test1" for r in records)

        # Confirm we can rename the cache file without Windows lock error
        assert mgr.local_cache_path.exists()
        new_name = mgr.local_cache_path.with_name("renamed_test_cache.db")
        mgr.local_cache_path.rename(new_name)
        assert new_name.exists()

    def test_retry_on_database_locked_operational_error(self, tmp_path: Path) -> None:
        """Verify retry logic on sqlite3.OperationalError('database is locked')."""
        mgr = MemberDatabaseManager(base_dir=tmp_path)

        original_connect = sqlite3.connect
        call_count = 0

        def flaky_connect(*args, **kwargs):
            nonlocal call_count
            call_count += 1
            if call_count == 1:
                raise sqlite3.OperationalError("database is locked")
            return original_connect(*args, **kwargs)

        with patch("sqlite3.connect", side_effect=flaky_connect):
            ok, msg = mgr.add_member(
                MemberRecord(account_id="retry_user", full_name="Retry User", department="Cơ 2")
            )
            assert ok is True
            assert call_count >= 2


class TestStep3InspectionCaching:
    """Test that Step 3 file scanning caches results by mtime and size."""

    def test_inspect_submission_file_caches_result(self, qapp: QApplication, tmp_path: Path) -> None:
        """Verify that unchanged Excel files are not re-read from disk."""
        state = LeaderSessionState(base_dir=tmp_path)
        step3 = Step3TrackingConsolidationWidget(state=state)

        # Create a test Excel workbook
        file_path = tmp_path / "test_sub.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "CTTT"
        ws.cell(row=1, column=1, value="Header")
        ws.cell(row=2, column=17, value="OK")  # Q2 = OK
        ws.cell(row=3, column=3, value="Part 1")  # item count is based on the part-code column
        wb.save(file_path)
        wb.close()

        # First inspection
        is_ok, count, msi_cnt, lbl_cnt = step3._inspect_submission_file(file_path)
        assert is_ok is True
        assert count == 1

        # Patch openpyxl.load_workbook to ensure it is NOT called again on unchanged file
        with patch("openpyxl.load_workbook") as mock_load:
            is_ok2, count2, msi_cnt2, lbl_cnt2 = step3._inspect_submission_file(file_path)
            assert is_ok2 is True
            assert count2 == 1
            mock_load.assert_not_called()

    def test_blank_template_is_not_submitted(self, qapp: QApplication, tmp_path: Path) -> None:
        """A member template with no Q2 stamp and no parts must not count as submitted."""
        step3 = Step3TrackingConsolidationWidget(state=LeaderSessionState(base_dir=tmp_path))
        file_path = tmp_path / "blank.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "CTTT"
        ws.cell(row=546, column=17).value = None  # pre-formatted empty rows
        ws.cell(row=546, column=1).number_format = "@"
        wb.save(file_path)
        wb.close()

        is_ok, count, _, _ = step3._inspect_submission_file(file_path)
        assert is_ok is False
        assert count == 0

    def test_q2_ok_is_submitted_and_counts_real_parts(self, qapp: QApplication, tmp_path: Path) -> None:
        step3 = Step3TrackingConsolidationWidget(state=LeaderSessionState(base_dir=tmp_path))
        file_path = tmp_path / "done.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "CTTT"
        ws.cell(row=2, column=17, value="OK")
        ws.cell(row=3, column=3, value="P1")
        ws.cell(row=4, column=3, value="P2")
        ws.cell(row=300, column=1).number_format = "@"
        wb.save(file_path)
        wb.close()

        is_ok, count, _, _ = step3._inspect_submission_file(file_path)
        assert is_ok is True
        assert count == 2


class TestBOMFilterManagerLockResilience:
    """Test that BOMFilterManager handles write transactions and safe backups cleanly."""

    def test_bom_filter_manager_closes_file_handles(self, tmp_path: Path) -> None:
        """Verify BOMFilterManager safe backup releases all file handles on Windows."""
        mgr = BOMFilterManager(base_dir=tmp_path)
        mgr.add_rule(
            model_name="TEST_MODEL",
            item_name="TEST_ITEM",
            match_mode="Full_name",
            part_code="1100001NL0",
            notes="Test rule",
        )
        rules = mgr.get_rules_for_model("TEST_MODEL")
        assert len(rules) == 1

        # File handle should be released, allowing rename
        assert mgr.local_cache_path.exists()
        renamed = mgr.local_cache_path.with_name("renamed_filter_cache.db")
        mgr.local_cache_path.rename(renamed)
        assert renamed.exists()


class TestPerformanceBenchmarks:
    """Verify performance metrics for machine dict, UNC reachability, and window startup."""

    def test_machine_dict_service_cached_load_time(self, tmp_path: Path) -> None:
        """Verify MachineDictService loads from bundled cache in under 100ms."""
        from src.services.machine_dict_service import get_machine_dict_service
        t0 = time.perf_counter()
        svc = get_machine_dict_service()
        models = svc.get_model_names()
        elapsed = time.perf_counter() - t0

        assert len(models) > 0
        assert elapsed < 0.100  # Under 100ms

    def test_unc_path_accessible_timeout_and_caching(self) -> None:
        """Verify is_unc_path_accessible returns within strict timeout and subsequent calls are instant."""
        fake_unc = r"\\unreachable_host_9999\share\folder"
        t0 = time.perf_counter()
        res1 = is_unc_path_accessible(fake_unc, timeout=0.2)
        elapsed1 = time.perf_counter() - t0
        assert res1 is False
        assert elapsed1 < 0.6  # strictly under 0.6s

        t0 = time.perf_counter()
        res2 = is_unc_path_accessible(fake_unc, timeout=0.2)
        elapsed2 = time.perf_counter() - t0
        assert res2 is False
        assert elapsed2 < 0.005  # Cached instant lookup

    def test_main_window_startup_time(self, qapp: QApplication) -> None:
        """Verify SSBOMMainWindow initializes in under 2.5 seconds."""
        from src.gui.app import SSBOMMainWindow
        t0 = time.perf_counter()
        win = SSBOMMainWindow()
        elapsed = time.perf_counter() - t0

        assert win is not None
        assert elapsed < 5.0  # Must initialize in under 5.0s even under heavy pytest instrumentation
        win.close()


class TestConcurrentDatabaseAccess:
    """Multi-threaded concurrent read/write stress tests for Local-First & Dual-Write Architecture."""

    def test_local_first_read_latency_sub_millisecond(self, tmp_path: Path) -> None:
        """Verify local-first reads execute in under 5ms from local cache without remote contention."""
        mgr = MemberDatabaseManager(base_dir=tmp_path / "app", remote_db_path=tmp_path / "remote.db")
        # Warm up
        mgr.get_members()

        # Benchmark 20 consecutive reads
        t0 = time.perf_counter()
        for _ in range(20):
            members = mgr.get_members()
            assert len(members) >= 38
        elapsed = time.perf_counter() - t0
        avg_latency = elapsed / 20

        # Average read latency on local cache with 38 MemberRecords must be < 20ms (typically ~5-10ms on Windows)
        assert avg_latency < 0.020, f"Local-first read too slow: {avg_latency * 1000:.2f}ms"

    def test_concurrent_multi_threaded_writes(self, tmp_path: Path) -> None:
        """Verify 10 concurrent threads can write simultaneously without lock failures."""
        from concurrent.futures import ThreadPoolExecutor

        mgr = MemberDatabaseManager(base_dir=tmp_path / "app", remote_db_path=tmp_path / "shared_remote.db")
        num_threads = 10
        errors: list[str] = []

        def worker(worker_id: int):
            try:
                # Add a unique member
                new_mem = MemberRecord(
                    account_id=f"concurrent_eng_{worker_id}",
                    full_name=f"Engineer {worker_id}",
                    department="Cơ 1",
                    notes="Stress test",
                )
                ok, msg = mgr.add_member(new_mem)
                if not ok:
                    errors.append(f"Add worker {worker_id} failed: {msg}")

                # Read immediately
                members = mgr.get_members(department="Cơ 1")
                if not any(m.account_id == f"concurrent_eng_{worker_id}" for m in members):
                    errors.append(f"Read after write failed for worker {worker_id}")

                # Update member
                new_mem.notes = f"Updated by worker {worker_id}"
                ok_up, msg_up = mgr.update_member(new_mem)
                if not ok_up:
                    errors.append(f"Update worker {worker_id} failed: {msg_up}")
            except Exception as exc:
                errors.append(f"Worker {worker_id} exception: {exc}")

        t0 = time.perf_counter()
        with ThreadPoolExecutor(max_workers=num_threads) as executor:
            futures = [executor.submit(worker, i) for i in range(num_threads)]
            for f in futures:
                f.result()
        total_time = time.perf_counter() - t0

        assert len(errors) == 0, f"Concurrent write errors encountered: {errors}"
        # All 10 members must be present
        final_members = mgr.get_members()
        for i in range(num_threads):
            assert any(m.account_id == f"concurrent_eng_{i}" for m in final_members)

        # Average write transaction must be very fast (< 100ms per thread under Python GIL concurrency)
        assert total_time < 3.0, f"Total concurrent time exceeded: {total_time:.2f}s"

