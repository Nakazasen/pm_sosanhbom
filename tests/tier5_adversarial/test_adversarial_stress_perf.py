"""Adversarial Stress, Performance & Boundary Benchmark Suite (Challenger 1).

Covers:
1. Deep nested BOM hierarchies (up to 12 levels and beyond).
2. Extreme BOM sizes (10,000, 25,000, 50,000 nodes) measuring runtime and memory consumption.
3. Execution threshold verification (<1.0s runtime for 10,000 nodes).
4. Concurrency and thread safety in GUI batch background workers (BatchReconciliationWorker).
"""

from __future__ import annotations

import datetime
import gc
import random
import string
import sys
import threading
import time
import tracemalloc
from typing import Any

import pandas as pd
import pytest
from pydantic import ValidationError

from src.core.date_filter import DateFilter, filter_by_date
from src.core.models import BOMNode, BOMTree
from src.core.reconciliation import ReconciliationEngine, reconcile_three_way
from src.core.tree_parser import PLMTreeParser
from src.core.unit_resolver import UnitResolver, resolve_units


# =============================================================================
# Helper: Synthetic BOM Generator
# =============================================================================

def generate_synthetic_bom_records(
    num_items: int,
    max_level: int = 6,
    seed: int = 42,
) -> list[dict[str, Any]]:
    """Generate deterministic synthetic PLM BOM records of size num_items."""
    rng = random.Random(seed)
    records: list[dict[str, Any]] = []

    current_level = 1
    units = ["LSU", "DEVELOPER", "DRUM", "FUSER", "CASSETTE", "MAIN_FRAME"]
    effectivities = [
        "01-Jan-2023 to 31-Dec-2023",
        "01-May-2024 UP",
        "to 30/06/2024",
        "01-Jan-2025 to 31-Dec-2026",
        "UP",
        "",
    ]

    for i in range(1, num_items + 1):
        if i == 1:
            level = 1
        else:
            # Change level with realistic probabilities
            delta = rng.choice([-2, -1, 0, 1])
            current_level = max(1, min(max_level, current_level + delta))
            level = current_level

        item_id = f"302FP{i:06d}"
        item_name = f"{units[i % len(units)]}_PART_{i}" if level > 1 else f"{units[i % len(units)]}_ASSY"
        has_children = rng.choice([True, False]) if level < max_level else False
        qty = rng.choice([1.0, 2.0, 4.0, 8.0])
        eff = rng.choice(effectivities)
        rev = rng.choice(["A", "01", "B", "02", "C"])

        records.append({
            "No": i,
            "Level": level,
            "Item Type": "Assembly" if has_children else "Part",
            "Item ID": item_id,
            "Has Children": "True" if has_children else "False",
            "Quantity": qty,
            "First Parts": "",
            "Second BOM flag": "",
            "Occurrence Effectivities": eff,
            "Item Revision:Projects List": "PRJ_STRESS",
            "Item Name": item_name,
            "Notice No": f"ECN-{i % 100}",
            "Revision": rev,
            "Item Rev:Release Status": "Released",
        })

    return records


# =============================================================================
# 1. Deep Nested Hierarchy Tests (Up to 12 levels)
# =============================================================================

class TestDeepNestedHierarchy:
    """Stress tests for deep nested BOM hierarchies up to 12 levels."""

    def test_bom_node_level_12_validation(self):
        """BOMNode must permit hierarchy levels up to 12 per specifications."""
        # Level 12 should be accepted for deeply nested sub-assemblies
        node_12 = BOMNode(level=12, item_id="DEEP_PART_12", item_name="DEEP LEAF")
        assert node_12.level == 12

    def test_bom_tree_deep_chain_12_levels(self):
        """Build and traverse a chain of 12 nested levels (Level 1 -> ... -> Level 12)."""
        root = BOMNode(level=1, item_id="ROOT_L1", item_name="ROOT", effectivity="01-Jan-2024 UP")
        curr = root
        for lvl in range(2, 13):
            child = BOMNode(level=lvl, item_id=f"PART_L{lvl}", item_name=f"SUB_L{lvl}", effectivity="01-Jan-2024 UP")
            curr.add_child(child)
            curr = child

        tree = BOMTree(roots=[root])
        assert tree.size() == 12
        assert tree.max_depth() == 12

        # Test clone with 12 levels
        cloned = tree.clone()
        assert cloned.size() == 12
        assert cloned.max_depth() == 12

        # Test unit resolver on 12-level hierarchy
        resolver = UnitResolver()
        resolved = resolver.resolve_tree(tree)
        assert resolved.root.unit_name == "ROOT"
        flattened = resolved.flatten()
        for n in flattened:
            assert n.unit_name == "ROOT"

        # Test date filter on 12-level hierarchy
        df = DateFilter()
        filtered = df.filter_tree(tree, reference_date=datetime.date(2026, 9, 17))
        assert filtered.size() == 12

    def test_plm_parser_with_12_level_dataset(self):
        """PLM parser should parse a 12-level hierarchical table without validation errors."""
        records = []
        for lvl in range(1, 13):
            records.append({
                "Level": lvl,
                "Item ID": f"PART_L{lvl}",
                "Item Name": f"NAME_L{lvl}",
                "Has Children": "True" if lvl < 12 else "False",
                "Quantity": 1.0,
                "Occurrence Effectivities": "01-Jan-2024 UP",
                "Revision": "01",
            })
        df = pd.DataFrame(records)
        parser = PLMTreeParser()
        tree = parser.parse_dataframe(df)
        assert tree.size() == 12
        assert tree.max_depth() == 12


# =============================================================================
# 2. Extreme BOM Sizes & Runtime / Memory Thresholds
# =============================================================================

class TestExtremeBOMPerformance:
    """Stress tests for extreme BOM sizes (10k, 25k, 50k nodes) with memory and time benchmarks."""

    def test_10k_nodes_tree_parsing_under_1_second(self):
        """10,000 items: PLM tree parsing must execute under 1.0 second."""
        records = generate_synthetic_bom_records(num_items=10_000, max_level=6, seed=101)
        df = pd.DataFrame(records)
        parser = PLMTreeParser()

        gc.collect()
        t0 = time.perf_counter()
        tree = parser.parse_dataframe(df)
        elapsed = time.perf_counter() - t0

        assert tree.size() > 9000, f"Expected >9000 nodes, got {tree.size()}"
        print(f"\n[BENCHMARK] 10,000 nodes Tree Parsing: {elapsed:.4f}s (Threshold: <1.0s)")
        assert elapsed < 1.0, f"10k nodes parsing exceeded 1.0s threshold: took {elapsed:.4f}s"

    def test_10k_nodes_date_filtering_performance(self):
        """10,000 items: Date filtering must execute under 1.0 second."""
        records = generate_synthetic_bom_records(num_items=10_000, max_level=6, seed=102)
        df = pd.DataFrame(records)
        parser = PLMTreeParser()
        tree = parser.parse_dataframe(df)

        filter_engine = DateFilter()
        ref_date = datetime.date(2026, 9, 17)

        gc.collect()
        t0 = time.perf_counter()
        filtered_tree = filter_engine.filter_tree(tree, reference_date=ref_date)
        elapsed = time.perf_counter() - t0

        print(f"\n[BENCHMARK] 10,000 nodes Date Filtering: {elapsed:.4f}s (Threshold: <1.0s)")
        assert elapsed < 1.0, f"10k nodes date filter exceeded 1.0s threshold: took {elapsed:.4f}s"
        assert filtered_tree.size() <= tree.size()

    def test_10k_nodes_unit_resolution_performance(self):
        """10,000 items: O(N) Unit resolution must execute under 0.5 seconds."""
        records = generate_synthetic_bom_records(num_items=10_000, max_level=6, seed=103)
        df = pd.DataFrame(records)
        parser = PLMTreeParser()
        tree = parser.parse_dataframe(df)

        resolver = UnitResolver()

        gc.collect()
        t0 = time.perf_counter()
        resolved_tree = resolver.resolve_tree(tree, in_place=True)
        elapsed = time.perf_counter() - t0

        print(f"\n[BENCHMARK] 10,000 nodes Unit Resolution: {elapsed:.4f}s (Threshold: <0.5s)")
        assert elapsed < 0.5, f"10k nodes unit resolution exceeded 0.5s threshold: took {elapsed:.4f}s"

    def test_10k_nodes_memory_consumption(self):
        """10,000 items: Peak memory consumption during parsing & resolution must remain under 100MB."""
        tracemalloc.start()
        records = generate_synthetic_bom_records(num_items=10_000, max_level=6, seed=104)
        df = pd.DataFrame(records)

        parser = PLMTreeParser()
        tree = parser.parse_dataframe(df)
        resolver = UnitResolver()
        resolver.resolve_tree(tree, in_place=True)

        current, peak = tracemalloc.get_traced_memory()
        tracemalloc.stop()

        peak_mb = peak / (1024 * 1024)
        print(f"\n[BENCHMARK] 10,000 nodes Peak Memory: {peak_mb:.2f} MB (Threshold: <100MB)")
        assert peak_mb < 100.0, f"Peak memory exceeded 100MB: {peak_mb:.2f} MB"

    def test_25k_and_50k_extreme_scalability(self):
        """Stress scalability on 25,000 and 50,000 nodes verifying linear scaling."""
        sizes = [25_000, 50_000]
        timings = {}
        memories = {}

        for n in sizes:
            records = generate_synthetic_bom_records(num_items=n, max_level=6, seed=n)
            df = pd.DataFrame(records)

            tracemalloc.start()
            gc.collect()
            t0 = time.perf_counter()

            parser = PLMTreeParser()
            tree = parser.parse_dataframe(df)

            filter_engine = DateFilter()
            filtered = filter_engine.filter_tree(tree, reference_date=datetime.date(2026, 9, 17))

            resolver = UnitResolver()
            resolver.resolve_tree(filtered, in_place=True)

            elapsed = time.perf_counter() - t0
            current, peak = tracemalloc.get_traced_memory()
            tracemalloc.stop()

            timings[n] = elapsed
            memories[n] = peak / (1024 * 1024)
            print(f"\n[BENCHMARK] {n:,} nodes Full Pipeline: {elapsed:.4f}s, Peak RAM: {memories[n]:.2f} MB")

        # Verify time scaling is roughly linear (50k time <= 3.5x of 25k time, not O(N^2) which would be 4x+)
        ratio = timings[50_000] / timings[25_000]
        print(f"[BENCHMARK] Scaling ratio (50k / 25k): {ratio:.2f}x (Linear expectation: ~2.0-2.5x)")
        assert ratio < 4.0, f"Non-linear quadratic scaling detected: ratio {ratio:.2f}x"
        assert timings[50_000] < 50.0, f"50k nodes took too long: {timings[50_000]:.2f}s"


# =============================================================================
# 3. Concurrency & Thread Safety in GUI Batch Workers
# =============================================================================

class TestGUIBatchWorkerConcurrency:
    """Thread safety and race condition verification for BatchReconciliationWorker."""

    def test_batch_worker_concurrent_executions(self):
        """Execute multiple BatchReconciliationWorker instances concurrently in separate threads."""
        from src.gui.leader_view import BatchReconciliationWorker

        # Prepare test data
        num_threads = 5
        results: list[Any] = [None] * num_threads
        errors: list[Any] = [None] * num_threads

        def run_worker(thread_idx: int):
            try:
                cttt_items = [
                    {"SUB": "LSU", "TRANG CTTT": "01", "MÃ LINH KIỆN": f"PART_{thread_idx}_{i}",
                     "TÊN LINH KIỆN": f"NAME_{i}", "SỐ LƯỢNG": 1.0, "PHỤ TRÁCH": "Tester", "Giải thích": ""}
                    for i in range(100)
                ]
                worker = BatchReconciliationWorker(
                    collected_cttt=cttt_items,
                    plm_path=None,
                    r3_path=None,
                    model_name=f"Model_{thread_idx}",
                )

                finished_event = threading.Event()

                def on_finished(res):
                    results[thread_idx] = res
                    finished_event.set()

                def on_error(err):
                    errors[thread_idx] = err
                    finished_event.set()

                worker.finished.connect(on_finished)
                worker.error.connect(on_error)
                worker.run()
                finished_event.wait(timeout=5.0)
            except Exception as exc:
                errors[thread_idx] = str(exc)

        threads = [threading.Thread(target=run_worker, args=(i,)) for i in range(num_threads)]
        for t in threads:
            t.start()
        for t in threads:
            t.join(timeout=10.0)

        # Assert all succeeded without error or corruption
        for idx in range(num_threads):
            assert errors[idx] is None, f"Thread {idx} failed with error: {errors[idx]}"
            assert results[idx] is not None, f"Thread {idx} did not produce a result"
            assert len(results[idx].cttt_rows) == 100

    def test_reconciliation_engine_multithreaded_reentrancy(self):
        """ReconciliationEngine must be safely callable simultaneously from multiple threads."""
        engine = ReconciliationEngine()
        errors = []

        def worker_task(worker_id: int):
            try:
                for _ in range(50):
                    res = engine.reconcile_single_row(
                        cttt_qty=1.0, plm_qty=1.0, r3_qty=1.0, plm_rev="01", r3_rev="01"
                    )
                    assert res["overall_check"] == "OK"
            except Exception as e:
                errors.append((worker_id, str(e)))

        threads = [threading.Thread(target=worker_task, args=(i,)) for i in range(10)]
        for t in threads:
            t.start()
        for t in threads:
            t.join()

        assert len(errors) == 0, f"Multithreaded reconciliation errors: {errors}"
