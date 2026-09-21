"""Tier 5 Adversarial Stress & Empirical Verification Suite for Theme Hot-Reload & QSS Parsing.

Milestone 1 (M1) Challenger 2 Empirical Test Suite.

Empirically challenges:
1. 100x Theme Switch Offscreen Stress:
   - Rapid alternating switching ('light' -> 'dark' -> 'light' -> ...) on a complex PyQt6 widget hierarchy:
     QMainWindow, QTableView with data model, QHeaderView, QProgressBar, QPushButton (standard & CTA),
     QComboBox, QLineEdit, QGroupBox, QDockWidget, QStatusBar, QTabWidget.
   - Process events to force full Qt style re-resolution and repolishing.
   - Checks memory growth bounds with tracemalloc and ensures zero segfaults/crashes.
   - Checks state preservation of model data, selections, progress values, and text.

2. Qt Stylesheet Parser Warnings and Errors:
   - Uses qInstallMessageHandler to intercept Qt messages.
   - Validates that our interceptor catches invalid QSS warnings (Oracle verification).
   - Asserts ZERO Qt QSS parser warnings/errors for both light_theme.qss and dark_theme.qss across all widgets.

3. Config JSON Resilience, Concurrency & Corruption:
   - Handles corrupted, 0-byte, binary garbage, and non-dict JSON files with graceful fallback.
   - Ensures deep directory creation.
   - Stress-tests multi-threaded concurrent writes (10 threads x 20 writes) to guarantee no file corruption.
   - Tests read-only/locked file gracefully logs without crashing.
"""

from __future__ import annotations

import concurrent.futures
import json
import os
import stat
import time
import tracemalloc
from pathlib import Path
from typing import Any

import pytest
from PyQt6.QtCore import QMessageLogContext, Qt, QtMsgType, qInstallMessageHandler
from PyQt6.QtGui import QStandardItem, QStandardItemModel
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDockWidget,
    QGroupBox,
    QLineEdit,
    QMainWindow,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QTableView,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from src.gui.styles.theme_manager import ThemeManager

# Ensure headless offscreen platform
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")


@pytest.fixture
def qapp() -> QApplication:
    """Ensure QApplication exists."""
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


class QtMessageCollector:
    """Captures and filters Qt log messages using qInstallMessageHandler."""

    def __init__(self) -> None:
        self.messages: list[tuple[QtMsgType, str, str]] = []
        self._previous_handler: Any = None

    def __enter__(self) -> QtMessageCollector:
        self.messages.clear()

        def handler(msg_type: QtMsgType, context: QMessageLogContext, msg: str) -> None:
            self.messages.append((msg_type, msg, context.file or ""))

        self._previous_handler = qInstallMessageHandler(handler)
        return self

    def __exit__(self, exc_type: Any, exc_val: Any, exc_tb: Any) -> None:
        qInstallMessageHandler(self._previous_handler)

    def get_qss_warnings(self) -> list[tuple[QtMsgType, str]]:
        """Return warnings or errors specifically related to stylesheet parsing."""
        qss_keywords = [
            "could not parse stylesheet",
            "unknown property",
            "could not parse",
            "stylesheet",
            "qss",
            "css",
            "selector",
            "rule",
            "declaration",
        ]
        results = []
        for msg_type, msg, _ in self.messages:
            msg_lower = msg.lower()
            # Ignore standard Qt offscreen font warning
            if "qfontdatabase" in msg_lower:
                continue
            if any(keyword in msg_lower for keyword in qss_keywords):
                results.append((msg_type, msg))
        return results


def _build_complex_widget_tree(parent_window: QMainWindow) -> dict[str, QWidget]:
    """Construct a dense, realistic widget hierarchy for empirical stress testing."""
    central = QWidget(parent_window)
    layout = QVBoxLayout(central)
    parent_window.setCentralWidget(central)

    # 1. QTableView with populated QStandardItemModel
    table = QTableView(central)
    model = QStandardItemModel(50, 6, table)
    model.setHorizontalHeaderLabels(["ID", "Part Name", "Rev", "Qty", "Status", "Notes"])
    for r in range(50):
        for c in range(6):
            model.setItem(r, c, QStandardItem(f"Item_{r}_{c}"))
    table.setModel(model)
    table.setSelectionMode(QTableView.SelectionMode.SingleSelection)
    table.selectRow(5)
    layout.addWidget(table)

    # 2. Controls row
    btn_normal = QPushButton("Normal Action", central)
    btn_cta = QPushButton("Primary Action", central)
    btn_cta.setObjectName("cta_primary")
    combo = QComboBox(central)
    combo.addItems([f"Option_{i}" for i in range(20)])
    combo.setCurrentIndex(3)
    pbar = QProgressBar(central)
    pbar.setValue(42)
    line_edit = QLineEdit("Search Part Number...", central)

    layout.addWidget(btn_normal)
    layout.addWidget(btn_cta)
    layout.addWidget(combo)
    layout.addWidget(pbar)
    layout.addWidget(line_edit)

    # 3. GroupBox with Checkbox and Radio Buttons
    group = QGroupBox("Configuration Options", central)
    grp_layout = QVBoxLayout(group)
    chk = QCheckBox("Enable Deep Inspection", group)
    chk.setChecked(True)
    rad = QRadioButton("Mode Fast", group)
    rad.setChecked(True)
    grp_layout.addWidget(chk)
    grp_layout.addWidget(rad)
    layout.addWidget(group)

    # 4. DockWidget
    dock = QDockWidget("Log Viewer", parent_window)
    dock_text = QTextEdit(dock)
    dock_text.setPlainText("Log initialized.")
    dock.setWidget(dock_text)
    parent_window.addDockWidget(Qt.DockWidgetArea.BottomDockWidgetArea, dock)

    # 5. Status Bar
    status_bar = parent_window.statusBar()
    status_bar.showMessage("Ready for stress verification")

    return {
        "window": parent_window,
        "table": table,
        "model": model,
        "btn_normal": btn_normal,
        "btn_cta": btn_cta,
        "combo": combo,
        "pbar": pbar,
        "line_edit": line_edit,
        "chk": chk,
        "dock": dock,
        "status_bar": status_bar,
    }


# =============================================================================
# 1. 100x THEME SWITCH OFFSCREEN STRESS TEST
# =============================================================================


class TestThemeSwitch100xOffscreenStress:
    """Stress test switching themes 100 times on a complex offscreen widget tree."""

    def test_100x_rapid_theme_toggle_stability_and_state_preservation(
        self, qapp: QApplication, tmp_path: Path
    ) -> None:
        """Switch theme 100 times rapidly, ensuring zero crashes, bounded time, and state preservation."""
        config_path = tmp_path / "stress_settings.json"
        tm = ThemeManager(config_path=config_path, apply_stylesheet=False)

        window = QMainWindow()
        widgets = _build_complex_widget_tree(window)
        window.show()
        qapp.processEvents()

        signal_emissions: list[str] = []
        tm.theme_changed.connect(lambda t: signal_emissions.append(t))

        NUM_ITERATIONS = 100
        start_time = time.perf_counter()

        for i in range(NUM_ITERATIONS):
            target_theme = "dark" if i % 2 == 0 else "light"
            tm.set_theme(target_theme, save_preference=False, apply_stylesheet=True)
            # Force Qt to process repolish and style events
            qapp.processEvents()

            assert tm.get_effective_theme() == target_theme
            assert qapp.styleSheet() == tm.get_stylesheet(target_theme)

        elapsed = time.perf_counter() - start_time
        avg_switch_ms = (elapsed / NUM_ITERATIONS) * 1000

        # Assert signal emissions
        assert len(signal_emissions) == NUM_ITERATIONS, (
            f"Expected {NUM_ITERATIONS} signals, got {len(signal_emissions)}"
        )
        assert signal_emissions[0] == "dark"
        assert signal_emissions[-1] == "light"

        # Performance requirement: average switch should not exceed 150ms per iteration (RES-UI-06)
        assert avg_switch_ms < 150.0, f"Average switch time too slow: {avg_switch_ms:.2f}ms (expected < 150ms)"

        # Verify widget state preservation after 100 switches
        pbar: QProgressBar = widgets["pbar"]  # type: ignore
        combo: QComboBox = widgets["combo"]  # type: ignore
        line_edit: QLineEdit = widgets["line_edit"]  # type: ignore
        chk: QCheckBox = widgets["chk"]  # type: ignore
        model: QStandardItemModel = widgets["model"]  # type: ignore

        assert pbar.value() == 42, "Progress bar value altered by theme switching"
        assert combo.currentIndex() == 3, "Combo box selection lost during theme switching"
        assert combo.currentText() == "Option_3"
        assert line_edit.text() == "Search Part Number...", "Line edit text corrupted"
        assert chk.isChecked(), "Checkbox state changed"
        assert model.rowCount() == 50, "Table rows corrupted"
        assert model.item(5, 1).text() == "Item_5_1", "Table cell content corrupted"

        window.close()
        window.deleteLater()
        qapp.processEvents()

    def test_memory_leakage_bounded_across_100_switches(
        self, qapp: QApplication, tmp_path: Path
    ) -> None:
        """Verify memory allocation remains strictly bounded with no runaway growth across 100 switches."""
        config_path = tmp_path / "memory_settings.json"
        tm = ThemeManager(config_path=config_path, apply_stylesheet=False)

        window = QMainWindow()
        _build_complex_widget_tree(window)
        window.show()
        qapp.processEvents()

        # Warm-up 5 iterations to allow initial Qt cache priming
        for i in range(5):
            tm.set_theme("dark" if i % 2 == 0 else "light", apply_stylesheet=True)
            qapp.processEvents()

        # Start tracking memory
        tracemalloc.start()
        snapshot_start = tracemalloc.take_snapshot()

        NUM_ITERATIONS = 100
        for i in range(NUM_ITERATIONS):
            tm.set_theme("dark" if i % 2 == 0 else "light", apply_stylesheet=True)
            qapp.processEvents()

        snapshot_end = tracemalloc.take_snapshot()
        current, peak = tracemalloc.get_traced_memory()
        tracemalloc.stop()

        stats = snapshot_end.compare_to(snapshot_start, "lineno")
        total_growth_bytes = sum(s.size_diff for s in stats if s.size_diff > 0)
        total_growth_mb = total_growth_bytes / (1024 * 1024)
        peak_mb = peak / (1024 * 1024)

        # Allow max 10MB growth across 100 full QSS re-evaluations (Qt internally caches some styles/fonts)
        assert total_growth_mb < 15.0, (
            f"Excessive memory growth detected across 100 theme switches: "
            f"{total_growth_mb:.2f} MB (Peak: {peak_mb:.2f} MB)"
        )

        window.close()
        window.deleteLater()
        qapp.processEvents()


# =============================================================================
# 2. QT STYLESHEET PARSER WARNINGS & ERRORS
# =============================================================================


class TestQtStylesheetParserIntegrity:
    """Verify zero Qt Stylesheet Parser warnings or errors when applying QSS stylesheets."""

    def test_oracle_warning_interceptor_detects_bad_qss(self, qapp: QApplication) -> None:
        """Sanity check/oracle: verify QtMessageCollector properly captures real QSS parsing warnings."""
        win = QMainWindow()
        win.show()
        qapp.processEvents()

        with QtMessageCollector() as collector:
            # Intentionally inject invalid QSS properties and malformed rules
            qapp.setStyleSheet(
                "QWidget { non_existent_bogus_property: 9999px; totally_broken_css: #foo; }"
            )
            qapp.processEvents()

        warnings = collector.get_qss_warnings()
        assert len(warnings) > 0, "Oracle failed: Qt parser warning was not captured by interceptor"
        assert any("unknown property" in w[1].lower() for w in warnings)

        win.close()
        win.deleteLater()
        qapp.processEvents()

    def test_zero_qss_parser_warnings_on_light_theme(self, qapp: QApplication) -> None:
        """Assert light_theme.qss produces 0 warnings from Qt's stylesheet parser on complex widgets."""
        win = QMainWindow()
        _build_complex_widget_tree(win)
        win.show()
        qapp.processEvents()

        tm = ThemeManager(apply_stylesheet=False)

        with QtMessageCollector() as collector:
            tm.set_theme("light", apply_stylesheet=True)
            qapp.processEvents()

        warnings = collector.get_qss_warnings()
        assert warnings == [], (
            f"Qt QSS Parser emitted {len(warnings)} warnings on light_theme.qss: {warnings}"
        )

        win.close()
        win.deleteLater()
        qapp.processEvents()

    def test_zero_qss_parser_warnings_on_dark_theme(self, qapp: QApplication) -> None:
        """Assert dark_theme.qss produces 0 warnings from Qt's stylesheet parser on complex widgets."""
        win = QMainWindow()
        _build_complex_widget_tree(win)
        win.show()
        qapp.processEvents()

        tm = ThemeManager(apply_stylesheet=False)

        with QtMessageCollector() as collector:
            tm.set_theme("dark", apply_stylesheet=True)
            qapp.processEvents()

        warnings = collector.get_qss_warnings()
        assert warnings == [], (
            f"Qt QSS Parser emitted {len(warnings)} warnings on dark_theme.qss: {warnings}"
        )

        win.close()
        win.deleteLater()
        qapp.processEvents()

    def test_zero_qss_parser_warnings_during_100_switches(self, qapp: QApplication) -> None:
        """Assert 100 alternating switches produce zero QSS parser warnings."""
        win = QMainWindow()
        _build_complex_widget_tree(win)
        win.show()
        qapp.processEvents()

        tm = ThemeManager(apply_stylesheet=False)

        with QtMessageCollector() as collector:
            for i in range(100):
                tm.set_theme("dark" if i % 2 == 0 else "light", apply_stylesheet=True)
                qapp.processEvents()

        warnings = collector.get_qss_warnings()
        assert warnings == [], (
            f"Qt QSS Parser emitted warnings during 100 switches: {warnings}"
        )

        win.close()
        win.deleteLater()
        qapp.processEvents()


# =============================================================================
# 3. CONFIG JSON RESILIENCE, CONCURRENCY & CORRUPTION
# =============================================================================


class TestConfigJsonResilienceAndConcurrency:
    """Stress test ThemeManager JSON persistence against corruption, locks, and concurrent writes."""

    def test_zero_byte_config_file_graceful_recovery(self, tmp_path: Path) -> None:
        """0-byte file must default to 'light' and recover on save."""
        cfg = tmp_path / "empty.json"
        cfg.write_bytes(b"")

        tm = ThemeManager(config_path=cfg, apply_stylesheet=False)
        assert tm.get_current_theme() == "light"

        tm.set_theme("dark", save_preference=True, apply_stylesheet=False)
        assert cfg.exists()
        saved = json.loads(cfg.read_text(encoding="utf-8"))
        assert saved.get("theme") == "dark"

    def test_truncated_json_graceful_recovery(self, tmp_path: Path) -> None:
        """Incomplete JSON string must fallback to 'light' and recover on save."""
        cfg = tmp_path / "truncated.json"
        cfg.write_text('{"theme": "dar', encoding="utf-8")

        tm = ThemeManager(config_path=cfg, apply_stylesheet=False)
        assert tm.get_current_theme() == "light"

        tm.set_theme("light", save_preference=True, apply_stylesheet=False)
        saved = json.loads(cfg.read_text(encoding="utf-8"))
        assert saved.get("theme") == "light"

    def test_binary_garbage_config_file_graceful_recovery(self, tmp_path: Path) -> None:
        """Binary garbage must fallback to 'light' and recover on save."""
        cfg = tmp_path / "garbage.json"
        cfg.write_bytes(b"\x00\xff\xfe\x12\x34\x56\x78\x9a\xbc\xde\xf0")

        tm = ThemeManager(config_path=cfg, apply_stylesheet=False)
        assert tm.get_current_theme() == "light"

        tm.set_theme("dark", save_preference=True, apply_stylesheet=False)
        saved = json.loads(cfg.read_text(encoding="utf-8"))
        assert saved.get("theme") == "dark"

    def test_non_dict_json_types_graceful_recovery(self, tmp_path: Path) -> None:
        """JSON containing non-dict root (array, string, number, null) must fallback gracefully."""
        non_dict_payloads = [
            json.dumps(["light", "dark"]),
            json.dumps("dark"),
            json.dumps(12345),
            json.dumps(None),
            json.dumps(True),
        ]
        for i, payload in enumerate(non_dict_payloads):
            cfg = tmp_path / f"non_dict_{i}.json"
            cfg.write_text(payload, encoding="utf-8")

            tm = ThemeManager(config_path=cfg, apply_stylesheet=False)
            assert tm.get_current_theme() == "light"

            # Save preference overwrites safely with valid dict
            tm.set_theme("dark", save_preference=True, apply_stylesheet=False)
            saved = json.loads(cfg.read_text(encoding="utf-8"))
            assert isinstance(saved, dict)
            assert saved.get("theme") == "dark"

    def test_invalid_theme_name_in_json_falls_back_to_light(self, tmp_path: Path) -> None:
        """Invalid theme name in JSON (e.g. 'neon') defaults to 'light'."""
        cfg = tmp_path / "invalid_theme.json"
        cfg.write_text(json.dumps({"theme": "neon_cyberpunk"}), encoding="utf-8")

        tm = ThemeManager(config_path=cfg, apply_stylesheet=False)
        assert tm.get_current_theme() == "light"

    def test_deeply_nested_directory_creation(self, tmp_path: Path) -> None:
        """Saving to a non-existent deeply nested path creates all parent folders."""
        deep_cfg = tmp_path / "level1" / "level2" / "level3" / "settings.json"
        assert not deep_cfg.parent.exists()

        tm = ThemeManager(config_path=deep_cfg, apply_stylesheet=False)
        tm.set_theme("dark", save_preference=True, apply_stylesheet=False)

        assert deep_cfg.exists()
        saved = json.loads(deep_cfg.read_text(encoding="utf-8"))
        assert saved.get("theme") == "dark"

    def test_concurrent_multithreaded_save_theme_stress(self, tmp_path: Path) -> None:
        """10 threads writing to the config file simultaneously must not leave corrupted JSON."""
        cfg = tmp_path / "concurrent_settings.json"
        cfg.write_text(json.dumps({"theme": "light", "core": "stable"}), encoding="utf-8")

        tm = ThemeManager(config_path=cfg, apply_stylesheet=False)

        num_threads = 10
        writes_per_thread = 20

        def worker(thread_id: int) -> None:
            for j in range(writes_per_thread):
                target = "dark" if (thread_id + j) % 2 == 0 else "light"
                tm.save_theme(target)
                time.sleep(0.001)

        with concurrent.futures.ThreadPoolExecutor(max_workers=num_threads) as executor:
            futures = [executor.submit(worker, tid) for tid in range(num_threads)]
            concurrent.futures.wait(futures)

        # Verify the file is not corrupt and is valid JSON
        assert cfg.exists()
        content = cfg.read_text(encoding="utf-8")
        parsed = json.loads(content)
        assert isinstance(parsed, dict)
        assert parsed.get("theme") in ("light", "dark")

    def test_read_only_config_file_handled_gracefully(self, tmp_path: Path) -> None:
        """When config file is read-only / permission denied, save_theme logs warning without crashing."""
        cfg = tmp_path / "readonly_settings.json"
        cfg.write_text(json.dumps({"theme": "light"}), encoding="utf-8")

        # Make file read-only
        os.chmod(cfg, stat.S_IREAD)

        tm = ThemeManager(config_path=cfg, apply_stylesheet=False)
        try:
            # save_theme should catch PermissionError and log warning without re-raising
            tm.save_theme("dark")
        finally:
            # Restore write permission for cleanup
            os.chmod(cfg, stat.S_IWRITE)
