"""Progress Reporter subsystem for TC2412 automation pipeline.

Implements non-blocking 120-second heartbeat reports conforming to
Phase 8 (REP-01 / R8) of specs/SPEC_PLM_AUTO_DOWNLOAD.md and PROJECT.md:88-91.
"""

from __future__ import annotations

import threading
import time
from typing import Callable, Optional


class ProgressReporter:
    """Periodic progress reporter (default interval: 120s / 2 minutes).

    Conforms to PROJECT.md:88-91 and specs/SPEC_PLM_AUTO_DOWNLOAD.md Section 11:
    - Provides `update(phase_index, total_phases, phase_name, percent) -> str`.
    - Supports 5-argument callback:
      `(phase_index: int, total_phases: int, phase_name: str, percent: float, elapsed: float) -> None`
    - Supports string_callback for console/UI/logging integration.
    - Thread-safe background heartbeat timer.
    """

    PHASE_NAMES: dict[int, str] = {
        1: "Authentication & Keyring Persistence",
        2: "Search & Navigation",
        3: "Deep Tree Expansion (Level 7)",
        4: "Select All Rows & Open Menu",
        5: "Safe Export Trigger (Anti-Import)",
        6: "14-Column Configuration",
        7: "Run Export & File Verification",
        8: "Pipeline Finalized",
    }

    PHASE_PERCENTAGES: dict[int, float] = {
        1: 12.5,
        2: 25.0,
        3: 37.5,
        4: 50.0,
        5: 62.5,
        6: 75.0,
        7: 87.5,
        8: 100.0,
    }

    def __init__(
        self,
        callback: Optional[Callable[[int, int, str, float, float], None]] = None,
        string_callback: Optional[Callable[[str], None]] = None,
        interval_seconds: float = 120.0,
    ) -> None:
        self.callback = callback
        self.string_callback = string_callback
        self.interval_seconds = max(0.01, float(interval_seconds))

        self.start_time = time.time()
        self.current_phase = 1
        self.total_phases = 8
        self.current_name = self.PHASE_NAMES.get(1, "Authentication")
        self.current_percent = self.PHASE_PERCENTAGES.get(1, 12.5)
        self.running = False
        self._lock = threading.Lock()
        self._thread: Optional[threading.Thread] = None

    def start(self) -> ProgressReporter:
        """Start the background 2-minute heartbeat timer."""
        with self._lock:
            if not self.running:
                self.running = True
                self._thread = threading.Thread(
                    target=self._run_timer,
                    name="TC2412ProgressReporterHeartbeat",
                    daemon=True,
                )
                self._thread.start()
        return self

    def stop(self) -> None:
        """Stop the background heartbeat timer safely."""
        with self._lock:
            self.running = False
        if self._thread and self._thread.is_alive() and threading.current_thread() != self._thread:
            self._thread.join(timeout=2.0)

    def __enter__(self) -> ProgressReporter:
        return self.start()

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        self.stop()

    def update(
        self,
        phase_index: int,
        total_phases: int = 8,
        phase_name: Optional[str] = None,
        percent: Optional[float] = None,
    ) -> str:
        """Update progress state, notify callbacks, and return the formatted progress message.

        Args:
            phase_index: Current phase index (1-8).
            total_phases: Total phases count (default 8).
            phase_name: Optional human-readable name of phase.
            percent: Optional percentage complete (0.0 - 100.0).

        Returns:
            Formatted progress string: 'Phase [X/8] - [Name] - [XX.X%] - [MM:SS elapsed]'.
        """
        elapsed = time.time() - self.start_time
        name = phase_name or self.PHASE_NAMES.get(phase_index, f"Phase {phase_index}")
        pct = percent if percent is not None else self.PHASE_PERCENTAGES.get(phase_index, (phase_index / total_phases) * 100.0)

        with self._lock:
            self.current_phase = phase_index
            self.total_phases = total_phases
            self.current_name = name
            self.current_percent = pct

        minutes = int(elapsed // 60)
        seconds = int(elapsed % 60)
        msg = f"Phase [{phase_index}/{total_phases}] - {name} - {pct:.1f}% - {minutes:02d}:{seconds:02d} elapsed"

        # Trigger 5-argument callback
        if self.callback:
            try:
                self.callback(phase_index, total_phases, name, pct, elapsed)
            except Exception:
                pass

        # Trigger string callback
        if self.string_callback:
            try:
                self.string_callback(msg)
            except Exception:
                pass

        return msg

    def next_phase(self, phase_name: Optional[str] = None, percent: Optional[float] = None) -> str:
        """Advance to the next phase sequentially."""
        with self._lock:
            next_idx = min(self.current_phase + 1, self.total_phases)
        return self.update(next_idx, self.total_phases, phase_name, percent)

    def _emit_heartbeat(self) -> None:
        """Emit periodic heartbeat message."""
        with self._lock:
            phase = self.current_phase
            total = self.total_phases
            name = self.current_name
            pct = self.current_percent
        self.update(phase, total, name, pct)

    def _run_timer(self) -> None:
        """Timer loop that triggers heartbeat updates."""
        while True:
            # Sleep in small slices to allow rapid shutdown
            slice_time = min(0.05, self.interval_seconds / 2.0)
            waited = 0.0
            while waited < self.interval_seconds:
                with self._lock:
                    if not self.running:
                        return
                time.sleep(slice_time)
                waited += slice_time

            with self._lock:
                if not self.running:
                    return
            self._emit_heartbeat()

