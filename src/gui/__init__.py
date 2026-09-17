"""PyQt6 Desktop Graphical User Interface Package.

Provides:
- SSBOMMainWindow: Main desktop window hosting Leader & Member workspaces.
- LeaderWorkspaceView: Leader management, folder generation, batch processing, and reporting.
- MemberWorkspaceView: Production member data entry, preliminary self-check, and submission.
- SettingsDialog: System configuration dialog for TC14, SAP R3, and project directories.
"""

from src.gui.app import SSBOMMainWindow, main
from src.gui.leader_view import (
    BatchReconciliationWorker,
    EmailPreviewDialog,
    LeaderWorkspaceView,
)
from src.gui.member_view import MemberWorkspaceView
from src.gui.settings_dialog import (
    DEFAULT_SETTINGS,
    SettingsDialog,
)

__all__ = [
    "DEFAULT_SETTINGS",
    "BatchReconciliationWorker",
    "EmailPreviewDialog",
    "LeaderWorkspaceView",
    "MemberWorkspaceView",
    "SSBOMMainWindow",
    "SettingsDialog",
    "main",
]
