import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.stdout.reconfigure(encoding="utf-8")
from PyQt6.QtWidgets import QApplication

app = QApplication.instance() or QApplication(sys.argv)
from src.gui.app import SSBOMMainWindow

w = SSBOMMainWindow()
print("=== WINDOW TITLE ===", w.windowTitle())
print("=== TOOLBAR ACTION COUNT ===", len(w.toolbar.actions()))
for i, act in enumerate(w.toolbar.actions()):
    print(f"{i}: text='{act.text()}' | tooltip='{act.toolTip()}'")

print("=== FILE MENU ACTION COUNT ===", len(w.menuBar().actions()[0].menu().actions()))
for i, act in enumerate(w.menuBar().actions()[0].menu().actions()):
    print(f"{i}: text='{act.text()}' | shortcut='{act.shortcut().toString()}'")

w.close()
sys.exit(0)
