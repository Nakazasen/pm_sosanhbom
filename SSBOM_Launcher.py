"""Stable Root Launcher for SSBOM Manager.

Delegates execution to scripts.update_launcher.
"""

import sys
from pathlib import Path

# Ensure root is in sys.path
root_dir = Path(__file__).resolve().parent
if str(root_dir) not in sys.path:
    sys.path.insert(0, str(root_dir))

from scripts.update_launcher import main

if __name__ == "__main__":
    sys.exit(main())
