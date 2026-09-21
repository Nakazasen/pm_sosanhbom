from __future__ import annotations

import json
import shutil
import sys
from pathlib import Path
from typing import Any, Dict, Optional

# Ensure root is in sys.path
root_dir = Path(__file__).resolve().parent
if str(root_dir) not in sys.path:
    sys.path.insert(0, str(root_dir))

from scripts.update_launcher import main


class AppLauncher:
    """Launcher management helper matching MP2027 spec."""

    def __init__(self, root_dir: Optional[Path | str] = None) -> None:
        self.root_dir = Path(root_dir) if root_dir else Path(__file__).resolve().parent

    def read_current_pointer(self) -> Dict[str, Any]:
        cur_file = self.root_dir / "current.json"
        if cur_file.is_file():
            try:
                with open(cur_file, "r", encoding="utf-8") as fp:
                    return json.load(fp)
            except Exception:
                pass
        return {"is_dev": True, "entrypoint": "src.gui.app"}

    def rollback(self) -> bool:
        prev_file = self.root_dir / "previous.json"
        cur_file = self.root_dir / "current.json"
        if prev_file.is_file():
            shutil.copy2(prev_file, cur_file)
            return True
        return False


if __name__ == "__main__":
    sys.exit(main())
