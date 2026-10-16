"""Re-filter all Virgo2 full BOMs into 'down bom' after the Has Children fix.

Pipeline (same as the previous *_loc output): 24-col TC2412 export -> canonical 14-col
(electrical units removed, ACCESSORIES collapsed to L2) -> Virgo BolocBom rules.
Full backups are only read, never modified.
"""

from __future__ import annotations

import io
import logging
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
logging.disable(logging.WARNING)

from src.automation.tc2412.standardizer import standardize_plm_file  # noqa: E402
from src.core.model_pruner import ModelPruner  # noqa: E402
from src.core.tree_parser import PLMTreeParser  # noqa: E402

VIRGO2 = Path(
    r"\\fstvn01\Data\10_Production Engineering Department(製造技術部)\02.製造技術課\PE Dept"
    r"\4A. QUAN LY BOM-TDTK-BOM管理-設計変更\SO SANH PLM-CTTT-R3\Virgo2"
)
SRC_DIR = VIRGO2 / "down bom" / "backup_before_virgo2_filter"
OUT_DIR = VIRGO2 / "down bom"


def main() -> int:
    files = sorted(p for p in SRC_DIR.glob("PLM_*.xlsx") if not p.name.startswith("~$"))
    pruner, parser = ModelPruner(), PLMTreeParser()
    print(f"{len(files)} file full trong {SRC_DIR}")
    print(f"{'File':<22}{'Full':>7}{'14 cột':>8}{'Sau lọc':>9}{'FILL UP':>9}")
    failed = 0
    with tempfile.TemporaryDirectory() as tmp:
        for src in files:
            try:
                std = Path(tmp) / src.name
                full_rows, std_rows = standardize_plm_file(src, std, prune_electrical=True)
                out = OUT_DIR / src.name
                res = pruner.prune_file(std, out, model_name="Virgo", create_backup=False)
                fill = [n for n in parser.parse_file(out).flatten() if "FILL UP CONTAINER" in n.item_name.upper()]
                print(f"{src.name:<22}{full_rows:>7}{std_rows:>8}{res['final_rows']:>9}{len(fill):>9}")
            except Exception as exc:  # keep going for the remaining files
                failed += 1
                print(f"{src.name:<22} LỖI: {exc}")
    print(f"Hoàn tất: {len(files) - failed}/{len(files)} file -> {OUT_DIR}")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
