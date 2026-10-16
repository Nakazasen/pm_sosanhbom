"""Batch prune 29 Spica PLM files with rules from BolocBom_Spica.xlsx."""

from __future__ import annotations

import sys
from pathlib import Path

# Add project root to sys.path
ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import pandas as pd
from src.core.model_pruner import ModelPruner, ModelRule

FILES = [
    r"D:\Sandbox\pm_sosanhbom\20261002\PLM_T10C362US0.xlsx",
    r"D:\Sandbox\pm_sosanhbom\20261002\PLM_T10C363NL0.xlsx",
    r"D:\Sandbox\pm_sosanhbom\20261002\PLM_T10C382US0.xlsx",
    r"D:\Sandbox\pm_sosanhbom\20261002\PLM_T10C383NL0.xlsx",
    r"D:\Sandbox\pm_sosanhbom\20261002\PLM_T10C392US0.xlsx",
    r"D:\Sandbox\pm_sosanhbom\20261002\PLM_T10C393NL0.xlsx",
    r"D:\Sandbox\pm_sosanhbom\20261002\PLM_T10C0L2US0.xlsx",
    r"D:\Sandbox\pm_sosanhbom\20261002\PLM_T10C0L3NL0.xlsx",
    r"D:\Sandbox\pm_sosanhbom\20261002\PLM_T10C0L9JP0.xlsx",
    r"D:\Sandbox\pm_sosanhbom\20261002\PLM_T10C0M3NL0.xlsx",
    r"D:\Sandbox\pm_sosanhbom\20261002\PLM_T10C0N3NL0.xlsx",
    r"D:\Sandbox\pm_sosanhbom\20261002\PLM_T10C0P3NL0.xlsx",
    r"D:\Sandbox\pm_sosanhbom\20261002\PLM_T10C0R2US0.xlsx",
    r"D:\Sandbox\pm_sosanhbom\20261002\PLM_T10C0R3NL0.xlsx",
    r"D:\Sandbox\pm_sosanhbom\20261002\PLM_T10C0R9JP0.xlsx",
    r"D:\Sandbox\pm_sosanhbom\20261002\PLM_T10C0S3NL0.xlsx",
    r"D:\Sandbox\pm_sosanhbom\20261002\PLM_T10C1T2US0.xlsx",
    r"D:\Sandbox\pm_sosanhbom\20261002\PLM_T10C1T3NL0.xlsx",
    r"D:\Sandbox\pm_sosanhbom\20261002\PLM_T10C1T9JP0.xlsx",
    r"D:\Sandbox\pm_sosanhbom\20261002\PLM_T10C1V3NL0.xlsx",
    r"D:\Sandbox\pm_sosanhbom\20261002\PLM_T10C1W3NL0.xlsx",
    r"D:\Sandbox\pm_sosanhbom\20261002\PLM_T10C1Y2US0.xlsx",
    r"D:\Sandbox\pm_sosanhbom\20261002\PLM_T10C1Y3NL0.xlsx",
    r"D:\Sandbox\pm_sosanhbom\20261002\PLM_T10C1Y9JP0.xlsx",
    r"D:\Sandbox\pm_sosanhbom\20261002\PLM_T10C1Z3NL0.xlsx",
    r"D:\Sandbox\pm_sosanhbom\20261002\PLM_T10C342US0.xlsx",
    r"D:\Sandbox\pm_sosanhbom\20261002\PLM_T10C343NL0.xlsx",
    r"D:\Sandbox\pm_sosanhbom\20261002\PLM_T10C352US0.xlsx",
    r"D:\Sandbox\pm_sosanhbom\20261002\PLM_T10C353NL0.xlsx",
]

def main():
    sys.stdout.reconfigure(encoding="utf-8")
    rule_file = r"D:\Sandbox\pm_sosanhbom\20261002\BolocBom_Spica.xlsx"
    df_rules = pd.read_excel(rule_file)
    rules: list[ModelRule] = []
    for _, r in df_rules.iterrows():
        name = str(r.iloc[1]).strip() if pd.notna(r.iloc[1]) else None
        mode = str(r.iloc[2]).strip() if pd.notna(r.iloc[2]) else "Full_name"
        code = str(r.iloc[3]).strip() if pd.notna(r.iloc[3]) else None
        rules.append(ModelRule(item_name=name, match_mode=mode, part_code=code))

    print(f"=== BẮT ĐẦU LỌC 29 FILE PLM SPICA ===")
    print(f"[*] Nạp quy tắc từ: {rule_file}")
    print(f"[*] Tổng số quy tắc Spica: {len(rules)}")
    print("-" * 75)

    pruner = ModelPruner(custom_rules={"Spica": rules})
    results = []

    for i, file_path_str in enumerate(FILES, start=1):
        p = Path(file_path_str)
        if not p.is_file():
            print(f"[{i:02d}/29] [-] TỆP KHÔNG TỒN TẠI: {p.name}")
            results.append({"file": p.name, "initial": 0, "final": 0, "deleted": 0, "status": "FILE_NOT_FOUND"})
            continue

        try:
            stats = pruner.prune_file(p, model_name="Spica", create_backup=True)
            ini = stats.get("initial_rows", 0)
            fin = stats.get("final_rows", 0)
            dele = stats.get("deleted_rows", 0)
            print(f"[{i:02d}/29] [✓] {p.name:<25} | Ban đầu: {ini:>4} | Còn lại: {fin:>4} | Đã xóa: {dele:>3} dòng")
            results.append({"file": p.name, "initial": ini, "final": fin, "deleted": dele, "status": "OK"})
        except Exception as exc:
            print(f"[{i:02d}/29] [-] LỖI {p.name}: {exc}")
            results.append({"file": p.name, "initial": 0, "final": 0, "deleted": 0, "status": f"ERROR: {exc}"})

    print("-" * 75)
    total_deleted = sum(r["deleted"] for r in results)
    ok_count = sum(1 for r in results if r["status"] == "OK")
    print(f"=== TỔNG KẾT: Hoàn thành {ok_count}/29 file | Tổng số dòng đã lọc xóa: {total_deleted} ===")

if __name__ == "__main__":
    main()
