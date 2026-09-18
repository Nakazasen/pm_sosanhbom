"""Format all downloaded BOMs in Virgo2 to match exact 24-column structure of T10C423NL0 Mới.xlsm."""

import os
import sys
from pathlib import Path

sys.path.insert(0, os.path.abspath("."))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

import numpy as np
import pandas as pd

DEST_DIR = Path(
    r"\\fstvn01\Data\10_Production Engineering Department(製造技術部)\02.製造技術課\PE Dept\4A. QUAN LY BOM-TDTK-BOM管理-設計変更\SO SANH PLM-CTTT-R3\Virgo2"
)


def format_as_standard_24col(df: pd.DataFrame) -> pd.DataFrame:
    """Formats DataFrame to match the exact 24 columns of T10C423NL0 Mới.xlsm."""
    out = pd.DataFrame()
    out["Level"] = np.nan
    out["Level.1"] = df["Level"] if "Level" in df else df.get("Level.1", np.nan)
    out["Item Type"] = "Parts"
    out["Name"] = df["Revision Name"] if "Revision Name" in df else df.get("Name", df.get("ID", np.nan))
    out["1st parts"] = np.nan
    out["Quantity"] = df.get("Quantity", np.nan)
    out["2nd BOM Flag"] = False
    out["Parts Text"] = df.get("Description", df.get("Parts Text", np.nan))
    out["Notice No"] = np.nan
    out["Revision"] = df.get("Revision", np.nan)
    out["Release Status"] = df.get("Release Status", np.nan)
    out["Unit Of Measure"] = df.get("Unit Of Measure", np.nan)
    out["Date Released"] = df.get("Date Released", np.nan)
    out["Is Variant Item"] = False
    out["Assembly Indicator"] = df.get("Assembly Indicator", np.nan)
    out["Parts Text.1"] = np.nan
    out["Element Effectivities"] = df.get("Release Effectivity", df.get("Element Effectivities", np.nan))
    out["Occurrence Name"] = np.nan
    out["Finish Type"] = np.nan
    out["Safety Part"] = False
    out["Multibody"] = False
    out["Spare Part"] = True
    out["Technology Intent"] = "In-house Design"
    out["TCUID"] = df.get("TCUID", np.nan)
    return out


def main():
    files = sorted(
        [
            f
            for f in DEST_DIR.glob("PLM_*.xlsx")
            if not f.name.endswith(".bak") and "cũ" not in f.name.lower()
        ]
    )
    print(f"[*] Định dạng chuẩn 24 cột cho {len(files)} tệp BOM đã tải...", flush=True)

    for f in files:
        try:
            df = pd.read_excel(f)
            # If already 24 columns, check headers
            if len(df.columns) == 24 and "Level.1" in df.columns and "Name" in df.columns:
                print(f"[+] {f.name} đã là chuẩn 24 cột ({len(df)} dòng). Bỏ qua.", flush=True)
                continue

            df_24 = format_as_standard_24col(df)
            df_24.to_excel(f, index=False)
            print(f"[OK] Đã chuẩn hoá 24 cột cho: {f.name} ({len(df_24)} dòng, {len(df_24.columns)} cột)", flush=True)
        except Exception as e:
            print(f"[-] Lỗi xử lý {f.name}: {e}", flush=True)


if __name__ == "__main__":
    main()
