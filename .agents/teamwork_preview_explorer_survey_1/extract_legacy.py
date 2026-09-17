import os
import sys
from oletools.olevba import VBA_Parser

work_dir = r"D:\Sandbox\pm_sosanhbom"
out_dir = r"D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_1\legacy_vba"
os.makedirs(out_dir, exist_ok=True)

# 1. Read tudongdangnhapR3.vbs
vbs_path = os.path.join(work_dir, "tudongdangnhapR3.vbs")
if os.path.exists(vbs_path):
    with open(vbs_path, "rb") as f:
        content_bytes = f.read()
    # Try utf-16le, utf-8, cp1252
    for enc in ["utf-16", "utf-16-le", "utf-8", "cp1252", "latin1"]:
        try:
            text = content_bytes.decode(enc)
            with open(os.path.join(out_dir, "tudongdangnhapR3.vbs.txt"), "w", encoding="utf-8") as out_f:
                out_f.write(text)
            print(f"Decoded tudongdangnhapR3.vbs using {enc}, length={len(text)}")
            break
        except Exception:
            pass

# 2. Extract VBA from all .xlsm files
xlsm_files = [f for f in os.listdir(work_dir) if f.endswith(".xlsm")]
for fname in xlsm_files:
    fpath = os.path.join(work_dir, fname)
    print(f"\nProcessing {fname}...")
    try:
        vbaparser = VBA_Parser(fpath)
        if vbaparser.detect_vba_macros():
            file_vba_dir = os.path.join(out_dir, fname.replace(".", "_"))
            os.makedirs(file_vba_dir, exist_ok=True)
            for (filename, stream_path, vba_filename, vba_code) in vbaparser.extract_macros():
                safe_name = vba_filename.replace("/", "_").replace("\\", "_")
                macro_path = os.path.join(file_vba_dir, f"{safe_name}.vba")
                with open(macro_path, "w", encoding="utf-8", errors="replace") as mf:
                    mf.write(f"' Stream: {stream_path}\n' File: {vba_filename}\n\n")
                    mf.write(vba_code)
                print(f"  Extracted macro: {safe_name} ({len(vba_code)} bytes)")
        else:
            print(f"  No VBA macros detected in {fname}")
        vbaparser.close()
    except Exception as e:
        print(f"  Error extracting VBA from {fname}: {e}")

print("\nExtraction complete!")
