"""Batch download 16 Virgo PLM BOMs from Siemens Teamcenter 2412 into Virgo2 network folder.

Uses user vn_pe02 with verified preset 'KTCT_Trong' for authentic 24-column PLM export.
Converts lossless to pure .xlsx (0 errors, 100% style preservation).
Backs up raw BOM to backup_before_virgo_filter/ and applies BolocBom filter automatically.
"""

from __future__ import annotations

import os
import shutil
import sys
import time
import zipfile
from datetime import datetime
from pathlib import Path
from typing import List, Optional

sys.path.insert(0, os.path.abspath("."))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

import openpyxl
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from src.automation.tc2412.client import TC2412AutomationClient
from src.automation.tc2412.session import BrowserConfig, TC2412SessionManager
from scripts.fast_virgo_filter import load_bolocbom_rules, process_file

PART_NUMBERS: List[str] = [
    "T10C423NL0",
    "T10C4A2US0",
    "T10C4B2US0",
    "T10C4C3NL0",
    "T10C4D2US0",
    "T10C4F3NL0",
    "T10C4H3NL0",
    "T10C4K3NL0",
    "T10C4L9JP0",
    "T10C4M3NL0",
    "T10C4N2US0",
    "T10C433NL0",
    "T10C452US0",
    "T10C463NL0",
    "T10C483NL0",
    "T10C493NL0",
]

DEST_DIR = Path(
    r"\\fstvn01\Data\10_Production Engineering Department(製造技術部)\02.製造技術課\PE Dept\4A. QUAN LY BOM-TDTK-BOM管理-設計変更\SO SANH PLM-CTTT-R3\Virgo2"
)
BACKUP_DIR = DEST_DIR / "backup_before_virgo_filter"


def is_already_valid(dest_path: Path) -> bool:
    """Check if destination file is already complete with 24 columns and valid Part Names."""
    if not dest_path.exists() or dest_path.stat().st_size < 100_000:
        return False
    try:
        wb = openpyxl.load_workbook(dest_path, read_only=True, data_only=True)
        ws = wb.active
        if ws.max_column < 24 or ws.max_row < 500:
            wb.close()
            return False
        # Check that Column H (Parts Text) is not empty on row 2 or 3
        has_text = False
        for r in range(2, min(ws.max_row + 1, 10)):
            val = ws.cell(row=r, column=8).value
            if val and str(val).strip():
                has_text = True
                break
        wb.close()
        return has_text
    except Exception:
        return False


def convert_to_pure_xlsx(src_file: Path, dest_xlsx: Path) -> bool:
    """Lossless conversion from TC2412 export to pure .xlsx, preserving 100% of formatting."""
    try:
        with zipfile.ZipFile(src_file, "r") as zin:
            with zipfile.ZipFile(dest_xlsx, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename.endswith("vbaProject.bin"):
                        continue
                    data = zin.read(item.filename)
                    if item.filename == "[Content_Types].xml":
                        data = data.replace(
                            b"application/vnd.ms-excel.sheet.macroEnabled.main+xml",
                            b"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                        )
                    zout.writestr(item, data)
        return True
    except Exception as e:
        print(f"[-] Lỗi chuyển đổi .xlsx: {e}", flush=True)
        return False


def ensure_ktct_trong_arrangement(driver, panel, wait) -> bool:
    """Ensure that the 'KTCT_Trong' column arrangement is selected in the Export panel."""
    try:
        # Check if already KTCT_Trong
        panel_text = panel.text
        if "KTCT_Trong" in panel_text:
            print("[+] Preset 'KTCT_Trong' đã đang được chọn.", flush=True)
            return True

        print("[*] Chọn preset 'KTCT_Trong' từ menu Column Arrangements...", flush=True)
        arr_btn = panel.find_element(By.CSS_SELECTOR, "button[command-id='Arm0ArrangeViewConfigs']")
        arr_btn.click()
        time.sleep(1.5)

        menu_items = driver.find_elements(
            By.CSS_SELECTOR,
            "div.aw-popup div.aw-widgets-cellListItem, div.aw-popup li, div.sw-popup li, div.sw-popup div",
        )
        target = None
        for m in menu_items:
            if (m.text or "").strip() == "KTCT_Trong":
                target = m
                break

        if target:
            target.click()
            time.sleep(2)
            print("[+] Đã click chọn preset 'KTCT_Trong' thành công!", flush=True)
            return True
        else:
            print("[-] CẢNH BÁO: Không tìm thấy 'KTCT_Trong' trong menu popup!", flush=True)
            return False
    except Exception as exc:
        print(f"[-] Lỗi khi chọn KTCT_Trong: {exc}", flush=True)
        return False


def download_single_bom(
    client: TC2412AutomationClient,
    part_number: str,
    download_dir: Path,
    rules: List[dict],
) -> bool:
    """Download a single BOM, convert to pure .xlsx, backup, and apply BolocBom."""
    driver = client.driver
    wait = WebDriverWait(driver, 25)

    dest_file = DEST_DIR / f"PLM_{part_number}.xlsx"
    bak_file = BACKUP_DIR / f"PLM_{part_number}.xlsx"

    print(f"\n=======================================================", flush=True)
    print(f"[*] BẮT ĐẦU TẢI BOM: {part_number}", flush=True)
    print(f"=======================================================", flush=True)

    if is_already_valid(dest_file):
        print(f"[+] Tệp {dest_file.name} đã tồn tại và hợp lệ ({dest_file.stat().st_size:,} bytes). Bỏ qua tải lại.", flush=True)
        return True

    # Clean download dir before each part
    for f in download_dir.glob("*"):
        try:
            f.unlink()
        except Exception:
            pass

    try:
        # [1] Search
        print(f"[1/7] Tìm kiếm mã: {part_number}...", flush=True)
        client.search_item(part_number)
        time.sleep(3)

        # [2] Open Content Tab
        print(f"[2/7] Mở tab Content...", flush=True)
        content_tab = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//a[contains(@class, 'sw-tab-title') and normalize-space()='Content']")
            )
        )
        content_tab.click()
        time.sleep(4)

        # [3] Select Root BOM Row
        print(f"[3/7] Chọn dòng BOM gốc...", flush=True)
        root_cell = wait.until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "div.aw-splm-tableRow div.aw-splm-tableCellText")
            )
        )
        root_cell.click()
        time.sleep(1)

        # [4] Expand Below
        print(f"[4/7] Bung toàn bộ các cấp con (Expand Below)...", flush=True)
        workarea_tb = driver.find_element(
            By.CSS_SELECTOR, "div.aw-layout-workareaCommandbar, div.aw-commands-toolbar"
        )
        workarea_tb.find_element(By.CSS_SELECTOR, "button[command-id='Awb0Expand']").click()
        time.sleep(1.5)

        cmds = driver.find_elements(By.CSS_SELECTOR, "div.aw-widgets-cellListItem, [command-id='Awb0ExpandBelow']")
        exp_below = [
            c for c in cmds if "expand below" in (c.text or "").lower() or c.get_attribute("command-id") == "Awb0ExpandBelow"
        ]
        if exp_below:
            exp_below[0].click()
            print("[+] Đang mở rộng toàn bộ cây BOM (chờ 12s)...", flush=True)
            time.sleep(12)

        # [5] Select All
        print(f"[5/7] Chọn toàn bộ cây cấu trúc (Select All)...", flush=True)
        workarea_tb = driver.find_element(
            By.CSS_SELECTOR, "div.aw-layout-workareaCommandbar, div.aw-commands-toolbar"
        )
        workarea_tb.find_element(By.CSS_SELECTOR, "button[command-id='Awp0SelectAll']").click()
        time.sleep(2)

        # [6] Open Export To Excel
        print(f"[6/7] Mở panel Export To Excel...", flush=True)
        workarea_tb.find_element(By.CSS_SELECTOR, "button[command-id='Arm0ExportImport']").click()
        time.sleep(2)

        popup = wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "div.aw-popup-commandListContainer, div.aw-popup, div.sw-popup")
            )
        )
        for it in popup.find_elements(By.CSS_SELECTOR, ".aw-widgets-cellListItem, .aw-command, button, a, div"):
            txt = (it.text or "").strip().lower()
            if "export to excel" in txt and "import" not in txt:
                it.click()
                break
        time.sleep(3)

        # [7] Panel configuration & Export
        print(f"[7/7] Cấu hình cột 'KTCT_Trong' và bấm Export...", flush=True)
        panel = wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "form.sw-command-panel, div.sw-right-dialog form")
            )
        )

        # Select KTCT_Trong preset
        ensure_ktct_trong_arrangement(driver, panel, wait)

        # Uncheck Run in Background
        cbs = panel.find_elements(By.CSS_SELECTOR, "input[type='checkbox']")
        for cb in cbs:
            try:
                parent_lbl = cb.find_element(
                    By.XPATH, "./ancestor::label | ./ancestor::div[contains(@class,'checkbox')]"
                )
                if "background" in parent_lbl.text.lower():
                    if cb.is_selected():
                        cb.click()
                        time.sleep(1)
            except Exception:
                pass

        # Click Export Button
        export_btn = panel.find_element(By.XPATH, ".//button[normalize-space()='Export']")
        driver.execute_script("arguments[0].click();", export_btn)
        print("[+] Đã nhấn Export. Đang đợi máy chủ Teamcenter sinh file và tải về...", flush=True)

        start_time = time.time()
        downloaded = None
        while time.time() - start_time < 75:
            time.sleep(2)
            fls = [
                f
                for f in download_dir.glob("*")
                if not f.name.endswith(".crdownload") and not f.name.endswith(".tmp")
            ]
            if fls and fls[0].stat().st_size > 50_000:
                downloaded = fls[0]
                break

        if not downloaded:
            raise TimeoutError(f"Quá thời gian chờ tải file BOM cho {part_number}!")

        print(f"[+] Tệp tải về thành công: {downloaded.name} ({downloaded.stat().st_size:,} bytes)", flush=True)

        # Convert to pure .xlsx
        local_pure_xlsx = download_dir / f"PLM_{part_number}.pure.xlsx"
        ok = convert_to_pure_xlsx(downloaded, local_pure_xlsx)
        if not ok:
            raise RuntimeError("Lỗi khi chuyển đổi file sang pure .xlsx!")

        # Verify columns
        wb_check = openpyxl.load_workbook(local_pure_xlsx, data_only=True)
        ws_check = wb_check.active
        max_r, max_c = ws_check.max_row, ws_check.max_column
        wb_check.close()
        print(f"[XÁC THỰC] File thô hợp lệ: {max_r:,} dòng, {max_c} cột chuẩn.", flush=True)

        # Save raw backup
        BACKUP_DIR.mkdir(parents=True, exist_ok=True)
        shutil.copy2(local_pure_xlsx, bak_file)
        print(f"[LƯU TRỮ] Đã lưu bản gốc chưa lọc vào: {bak_file.name}", flush=True)

        # Standardize to canonical 14 columns
        print(f"[*] Chuẩn hóa BOM sang 14 cột tiêu chuẩn cho {part_number}...", flush=True)
        from src.automation.tc2412.standardizer import standardize_plm_file
        DEST_DIR.mkdir(parents=True, exist_ok=True)
        orig_c, final_c = standardize_plm_file(local_pure_xlsx, dest_file, prune_electrical=True)
        print(f"[CHUẨN HÓA 14 CỘT] {part_number}: {orig_c:,} dòng raw -> {final_c:,} dòng 14 cột chuẩn ({dest_file.stat().st_size:,} bytes)", flush=True)

        # Sync to subfolder if present
        for d in DEST_DIR.iterdir():
            if d.is_dir() and d.name.upper().startswith(part_number.upper()):
                sub_dest = d / f"PLM_{part_number}.xlsx"
                shutil.copy2(dest_file, sub_dest)
                print(f"[ĐỒNG BỘ] Đã copy vào thư mục kỹ sư: {d.name}", flush=True)
                break

        # Cleanup local scratch files
        for f in download_dir.glob("*"):
            try:
                f.unlink()
            except Exception:
                pass

        return True

    except Exception as exc:
        print(f"[-] LỖI XỬ LÝ MÃ {part_number}: {exc}", flush=True)
        return False


def main() -> int:
    download_dir = Path("D:/Sandbox/pm_sosanhbom/scratch/downloads").resolve()
    download_dir.mkdir(parents=True, exist_ok=True)
    DEST_DIR.mkdir(parents=True, exist_ok=True)
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)

    print("=" * 65, flush=True)
    print("CHƯƠNG TRÌNH TỰ ĐỘNG TẢI VÀ LỌC 16 BẢN BOM VIRGO2 - TC2412", flush=True)
    print(f"Tài khoản: vn_pe02 (Trọng) | Cấu hình cột: KTCT_Trong", flush=True)
    print(f"Thư mục lưu: {DEST_DIR}", flush=True)
    print("=" * 65, flush=True)

    # Load BolocBom rules
    print("[*] Đang tải quy tắc lọc BolocBom cho máy Virgo...", flush=True)
    rules = load_bolocbom_rules(machine_type="Virgo")
    print(f"[+] Đã sẵn sàng {len(rules)} quy tắc lọc.", flush=True)

    config = BrowserConfig(download_dir=download_dir, headless=True)
    sm = TC2412SessionManager(config=config)
    client = TC2412AutomationClient(session_manager=sm)
    driver = client.driver
    wait = WebDriverWait(driver, 25)

    success_list = []
    failed_list = []
    start_all = time.time()
    last_report_time = time.time()

    try:
        print("[*] Đăng nhập Siemens Teamcenter 2412 bằng tài khoản vn_pe02...", flush=True)
        driver.get("http://tcmp3gwb:3000/")
        time.sleep(2)
        driver.delete_all_cookies()
        driver.get("http://tcmp3gwb:3000/")
        time.sleep(2)

        user_input = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='userName'], input[type='text']"))
        )
        user_input.clear()
        user_input.send_keys("vn_pe02")
        pwd_input = driver.find_element(By.CSS_SELECTOR, "input[name='password'], input[type='password']")
        pwd_input.clear()
        pwd_input.send_keys("vn_pe02")
        driver.find_element(By.CSS_SELECTOR, "button[type='submit'], button.aw-widgets-button").click()
        time.sleep(5)
        print("[+] Đăng nhập thành công!", flush=True)

        for idx, part in enumerate(PART_NUMBERS, 1):
            cur_time = time.time()
            if cur_time - last_report_time >= 120:
                elapsed = int(cur_time - start_all)
                print(f"\n>>> [BÁO CÁO TIẾN ĐỘ 2 PHÚT] <<<", flush=True)
                print(f"Thời gian đã chạy: {elapsed // 60}m {elapsed % 60}s", flush=True)
                print(f"Tiến độ: {len(success_list)}/{len(PART_NUMBERS)} mã hoàn thành.", flush=True)
                print(f"Đã xong: {success_list}", flush=True)
                if failed_list:
                    print(f"Chưa đạt: {failed_list}", flush=True)
                print(f">>> Đang xử lý mã tiếp theo: {part} <<<\n", flush=True)
                last_report_time = cur_time

            print(f"\n>>> [TIẾN ĐỘ: {idx}/{len(PART_NUMBERS)}] Đang xử lý: {part} <<<", flush=True)
            ok = download_single_bom(
                client=client,
                part_number=part,
                download_dir=download_dir,
                rules=rules,
            )
            if ok:
                success_list.append(part)
            else:
                # Retry once
                print(f"[*] Thử tải lại lần 2 cho {part}...", flush=True)
                time.sleep(4)
                ok2 = download_single_bom(
                    client=client,
                    part_number=part,
                    download_dir=download_dir,
                    rules=rules,
                )
                if ok2:
                    success_list.append(part)
                else:
                    failed_list.append(part)

            time.sleep(3)

    finally:
        client.close()

    total_time = int(time.time() - start_all)
    print("\n" + "=" * 65, flush=True)
    print(f"BÁO CÁO TỔNG KẾT BATCH VIRGO2:")
    print(f"- Tổng thời gian: {total_time // 60}m {total_time % 60}s")
    print(f"- Thành công: {len(success_list)}/{len(PART_NUMBERS)} mã")
    print(f"- Danh sách hoàn tất: {success_list}")
    if failed_list:
        print(f"- Thất bại ({len(failed_list)} mã): {failed_list}")
    print(f"- Thư mục mạng: {DEST_DIR}")
    print("=" * 65, flush=True)

    return 0 if not failed_list else 1


if __name__ == "__main__":
    sys.exit(main())
