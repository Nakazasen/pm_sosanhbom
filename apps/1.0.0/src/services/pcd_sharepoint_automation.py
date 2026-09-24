"""PCD Production Plan SharePoint Online Browser Automation Service.

Automates Microsoft Edge in headless mode to navigate to the official
Kyocera PCD Production Plan SharePoint Online library, leverage corporate SSO,
locate the monthly plan file (*定期計画後明細計画.xlsx), and download it directly
into the application workspace without requiring manual OneDrive sync configuration.
"""

from __future__ import annotations

import logging
import os
import re
import sys
import time
import urllib.parse
from pathlib import Path
from typing import Callable, Optional

logger = logging.getLogger(__name__)


class PCDSharePointDownloadError(Exception):
    """Raised when SharePoint browser automation fails to locate or download the file."""
    pass


def build_pcd_sharepoint_folder_url(year: int, month: int) -> str:
    """Build exact SharePoint AllItems.aspx folder URL for given year and month."""
    month_folder = f"{year}{month:02d}"
    folder_path = f"/sites/kdtvn_PCD/ProductionPlan/Theo tháng (月別)/{year}/{month_folder}"
    encoded_id = urllib.parse.quote(folder_path, safe="")
    return (
        f"https://kdcf.sharepoint.com/sites/kdtvn_PCD/ProductionPlan/Forms/AllItems.aspx"
        f"?id={encoded_id}"
    )


def download_pcd_plan_from_sharepoint(
    year: int,
    month: int,
    download_dir: Optional[Path] = None,
    progress_callback: Optional[Callable[[str], None]] = None,
    headless: bool = True,
    timeout_seconds: int = 35,
) -> Path:
    """Automate Microsoft Edge to download the monthly plan from SharePoint Online.

    Args:
        year: Target plan year (e.g. 2026).
        month: Target plan month (1-12).
        download_dir: Target destination directory for the downloaded file.
                      Defaults to data/pcd_plans/{year}{month:02d}.
        progress_callback: Optional callback for status messages.
        headless: Whether to run Edge without visible GUI window.
        timeout_seconds: Maximum time to wait for page load and download.

    Returns:
        Path to the downloaded and verified .xlsx plan file.

    Raises:
        PCDSharePointDownloadError: If download fails or file is not found.
    """
    def log(msg: str) -> None:
        logger.info(msg)
        if progress_callback:
            try:
                progress_callback(msg)
            except Exception:
                pass

    if download_dir is None:
        download_dir = Path("data/pcd_plans") / f"{year}{month:02d}"
    download_dir = Path(download_dir).resolve()
    download_dir.mkdir(parents=True, exist_ok=True)

    log(f"[*] Khởi động trình duyệt Edge để kết nối SharePoint PCD ({year}/{month:02d})...")

    try:
        from selenium import webdriver
        from selenium.webdriver.edge.options import Options as EdgeOptions
        from selenium.webdriver.common.by import By
    except ImportError as e:
        raise PCDSharePointDownloadError(
            "Thư viện selenium chưa được cài đặt. Vui lòng cài đặt: pip install selenium"
        ) from e

    options = EdgeOptions()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--ignore-certificate-errors")

    prefs = {
        "download.default_directory": str(download_dir),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }
    options.add_experimental_option("prefs", prefs)

    # Snapshot existing files in download_dir to detect new file accurately
    existing_files = {p.resolve(): p.stat().st_mtime for p in download_dir.glob("*") if p.is_file()}

    driver = None
    try:
        driver = webdriver.Edge(options=options)

        # Enable CDP Page.setDownloadBehavior to intercept downloads in headless Chromium
        try:
            driver.execute_cdp_cmd(
                "Page.setDownloadBehavior",
                {
                    "behavior": "allow",
                    "downloadPath": str(download_dir),
                    "eventsEnabled": True,
                },
            )
        except Exception as cdp_err:
            logger.debug("Could not set CDP download behavior: %s", cdp_err)

        target_url = build_pcd_sharepoint_folder_url(year, month)
        log(f"[*] Đang truy cập thư viện PCD: Tháng {month:02d}/{year}...")
        driver.get(target_url)

        # Poll for list elements to render
        start_time = time.time()
        rows = []
        log("[*] Đang tải danh sách tài liệu từ SharePoint...")
        while time.time() - start_time < 20:
            time.sleep(1.5)
            rows = driver.find_elements(
                By.CSS_SELECTOR,
                "[role='row'], [data-automationid='DetailsRow']"
            )
            if len(rows) > 1:  # Header + at least one data row
                break

        if not rows:
            raise PCDSharePointDownloadError(
                f"Không tìm thấy thư mục hoặc tệp nào trên SharePoint PCD cho Tháng {month:02d}/{year}."
            )

        # Identify candidate files
        # Priority order: 定期計画後明細計画 > 月初明細計画 > 明細計画
        target_row = None
        target_filename = ""

        # First pass: check for end_of_period plan (*定期計画後明細計画*)
        for r in rows:
            txt = r.text
            if "定期計画後明細計画" in txt and not any(kw in txt for kw in ["品質状況", "kaizo", "chat luong"]):
                target_row = r
                target_filename = f"{month}月度定期計画後明細計画.xlsx"
                break

        # Second pass: check for beginning_of_month plan (*月初明細計画*)
        if not target_row:
            for r in rows:
                txt = r.text
                if "月初明細計画" in txt and not any(kw in txt for kw in ["品質状況", "kaizo", "chat luong"]):
                    target_row = r
                    target_filename = f"{month}月度月初明細計画.xlsx"
                    break

        # Third pass: any plan file
        if not target_row:
            for r in rows:
                txt = r.text
                if ("明細計画" in txt or "計画" in txt) and ".xlsx" in txt:
                    if not any(kw in txt for kw in ["品質状況", "kaizo", "chat luong"]):
                        target_row = r
                        target_filename = f"{month}月度計画.xlsx"
                        break

        if not target_row:
            raise PCDSharePointDownloadError(
                f"Không tìm thấy tệp kế hoạch sản xuất (*明細計画.xlsx) trong thư mục Tháng {month:02d}/{year} trên SharePoint PCD."
            )

        log(f"[+] Đã tìm thấy tệp kế hoạch: {target_filename}. Đang kích hoạt tải xuống...")

        # Select target row using JavaScript click
        driver.execute_script("arguments[0].click();", target_row)
        time.sleep(1.5)

        # Locate Download button in CommandBar
        download_btn = None
        cmd_buttons = driver.find_elements(
            By.CSS_SELECTOR,
            "[data-automationid='CommandBar'] button, [role='menubar'] button, button[data-automationid='downloadCommand']"
        )
        for cb in cmd_buttons:
            btn_txt = (cb.text or "").strip().lower()
            aria = (cb.get_attribute("aria-label") or "").lower()
            name = (cb.get_attribute("name") or "").lower()
            data_id = (cb.get_attribute("data-automationid") or "").lower()
            if any(term in s for term in ["download", "tải xuống"] for s in [btn_txt, aria, name, data_id]):
                download_btn = cb
                break

        if not download_btn:
            # Fallback: look for row more actions button
            more_btns = target_row.find_elements(
                By.CSS_SELECTOR,
                "button[data-automationid='moreActionsHeroField'], button[aria-label*='Action']"
            )
            if more_btns:
                driver.execute_script("arguments[0].click();", more_btns[0])
                time.sleep(1)
                context_download = driver.find_elements(By.CSS_SELECTOR, "button[name*='Download'], button[aria-label*='Download'], button[name*='Tải xuống']")
                if context_download:
                    download_btn = context_download[0]

        if not download_btn:
            raise PCDSharePointDownloadError(
                "Không tìm thấy nút 'Tải xuống' (Download) trên thanh công cụ SharePoint."
            )

        # Click Download button
        driver.execute_script("arguments[0].click();", download_btn)
        log("[*] Đang nhận tệp từ máy chủ SharePoint...")

        # Wait for file download to complete
        downloaded_file: Optional[Path] = None
        wait_start = time.time()
        while time.time() - wait_start < timeout_seconds:
            time.sleep(1)

            # Check if any .crdownload or .tmp is active
            crdownloads = list(download_dir.glob("*.crdownload")) + list(download_dir.glob("*.tmp"))
            if crdownloads:
                continue

            # Look for new or modified file
            current_files = [p for p in download_dir.glob("*") if p.is_file()]
            for p in current_files:
                p_res = p.resolve()
                if p_res not in existing_files or p.stat().st_mtime > existing_files[p_res]:
                    # Ensure file has non-trivial size
                    if p.stat().st_size > 1024:
                        downloaded_file = p
                        break
            if downloaded_file:
                break

        if not downloaded_file:
            raise PCDSharePointDownloadError(
                "Quá thời gian chờ tải tệp từ SharePoint (Timeout). Vui lòng thử lại hoặc kiểm tra kết nối mạng."
            )

        # If file was downloaded with GUID name without .xlsx extension, rename to target_filename
        if downloaded_file.suffix.lower() not in [".xlsx", ".xlsm", ".xls"]:
            final_path = download_dir / target_filename
            try:
                if final_path.exists():
                    final_path.unlink()
                downloaded_file.replace(final_path)
                downloaded_file = final_path
            except Exception as ren_err:
                logger.warning("Could not rename downloaded file: %s", ren_err)

        log(f"[+] Tải thành công tệp: {downloaded_file.name} ({downloaded_file.stat().st_size:,} bytes)!")
        return downloaded_file

    except PCDSharePointDownloadError:
        raise
    except Exception as exc:
        logger.exception("Unexpected error during SharePoint download automation: %s", exc)
        raise PCDSharePointDownloadError(f"Lỗi kết nối trình duyệt: {exc}") from exc
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass
