# SPEC-KIT: TỰ ĐỘNG HÓA TẢI BOM PLM TỪ SIEMENS TEAMCENTER VERSION 2412 (ACTIVE WORKSPACE)

> **Tài liệu nguồn**: `\\fstvn01\Data\10_Production Engineering Department(製造技術部)\02.製造技術課\PE Dept\15. FORM（BIEU MAU）-形式\Form_VBA\Form_VBA_SS_BOM\tai_lieu_huong_dan_download_BOM.pptx`  
> **Hệ thống đích**: **Siemens Teamcenter Version 2412 (TC2412 / Active Workspace)**  
> *(Xác thực DOM thực tế: `<p class="aw-login-copyrightTitle">Version 2412</p>` & `Copyright © 2024 Siemens`)*  
> **Cổng truy cập**: `http://tcmp3gwb:3000/`  
> **Phiên bản đặc tả**: 1.0.0 (Authoritative Release)  
> **Trạng thái**: APPROVED & AUTHORITATIVE  
> **Sở hữu tài liệu**: Milestone M1 (Spec-Kit Specification)  

---

## MỤC LỤC
1. [TỔNG QUAN KIẾN TRÚC & QUY TRÌNH HỆ THỐNG](#1-tổng-quan-kiến-trúc--quy-trình-hệ-thống)
   - 1.1. [Kiến trúc Adapter Pattern](#11-kiến-trúc-adapter-pattern-tuân-thủ-r5-original_requestmd)
   - 1.2. [Chuỗi 7 Phân Hệ Tự Động Hóa + Trình Báo Cáo R8](#12-chuỗi-7-phân-hệ-tự-động-hóa--trình-báo-cáo-r8)
   - 1.3. [Hợp Đồng Điều Phối Toàn Cục (Top-Level Orchestration Contract)](#13-hợp-đồng-điều-phối-toàn-cục-top-level-orchestration-contract)
2. [CÁC QUYẾT ĐỊNH ĐÃ CHỐT TOÀN DIỆN (CLOSED RESOLUTIONS)](#2-các-quyết-định-đã-chốt-toàn-diện-closed-resolutions)
3. [DANH SÁCH 14 CỘT CHUẨN HÓA (CANONICAL 14-COLUMN CONTRACT)](#3-danh-sách-14-cột-chuẩn-hóa-canonical-14-column-contract)
4. [PHASE 1: AUTH-01 – XÁC THỰC & LƯU TRỮ TÀI KHOẢN DPAPI](#4-phase-1-auth-01--xác-thực--lưu-trữ-tài-khoản-dpapi)
5. [PHASE 2: BOM-SEARCH-01 – TÌM KIẾM ITEM & ĐIỀU HƯỚNG TAB CONTENT](#5-phase-2-bom-search-01--tìm-kiếm-item--điều-hướng-tab-content)
6. [PHASE 3: BOM-EXPAND-01 – BUNG CÂY CẤU TRÚC ĐẾN LEVEL 7](#6-phase-3-bom-expand-01--bung-cây-cấu-trúc-đến-level-7)
7. [PHASE 4: BOM-SELECT-01 – CHỌN TOÀN BỘ CÂY BOM & MỞ MENU EXCEL REPORT](#7-phase-4-bom-select-01--chọn-toàn-bộ-cây-bom--mở-menu-excel-report)
8. [PHASE 5: BOM-EXPORT-OPEN-01 – KÍCH HOẠT EXPORT TO EXCEL & CHỐT CHẶN IMPORT](#8-phase-5-bom-export-open-01--kích-hoạt-export-to-excel--chốt-chặn-import)
9. [PHASE 6: BOM-EXPORT-CONFIG-01 – CẤU HÌNH THUỘC TÍNH ITEM & 14 CỘT HIỂN THỊ](#9-phase-6-bom-export-config-01--cấu-hình-thuộc-tính-item--14-cột-hiển-thị)
10. [PHASE 7: BOM-EXPORT-RUN-01 – XUẤT DỮ LIỆU, GIÁM SÁT TẢI VỀ & KIỂM TRA OPENPYXL](#10-phase-7-bom-export-run-01--xuất-dữ-liệu-giám-sát-tải-về--kiểm-tra-openpyxl)
11. [PHASE 8: REP-01 (R8) – TRÌNH BÁO CÁO TIẾN ĐỘ ĐỊNH KỲ (2 PHÚT/CHU KỲ)](#11-phase-8-rep-01-r8--trình-báo-cáo-tiến-độ-định-kỳ-2-phútchu-kỳ)
12. [MA TRẬN TRUY XUẤT NGUỒN GỐC & KIỂM THỬ TOÀN DIỆN (TRACEABILITY MATRIX TIER 1 - TIER 5)](#12-ma-trận-truy-xuất-nguồn-gốc--kiểm-thử-toàn-diện-traceability-matrix-tier-1---tier-5)
13. [CHUẨN HÓA DỮ LIỆU CANONICAL & CẦU NỐI SHEET PLM BẢO TOÀN CÔNG THỨC EXCEL](#13-chuẩn-hóa-dữ-liệu-canonical--cầu-nối-sheet-plm-bảo-toàn-công-thức-excel)

---

## 1. TỔNG QUAN KIẾN TRÚC & QUY TRÌNH HỆ THỐNG

### 1.1. Kiến trúc Adapter Pattern (Tuân thủ R5 ORIGINAL_REQUEST.md)
Hệ thống tự động hóa tải dữ liệu BOM PLM từ Teamcenter Version 2412 được thiết kế theo mẫu kiến trúc Adapter Pattern, tách biệt tuyệt đối tầng thu thập dữ liệu tự động (Web Automation Layer) khỏi Lõi thuật toán so sánh BOM (Core Comparison & Tree Engine):

```
+-----------------------------------------------------------------------------------+
| Giao diện người dùng Desktop (CustomTkinter / i18n VN-JP-CN) / CLI Scripts        |
+-----------------------------------------------------------------------------------+
                                          │
                                          ▼
+-----------------------------------------------------------------------------------+
| Lõi So Sánh BOM & Phân Cấp Cây (src/core/)                                        |
| - engine.py (Bóc tách Model Virgo/Libra/Iris, duyệt cây 6 tầng, đối soát CTTT)   |
| - model_pruner.py (Cắt tỉa cụm không cần bung, prune_excel_file)                  |
| - adapters.py (PLMProvider Interface)                                             |
+-----------------------------------------------------------------------------------+
                                          ▲
                                          │ Truyền file Excel PLM 14 cột chuẩn
+-----------------------------------------------------------------------------------+
| Tầng Cung Cấp Dữ Liệu PLM (PLMProvider / TeamcenterSeleniumAdapter)               |
+-----------------------------------------------------------------------------------+
        │                                                     │
        ▼                                                     ▼
+───────────────────────────+                 +─────────────────────────────────────+
| Security & Storage        |                 | TC2412 Automation Client            |
| (src/security/)           |                 | (src/automation/tc2412/)            |
| - dpapi.py: Windows DPAPI |                 | - session.py: Headless Edge/Chrome  |
| - credentials.py: Keyring |                 | - client.py: 7-Phase Execution      |
| - Secure CLI Update/Delete|                 | - selectors.py: AW2412 DOM Locators |
|                           |                 | - progress_reporter.py: Chu kỳ 2 min|
+───────────────────────────+                 +─────────────────────────────────────+
                                                              │
                                                              ▼ (Lớp tương thích ngược)
                                              +─────────────────────────────────────+
                                              | src/automation/tc14/                |
                                              | Aliasing TC14* -> TC2412*           |
                                              +─────────────────────────────────────+
```

### 1.2. Chuỗi 7 Phân Hệ Tự Động Hóa + Trình Báo Cáo R8
Quy trình thực thi toàn trình (End-to-End Execution Pipeline) diễn ra tuần tự qua 7 giai đoạn nghiêm ngặt:
1. **AUTH-01**: Xác thực ngầm (Silent Auto-login) với Siemens Teamcenter Version 2412, bảo mật thông tin bằng Windows DPAPI.
2. **BOM-SEARCH-01**: Tìm kiếm Item ID qua Global Search Box, đối soát Item ID & Revision, điều hướng chính xác vào tab `Content`.
3. **BOM-EXPAND-01**: Chọn nút gốc (Root Node `data-indexnumber="0"`), kích hoạt menu `Expand Below` với độ sâu Level 7, áp dụng cơ chế chờ linh hoạt (Dynamic Explicit Waits) ở chế độ Headless.
4. **BOM-SELECT-01**: Chọn toàn bộ các dòng BOM (`Awp0SelectAll`), kiểm tra trạng thái chỉ số selection count và header checkbox, mở menu Excel Report (`Arm0ExportImport`).
5. **BOM-EXPORT-OPEN-01**: Kích hoạt tác vụ `Export to Excel`, thiết lập chốt chặn an toàn Fail-Closed tuyệt đối không kích hoạt `Import Changes`.
6. **BOM-EXPORT-CONFIG-01**: Chuyển nguồn thuộc tính sang `Item`, lọc tìm kiếm thuộc tính qua ô search box, xóa các cột mặc định sai, thiết lập chính xác 14 cột Displayed Columns theo thứ tự chuẩn hóa.
7. **BOM-EXPORT-RUN-01**: Bật `Run in Background`, bấm `Export`, bắt thông báo hoàn tất, click trigger tải tệp an toàn về thư mục đích, kiểm tra tính toàn vẹn bằng thư viện `openpyxl` (14 cột, dòng > 0, Unicode tiếng Nhật font `MS Gothic`).
8. **REP-01 (R8)**: Luồng chạy nền phát thông báo tiến độ định kỳ mỗi 2 phút (120 giây) theo định dạng: `Phase [X/8] - [Name] - [%] - [Elapsed]`.

### 1.3. Hợp Đồng Điều Phối Toàn Cục (Top-Level Orchestration Contract)

Tuân thủ nghiêm ngặt hợp đồng kiến trúc tại `PROJECT.md:76-84`, toàn bộ quy trình 7 phân hệ tải BOM được bao bọc và điều phối tập trung thông qua lớp điều khiển `TC2412AutomationClient`. Đây là giao diện lập trình duy nhất (Authoritative Entry Point) mà tầng Core Adapter (`src/core/adapters.py`) gọi tới.

#### Hợp đồng Lớp Python (Authoritative Class Contract)
```python
from pathlib import Path
from typing import Callable, Optional

class TC2412AutomationError(Exception):
    """Lớp cơ sở cho toàn bộ ngoại lệ phát sinh trong quá trình tự động hóa TC2412."""
    pass

class TC2412AutomationClient:
    """Điều phối viên toàn trình cho quy trình tải BOM tự động từ Siemens Teamcenter 2412."""

    def __init__(
        self,
        base_url: str = "http://tcmp3gwb:3000/",
        headless: bool = True,
        browser: str = "edge",  # 'edge' hoặc 'chrome'
        timeout: int = 180,
    ):
        """Khởi tạo client với cấu hình phiên làm việc."""
        self.base_url = base_url
        self.headless = headless
        self.browser = browser
        self.timeout = timeout

    def download_bom_full(
        self,
        part_number: str,
        output_dir: Path,
        part_rev: str | None = None,
        progress_callback: Callable[[int, int, str, float, float], None] | None = None,
    ) -> Path:
        """Thực thi tuần tự toàn bộ 7 phân hệ tải BOM từ TC2412 sang Excel.
        
        Args:
            part_number: Mã linh kiện / Item ID (ví dụ: '110C103NL0')
            output_dir: Thư mục đích lưu tệp Excel tải về
            part_rev: Ký hiệu Revision cụ thể (nếu None sẽ tự động chọn Latest Released Revision)
            progress_callback: Callback nhận tiến độ theo 5 tham số chuẩn (PROJECT.md:82):
                (phase_index: int, total_phases: int, phase_name: str, percent: float, elapsed_seconds: float) -> None
            
        Returns:
            Path: Đường dẫn tuyệt đối đến tệp Excel BOM 14 cột đã được kiểm tra tính toàn vẹn (openpyxl)
            
        Raises:
            AuthenticationError: Khi đăng nhập TC2412 thất bại
            ItemNotFoundError: Khi mã linh kiện không tồn tại trên hệ thống
            BOMExpandTimeoutError: Khi bung cây Level 7 vượt quá thời gian chờ (180s)
            NoRowsSelectedError: Khi không có dòng BOM nào được chọn trên bảng
            ImportChangesForbiddenError: Khi phát hiện nguy cơ kích hoạt lệnh Import Changes
            ColumnConfigurationError: Khi cấu hình cột không khớp đúng 14 cột chuẩn
            ExportDownloadTimeoutError: Khi máy chủ hoặc mạng quá hạn tải tệp (300s)
            BOMValidationError: Khi cấu trúc tệp Excel không khớp 14 header chuẩn hoặc thiếu dòng dữ liệu
            BOMCorruptFileError: Khi tệp tải về rỗng (< 5KB) hoặc hỏng định dạng
        """
        ...
```

#### Quy Trình Điều Phối Tuần Tự (Sequencing Lifecycle)
1. **Khởi Tạo Session & ProgressReporter**: Khởi động trình duyệt Headless WebDriver (Edge/Chrome) qua `src/automation/tc2412/session.py`. Đăng ký `progress_callback` vào `ProgressReporter` (bắt đầu phát heartbeat mỗi 120 giây).
2. **Thực Thi Tuần Tự 7 Phase**:
   - **Phase 1 [1/8]**: `AUTH-01` – Gọi `CredentialManager.get_credentials()`, điền form, xác thực session cookie.
   - **Phase 2 [2/8]**: `BOM-SEARCH-01` – Tìm kiếm `part_number`, đối soát `part_rev` (hoặc Latest Released), chuyển vào tab `Content`.
   - **Phase 3 [3/8]**: `BOM-EXPAND-01` – Chọn root node, gọi Expand Below đến Level 7, áp dụng dynamic wait theo dõi `div.aw-splm-tableSummary` và spinner.
   - **Phase 4 [4/8]**: `BOM-SELECT-01` – Thực thi Select All, xác thực chỉ số `div.aw-splm-tableSelectionCount` và header checkbox, mở menu Excel Report (`Arm0ExportImport`).
   - **Phase 5 [5/8]**: `BOM-EXPORT-OPEN-01` – Quét menu với chốt chặn Fail-Closed Anti-Import, mở panel `Export To Excel`.
   - **Phase 6 [6/8]**: `BOM-EXPORT-CONFIG-01` – Chọn nguồn `Item`, dùng ô filter `input.aw-uiwidgets-searchBox` chọn đủ 14 cột, xếp chuẩn thứ tự 1-14.
   - **Phase 7 [7/8]**: `BOM-EXPORT-RUN-01` – Tích `Run in Background`, bấm `Export`, bắt thông báo hoàn tất (`div.noty_message` / `a.aw-widgets-cellListCellTitle`), kích hoạt click tải tệp, giám sát download folder, đối soát `openpyxl`.
   - **Phase 8 [8/8]**: `COMPLETE` – Hoàn tất toàn trình (100%), trả về đối tượng `Path` tệp Excel.
3. **Quản Lý Vòng Đời & Thu Hồi Tài Nguyên (Resource Cleanup Lifecycle)**:
   - Toàn bộ khối thực thi được bọc trong cấu trúc bảo vệ `try ... finally: driver.quit()` và `reporter.stop()`.
   - Dù tiến trình thành công hay gặp bất kỳ ngoại lệ nào, WebDriver luôn được đóng sạch sẽ, tuyệt đối không để lại tiến trình trình duyệt mồ côi (`msedgedriver.exe`, `chromedriver.exe`) trên Windows.
4. **Phân Cấp Ngoại Lệ (Exception Hierarchy)**:
   - Tất cả ngoại lệ thuộc các phân hệ đều kế thừa từ lớp cơ sở `TC2412AutomationError(Exception)`, giúp tầng gọi `TeamcenterSeleniumAdapter` bắt lỗi thống nhất và an toàn.

---

## 2. CÁC QUYẾT ĐỊNH ĐÃ CHỐT TOÀN DIỆN (CLOSED RESOLUTIONS)

Căn cứ theo chỉ thị tại `ORIGINAL_REQUEST.md` (phiên `04:32:25Z`, `04:33:53Z`, `04:35:40Z`), toàn bộ các câu hỏi mở (`open_questions`) trong bản phác thảo cũ được chốt chính thức như sau:

| Mã Phase | Câu hỏi mở ban đầu (Old Draft) | Quyết định kỹ thuật chốt chính thức (Final Decision) | Căn cứ thẩm quyền |
|:---|:---|:---|:---|
| **AUTH-01** | Tự điền thông tin hay tự động đăng nhập ngầm? Lưu ở đâu? Đăng xuất có xóa thông tin không? | **Tự động đăng nhập ngầm hoàn toàn (Silent Auto-login)** khi đã lưu credentials. Lưu trữ mã hóa cấp OS qua **Windows DPAPI** (Service `PM_SOSANHBOM_TC2412`). Thông tin được giữ lâu dài; chỉ xóa khi người dùng gọi lệnh `--delete-credentials`. | Chỉ thị 04:32:25Z, 04:35:40Z |
| **BOM-SEARCH-01** | Tìm kiếm chính xác hay Wildcard? Quy tắc chọn Revision khi có nhiều kết quả? | Nhập **chính xác Item ID**. Khi có nhiều Revision, ưu tiên chọn **Revision được Release mới nhất** (`Release Status` hợp lệ) hoặc Revision do người dùng truyền vào qua tham số `part_rev`. | Slide 2, Chỉ thị 04:35:40Z |
| **BOM-EXPAND-01** | Bung cây đến mức sâu bao nhiêu? Cơ chế chờ nạp cây lớn? | Bung đến độ sâu tối đa **Level 7** (`Expand Below` = 7). Sử dụng **Dynamic Explicit Wait** theo dõi spinner loading `.aw-layout-progressBar` (`aria-busy="false"` + `hide`) và số lượng dòng `div.aw-splm-tableRow` ổn định trong 3 chu kỳ liên tiếp (timeout 180s). | Slide 3, Chỉ thị 04:35:40Z |
| **BOM-SELECT-01** | Phạm vi Select All là gì? Có gồm dòng gốc không? | Phạm vi là **toàn bộ 100% các dòng BOM đã bung** trên bảng `occTreeTable` (bao gồm cả dòng gốc Level 0/1 và toàn bộ các dòng con đến Level 7). Tất cả phải có `aria-selected="true"`. | Slide 4, Chỉ thị 04:35:40Z |
| **BOM-EXPORT-OPEN-01** | Cách phòng vệ chống click nhầm vào "Import Changes"? | Cài đặt **Fail-Closed Anti-Import Guard**: Quét thuộc tính và văn bản của phần tử menu, nếu chứa chuỗi `"import"` thì **tuyệt đối không click**. Chỉ click khi khớp chính xác `"export"` hoặc command-id `Awp0ExportToExcel`. Nếu phát hiện click nhầm, ngắt tiến trình lập tức. | Slide 5, Chỉ thị 04:35:40Z |
| **BOM-EXPORT-CONFIG-01** | Nguồn thuộc tính và danh sách cột bắt buộc là gì? | Nguồn thuộc tính bắt buộc là **`Item`**. Danh sách cột hiển thị bắt buộc gồm **đúng 14 cột** theo thứ tự quy định tại Mục 3. | Slide 6, Slide 16/20/27, Chỉ thị 04:35:40Z |
| **BOM-EXPORT-RUN-01** | Có chạy nền không? Quy tắc đặt tên file và kiểm tra tính toàn vẹn? | Luôn bật cờ **`Run in Background`** cho các BOM phức tạp. Tệp tải về được đặt tên dạng `PLM_<ItemID>.xlsx`. Kiểm tra toàn vẹn bằng `openpyxl`: file > 5KB, đúng 14 cột, số dòng > 1, bảo toàn Unicode tiếng Nhật (`MS Gothic`). | Slide 7, Chỉ thị 04:35:40Z |
| **Phiên bản đích** | Teamcenter 14 hay Teamcenter 2412? | **Siemens Teamcenter Version 2412 (TC2412)**. Mọi DOM locators, CSS classes và xử lý giao diện đều chuẩn hóa theo Active Workspace 2412. | Chỉ thị 04:33:53Z |

---

## 3. DANH SÁCH 14 CỘT CHUẨN HÓA (CANONICAL 14-COLUMN CONTRACT)

Thứ tự 14 cột trong tệp Excel xuất từ Teamcenter 2412 là hợp đồng dữ liệu bất biến (Immutable Data Contract), làm đầu vào trực tiếp cho Lõi phân tích cây `pm_sosanhbom` (bóc tách cụm Unit, kiểm tra ngày hiệu lực, lọc Model máy Virgo/Libra/Iris):

| Cột # | Vị trí Excel | Tên cột hiển thị chuẩn (Displayed Column) | Nguồn thuộc tính (Property Source) | Thuộc tính kỹ thuật trong TC2412 (Property Name) | Kiểu dữ liệu | Ý nghĩa nghiệp vụ trong thuật toán Core |
|:---:|:---:|:---|:---:|:---|:---:|:---|
| **1** | **A** | **Home** | Item | `object_name` / `item_id` | String | Định danh thư mục gốc / Vị trí phân cấp cao nhất |
| **2** | **B** | **Level** | BOM Line | `bl_level_starting_0` / `bl_indented_title` | Integer | Tầng phân cấp BOM (0, 1, 2, ..., 6, 7). Lõi dùng duyệt 6 tầng |
| **3** | **C** | **Item Type** | Item | `object_type` | String | Loại đối tượng chi tiết (Design Part, Commercial Part...) |
| **4** | **D** | **Item Id** | Item | `item_id` | String | Mã linh kiện (Part Number) dùng tra cứu Unit và so khớp SAP CS12 |
| **5** | **E** | **Has Children** | BOM Line | `bl_has_children` | Boolean | `true`/`false`. Quyết định cụm chi tiết có linh kiện con hay không |
| **6** | **F** | **Quantity** | BOM Line | `bl_quantity` | Float / Int | Số lượng cấu thành cụm (Quantity per assembly) |
| **7** | **G** | **1st Parts** | Item Revision | `item_revision_id` / `first_parts` | String | Phân loại linh kiện chính / linh kiện sơ cấp |
| **8** | **H** | **2nd BOM Flag** | BOM Line | `bom_flag_2` | String | Cờ quản lý phân rã BOM cấp 2 |
| **9** | **I** | **Occurrence Effectivities** | BOM Line | `bl_occ_effectivity` | String | Chuỗi ngày/số hiệu lực (ví dụ: `to 2026/12/31`, `UP`) để lọc hiệu lực |
| **10** | **J** | **Item Revision Project List** | Item Revision | `project_list` | String | Danh sách Model máy áp dụng (Virgo, Libra2, Iris2024...) |
| **11** | **K** | **Item Name** | Item | `object_name` | String | Tên tiếng Nhật / Anh của chi tiết (`MS Gothic`, Kanji/Katakana) |
| **12** | **L** | **Notice No** | Item Revision | `ec_notice_no` | String | Số thông báo kỹ thuật / ECN phát hành chi tiết |
| **13** | **M** | **Revision** | Item Revision | `item_revision_id` | String | Ký hiệu phiên bản chi tiết (ví dụ: `A`, `01`, `04`) |
| **14** | **N** | **Item Rev Status** | Item Revision | `release_status_list` | String | Trạng thái phát hành (`Approved`, `Released`, `Working`) |

---

## 4. PHASE 1: AUTH-01 – XÁC THỰC & LƯU TRỮ TÀI KHOẢN DPAPI

### 4.1. User Story
```
As a Production Engineering Technician or Automation System,
I want to authenticate automatically into Siemens Teamcenter Version 2412 using securely stored OS credentials (Windows DPAPI),
So that the automation pipeline executes unattended without exposing plaintext passwords on disk and allows updating or revoking credentials when needed.
```

### 4.2. Acceptance Criteria (Given-When-Then)

- **AC-AUTH-01 (Silent Auto-Login Thành Công)**:
  - **Given** thông tin tài khoản hợp lệ đã được lưu trong Windows Credential Manager thông qua DPAPI dưới Service `PM_SOSANHBOM_TC2412`,
  - **When** automation client khởi tạo WebDriver và truy cập `http://tcmp3gwb:3000/`,
  - **Then** client tự động điền `username` vào `input[name="username"]`, `password` vào `input[name="password"]`, bấm nút submit `button[type="submit"]`, và hoàn thành đăng nhập vào trang chủ (phát hiện `header.sw-row.aw-layout-header` hoặc `button[command-id="Awp0ShowAlertWithBubble"]`) trong thời gian tối đa 20 giây.

- **AC-AUTH-02 (Tài Khoản Không Hợp Lệ - Không Lưu Sai)**:
  - **Given** người dùng nhập thông tin đăng nhập sai hoặc tài khoản bị khóa trên máy chủ Teamcenter,
  - **When** form đăng nhập gửi dữ liệu xác thực,
  - **Then** màn hình hiển thị thông báo lỗi (phát hiện `.aw-login-error` hoặc `.sw-login-error`), client không lưu mật khẩu sai vào Keyring, ghi log cảnh báo an toàn (mật khẩu bị che `***`), và ném ngoại lệ `AuthenticationError`.

- **AC-AUTH-03 (Quản Lý Vòng Đời Credentials: Update & Delete)**:
  - **Given** tài khoản đã tồn tại trong Keyring,
  - **When** người dùng gọi lệnh cập nhật (`--update-credentials`) hoặc lệnh xóa (`--delete-credentials`),
  - **Then** hàm tương ứng của `CredentialManager` cập nhật hoặc xóa khóa trong Windows DPAPI thành công và phiên kế tiếp phản ánh đúng trạng thái lưu trữ mới.

- **AC-AUTH-04 (Tự Động Làm Mới Khi Session Hết Hạn)**:
  - **Given** cookie phiên cũ còn lưu nhưng máy chủ trả về redirect sang form đăng nhập (`div.signin-form`),
  - **When** client phát hiện URL chuyển về Login hoặc xuất hiện trường `input[name="username"]`,
  - **Then** client xóa cache phiên cũ và kích hoạt lại quy trình Silent Auto-login ngay lập tức.

- **AC-AUTH-05 (Bảo Mật Plaintext)**:
  - **Given** quy trình xác thực đang chạy,
  - **When** kiểm tra nhật ký log, các file tạm trên ổ đĩa, và crash dump,
  - **Then** chuỗi mật khẩu không bao giờ xuất hiện dưới dạng văn bản thuần (plaintext) tại bất kỳ vị trí nào.

### 4.3. Input / Output Schemas

#### Pydantic Schema (Python)
```python
from pydantic import BaseModel, Field, HttpUrl
from typing import Optional, Dict

class AuthInputSchema(BaseModel):
    base_url: HttpUrl = Field(default="http://tcmp3gwb:3000/", description="URL cổng TC2412")
    username: Optional[str] = Field(None, description="Tên tài khoản (nếu None sẽ lấy từ DPAPI)")
    password: Optional[str] = Field(None, description="Mật khẩu (nếu None sẽ lấy từ DPAPI)")
    save_credential: bool = Field(default=True, description="Lưu vào Windows DPAPI nếu đăng nhập thành công")
    force_reauth: bool = Field(default=False, description="Bắt buộc đăng nhập lại dù cookie phiên còn hiệu lực")
    timeout_seconds: int = Field(default=20, ge=5, le=120, description="Thời gian chờ tối đa cho đăng nhập")

class AuthOutputSchema(BaseModel):
    success: bool = Field(..., description="Trạng thái đăng nhập thành công")
    authenticated_user: str = Field(..., description="Tên tài khoản đã đăng nhập")
    session_cookies: Dict[str, str] = Field(default_factory=dict, description="Cookie phiên làm việc")
    redirect_url: str = Field(..., description="URL sau khi đăng nhập thành công")
    duration_seconds: float = Field(..., description="Thời gian thực thi đăng nhập tính bằng giây")
    error_message: Optional[str] = Field(None, description="Thông báo lỗi nếu thất bại")
```

#### Hợp Đồng Lớp Lưu Trữ Tài Khoản (CredentialManager Interface)

Tuân thủ nghiêm ngặt đặc tả tại `PROJECT.md:65-71`, `CredentialManager` cung cấp các phương thức tĩnh (static methods) độc lập, chịu trách nhiệm lưu trữ và truy xuất thông tin tài khoản an toàn thông qua Windows DPAPI (`CryptProtectData`/`CryptUnprotectData`) với fallback an toàn sang thư viện `keyring`:

```python
class CredentialManager:
    """Quản lý thông tin xác thực an toàn bằng Windows DPAPI và OS Keyring."""
    
    SERVICE_NAME: str = "PM_SOSANHBOM_TC2412"

    @staticmethod
    def get_credentials(service: str = "PM_SOSANHBOM_TC2412") -> tuple[str, str] | None:
        """Truy xuất cặp (username, password) từ Windows Credential Manager / DPAPI.
        
        Args:
            service: Tên định danh dịch vụ bảo mật (mặc định: 'PM_SOSANHBOM_TC2412')
            
        Returns:
            tuple[str, str] | None: Cặp (username, password) nếu tồn tại, None nếu chưa được lưu
            
        Raises:
            KeyringAccessError: Khi Windows DPAPI hoặc Keyring bị khóa bởi chính sách bảo mật OS
        """
        ...

    @staticmethod
    def save_credentials(username: str, password: str, service: str = "PM_SOSANHBOM_TC2412") -> bool:
        """Mã hóa và lưu trữ an toàn username và password vào Windows DPAPI.
        
        Args:
            username: Tên tài khoản TC2412 (ví dụ: 'vn_pe03')
            password: Mật khẩu plaintext (chỉ tồn tại trong RAM trong thời gian mã hóa)
            service: Tên định danh dịch vụ bảo mật (mặc định: 'PM_SOSANHBOM_TC2412')
            
        Returns:
            bool: True nếu mã hóa và lưu thành công, False nếu thất bại
        """
        ...

    @staticmethod
    def delete_credentials(service: str = "PM_SOSANHBOM_TC2412") -> bool:
        """Thu hồi và xóa vĩnh viễn thông tin tài khoản đã lưu khỏi Windows DPAPI.
        
        Args:
            service: Tên định danh dịch vụ bảo mật (mặc định: 'PM_SOSANHBOM_TC2412')
            
        Returns:
            bool: True nếu xóa thành công hoặc dịch vụ chưa từng tồn tại, False nếu gặp lỗi
        """
        ...
```

**Đặc tính An toàn Thông tin của CredentialManager**:
- **Cơ chế DPAPI Native**: Sử dụng Windows Data Protection API gắn liền với SID của người dùng hiện hành (`CurrentUser`), đảm bảo không tài khoản nào khác trên cùng máy tính có thể giải mã được.
- **Không Lưu Plaintext**: Tuyệt đối không ghi mật khẩu vào registry, file cấu hình `.ini`, `.json`, hay biến môi trường lâu dài.
- **An toàn Tuyệt Đối Khi Gỡ Bỏ**: Gọi `delete_credentials` sẽ xóa sạch key trong Credential Vault, đưa hệ thống về trạng thái yêu cầu đăng nhập thủ công an toàn.

#### JSON Schema
```json
{
  "$schema": "http://json-schema.org/draft-07/schema#",
  "title": "AuthInputSchema",
  "type": "object",
  "properties": {
    "base_url": { "type": "string", "format": "uri", "default": "http://tcmp3gwb:3000/" },
    "username": { "type": ["string", "null"] },
    "password": { "type": ["string", "null"] },
    "save_credential": { "type": "boolean", "default": true },
    "force_reauth": { "type": "boolean", "default": false },
    "timeout_seconds": { "type": "integer", "minimum": 5, "maximum": 120, "default": 20 }
  },
  "required": ["base_url"]
}
```

### 4.4. Verified DOM Locators Table (TC2412 Standard)

| UI Element | Tag | CSS Selector | XPath Expression | Identifying Attributes | Fallback Selector | Strategy / Notes |
|:---|:---|:---|:---|:---|:---|:---|
| **Login Container** | `div` | `div.signin-form.aw-layout-mainView` | `//div[contains(@class,'signin-form')]` | `class="signin-form"` | `div.aw-layout-loginContent` | Explicit Wait: 15s |
| **Version Label** | `p` | `p.aw-login-copyrightTitle` | `//p[contains(@class,'aw-login-copyrightTitle')]` | text: `"Version 2412"` | `//p[contains(text(),'2412')]` | Khẳng định bản TC2412 |
| **Username Input** | `input` | `input[name="username"]` | `//input[@name='username']` | `name="username"`, `aria-label="username"` | `input[data-locator="User Name"]` | `clear()` trước khi `send_keys()` |
| **Password Input** | `input` | `input[name="password"]` | `//input[@name='password']` | `name="password"`, `type="password"` | `input[data-locator="Password"]` | Che ký tự, không log chuỗi |
| **Sign In Button** | `button` | `div.aw-login-signInButton button[type="submit"]` | `//div[contains(@class,'aw-login-signInButton')]//button[@type='submit']` | `type="submit"`, class `sw-button accent-caution` | `button[type="submit"]` | Bấm khi form hết class disabled |
| **Progress Spinner** | `div` | `div.aw-login-progressContainer` | `//div[contains(@class,'aw-login-progressContainer')]` | `class="aw-login-progressContainer"` | `div.aw-layout-progressBar` | Chờ biến mất sau khi submit |
| **Post-Login Banner**| `header`| `header.sw-row.aw-layout-header` | `//header[contains(@class,'aw-layout-header')]` | `role="banner"` | `button[command-id="Awp0ShowAlertWithBubble"]` | Chỉ dấu đăng nhập thành công |

### 4.5. Edge Cases & Fail-Closed Guardrails
1. **Khóa tài khoản trên Active Directory / TC2412**: Khi hiển thị lỗi đăng nhập quá số lần, hệ thống dừng tiến trình lập tức, không cố thử lại để tránh khóa vĩnh viễn tài khoản domain của nhà máy.
2. **Mất kết nối mạng LAN (`ERR_CONNECTION_REFUSED`)**: Thử lại 3 lần với khoảng cách 5s, nếu vẫn mất kết nối ném `PLMNetworkUnreachableError`.
3. **Mật khẩu chứa ký tự đặc biệt**: Xử lý an toàn chuỗi raw UTF-8 khi truyền vào input field của Selenium, không qua URL query.

### 4.6. Error Handling & Recovery Matrix

| Tình huống lỗi | Mã lỗi ngoại lệ | Hành động khôi phục (Recovery Strategy) | Nhật ký & Thông báo người dùng |
|:---|:---|:---|:---|
| Mật khẩu sai / Hết hạn | `AuthenticationError` | Không thử lại tự động; xóa bộ nhớ tạm mật khẩu; yêu cầu người dùng cấu hình lại | Log ERROR: `"Đăng nhập TC2412 thất bại. Vui lòng kiểm tra lại tài khoản."` |
| Máy chủ TC2412 sập | `PLMNetworkUnreachableError` | Thử lại tối đa 3 lần với khoảng cách 5 giây; nếu không được thì dừng | Log ERROR: `"Không thể kết nối tới máy chủ http://tcmp3gwb:3000/. Lỗi mạng."` |
| Keyring / DPAPI bị khóa | `KeyringAccessError` | Yêu cầu nhập mật khẩu thủ công qua CLI prompt | Log WARN: `"Không thể truy cập DPAPI Windows. Yêu cầu nhập mật khẩu thủ công."` |

---

## 5. PHASE 2: BOM-SEARCH-01 – TÌM KIẾM ITEM & ĐIỀU HƯỚNG TAB CONTENT

### 5.1. User Story
```
As a BOM Analyst,
I want the automation system to accurately search for an Item ID in Teamcenter 2412 and navigate straight into its Content tab,
So that the correct product structure is loaded without picking the wrong revision or failing when search results are delayed.
```

### 5.2. Acceptance Criteria (Given-When-Then)

- **AC-SEARCH-01 (Tìm Kiếm Item ID Hợp Lệ & Vào Tab Content)**:
  - **Given** phiên làm việc đã đăng nhập thành công tại trang chủ TC2412,
  - **When** client nhập chuỗi mã sản phẩm (ví dụ: `110C103NL0`) vào ô `input.aw-uiwidgets-searchBox[name="searchBox"]` và gửi lệnh tìm kiếm,
  - **Then** danh sách kết quả `ul.aw-widgets-cellListWidget` xuất hiện, client tìm thấy phần tử có tiêu đề trùng khớp, click mở chi tiết và click vào tab `Content` (`a[data-locator="tab-tc_xrt_Content"]`), chuyển tab sang trạng thái active (`sw-tab-selected`) trong vòng 25 giây.

- **AC-SEARCH-02 (Xử Lý Mã Không Tồn Tại - Fail-Closed)**:
  - **Given** người dùng truyền vào mã Item không có trên hệ thống PLM,
  - **When** tìm kiếm hoàn tất sau 30 giây,
  - **Then** client phát hiện danh sách kết quả trống hoặc nhãn `.aw-widgets-noResultsLabel`, không thực hiện bất kỳ thao tác click nào tiếp theo và ném ngoại lệ `ItemNotFoundError`.

- **AC-SEARCH-03 (Phân Giải Đa Revision Chính Xác)**:
  - **Given** kết quả tìm kiếm trả về nhiều revision khác nhau (ví dụ: Rev A, Rev B, Rev 01),
  - **When** tham số `part_rev` được chỉ định (ví dụ: `"01"`),
  - **Then** client chọn đúng dòng có thuộc tính revision khớp `"01"`; nếu `part_rev` không được chỉ định, client tự động chọn dòng có `Release Status` mới nhất (Latest Released Revision).

- **AC-SEARCH-04 (Điều Hướng URL Trực Tiếp Khi Cần)**:
  - **Given** giao diện tìm kiếm qua UI bị trễ do mạng,
  - **When** client kích hoạt chế độ điều hướng nhanh,
  - **Then** client có khả năng điều hướng trực tiếp qua URL hash route `#/teamcenter.search.search?searchCriteria=<MÃ_ITEM>` để nạp kết quả tin cậy.

### 5.3. Input / Output Schemas

#### Pydantic Schema (Python)
```python
from pydantic import BaseModel, Field
from typing import Optional

class SearchInputSchema(BaseModel):
    item_id: str = Field(..., min_length=3, max_length=50, description="Mã sản phẩm / Item ID cần tìm (e.g. 110C103NL0)")
    target_revision: Optional[str] = Field(None, description="Ký hiệu Revision cụ thể (e.g. '01', 'A'). Nếu None chọn bản mới nhất")
    search_timeout_seconds: int = Field(default=30, ge=10, le=120, description="Timeout tối đa cho quá trình tìm kiếm")

class SearchOutputSchema(BaseModel):
    item_id: str = Field(..., description="Mã Item ID đã tìm thấy")
    selected_revision: str = Field(..., description="Revision đã được chọn")
    item_uid: Optional[str] = Field(None, description="UID nội bộ của Item trên Teamcenter nếu trích xuất được")
    content_tab_loaded: bool = Field(..., description="Xác nhận tab Content đã được nạp thành công")
    tree_container_ready: bool = Field(..., description="Bảng cây occTreeTable đã sẵn sàng nhận lệnh")
    duration_seconds: float = Field(..., description="Thời gian thực thi tìm kiếm và điều hướng")
```

### 5.4. Verified DOM Locators Table (TC2412 Standard)

| UI Element | Tag | CSS Selector | XPath Expression | Identifying Attributes | Fallback Selector | Strategy / Notes |
|:---|:---|:---|:---|:---|:---|:---|
| **Search Box Container** | `div` | `div.aw-search-globalSearchBoxContainer[role="search"]` | `//div[contains(@class,'aw-search-globalSearchBoxContainer') and @role='search']` | `role="search"` | `div.aw-uiwidgets-searchBoxContainer` | Chờ sẵn sàng |
| **Search Input Box** | `input` | `input.aw-uiwidgets-searchBox[name="searchBox"]` | `//input[@name='searchBox' and contains(@class,'aw-uiwidgets-searchBox')]` | `name="searchBox"`, `placeholder="Search"` | `input.aw-uiwidgets-searchBox` | `clear()` -> gõ ký tự |
| **Search Action Icon** | `span` | `span.aw-uiwidgets-searchBoxIcon` | `//span[contains(@class,'aw-uiwidgets-searchBoxIcon')]` | `icon-id="cmdSearch"` | `Keys.ENTER` trên input | Gửi phím Enter |
| **Search Results List**| `ul` | `ul.aw-widgets-cellListWidget` | `//ul[contains(@class,'aw-widgets-cellListWidget')]` | `class="aw-widgets-cellListWidget"` | `div.aw-widgets-cellListContainer` | Explicit wait: 30s |
| **Search Result Item** | `li` | `li.aw-widgets-cellListItem[role="option"]` | `//li[contains(@class,'aw-widgets-cellListItem') and @role='option']` | `role="option"` | `li.aw-widgets-cellListItem` | Đại diện cho 1 Item Rev |
| **Item Title Link** | `a` | `a.aw-widgets-cellListCellTitle[data-locator="aw-clickable-title"]` | `//a[@data-locator='aw-clickable-title' and contains(@class,'aw-widgets-cellListCellTitle')]` | `data-locator="aw-clickable-title"` | `a.aw-widgets-cellListCellTitle` | Click để vào Item View |
| **Open Item Button** | `button`| `button[command-id="Awp0ShowObjectCell"]` | `//button[@command-id='Awp0ShowObjectCell']` | `command-id="Awp0ShowObjectCell"` | `button.aw-commandId-Awp0ShowObjectCell` | Nút mở nhanh |
| **Content Tab Link** | `a` | `a[data-locator="tab-tc_xrt_Content"]` | `//a[@data-locator='tab-tc_xrt_Content']` | `data-locator="tab-tc_xrt_Content"` | `//li[contains(@class,'sw-tab')]//span[text()='Content']` | Chờ class `sw-tab-selected` |

### 5.5. Edge Cases & Fail-Closed Guardrails
1. **Ô tìm kiếm bị ẩn hoặc trễ hiển thị**: Chờ phần tử `div.aw-search-globalSearchBoxContainer` có `aria-disabled="false"`.
2. **Ký tự mã Item có khoảng trắng thừa**: Chuỗi tìm kiếm tự động gọi `.strip()` để tránh sinh kết quả rỗng.
3. **Màn hình hiển thị "No results found"**: Dừng ngay lập tức, không bấm các nút điều hướng khác.

### 5.6. Error Handling & Recovery Matrix

| Tình huống lỗi | Mã lỗi ngoại lệ | Hành động khôi phục | Nhật ký & Thông báo người dùng |
|:---|:---|:---|:---|
| Không tìm thấy mã Item | `ItemNotFoundError` | Dừng tiến trình, trả mã lỗi ngay | Log ERROR: `"Mã linh kiện '{item_id}' không tồn tại trên hệ thống TC2412."` |
| Quá hạn nạp kết quả Search | `SearchTimeoutError` | Thử lại tìm kiếm qua URL hash route 1 lần; nếu vẫn timeout thì ném lỗi | Log WARN: `"Tìm kiếm qua giao diện bị timeout, đang thử điều hướng trực tiếp bằng URL..."` |
| Không thấy Tab Content | `ContentTabNotFoundError` | Kiểm tra đối tượng có phải BOM assembly hay không; ném ngoại lệ nếu sai loại | Log ERROR: `"Đối tượng '{item_id}' không có tab Content (không phải đối tượng BOM)."` |

---

## 6. PHASE 3: BOM-EXPAND-01 – BUNG CÂY CẤU TRÚC ĐẾN LEVEL 7

### 6.1. User Story
```
As a Production Engineering Data Specialist,
I want the automation system to select the root BOM node and expand the hierarchical tree to Level 7 with headless dynamic waits,
So that all sub-assemblies and deep child parts (down to 7 levels) are fully loaded into memory without missing lines or failing on massive trees.
```

### 6.2. Acceptance Criteria (Given-When-Then)

- **AC-EXPAND-01 (Chọn Nút Gốc & Bung Cây Level 7 Thành Công)**:
  - **Given** người dùng đang ở tab `Content` với bảng cây `div.aw-splm-table#occTreeTable` đã hiển thị dòng gốc `data-indexnumber="0"`,
  - **When** client click vào dòng gốc để nhận trạng thái `aria-selected="true"`, bấm menu Expand `button[command-id="Awb0Expand"]`, chọn `"Expand Below"`, điền giá trị `7` vào ô nhập level và bấm xác nhận,
  - **Then** cây BOM bắt đầu bung dữ liệu, thanh tiến trình `.aw-layout-progressBar` chuyển `aria-busy="true"` rồi về `aria-busy="false"` + `hide`, và tổng số dòng `div.aw-splm-tableRow` tăng lên đầy đủ các cấp từ 0 đến 7.

- **AC-EXPAND-02 (Ổn Định Trong Chế Độ Headless - Chống Sập & Tránh Bẫy Virtual DOM)**:
  - **Given** trình duyệt đang chạy ở chế độ Headless (`--window-size=1920,1080`),
  - **When** cấu trúc cây BOM có dung lượng cực lớn (> 5.000 dòng) và Teamcenter 2412 áp dụng cơ chế Virtual DOM (chỉ render ~30-50 dòng trong viewport hiển thị),
  - **Then** client thực hiện dynamic explicit wait, theo dõi trạng thái ổn định của tổng số dòng hiển thị tại `div.aw-splm-tableSummary` và trạng thái tắt hẳn của spinner tiến trình `.aw-layout-progressBar` (`aria-busy="false"` + `hide`) trong 3 chu kỳ liên tiếp mỗi 1.0 giây với timeout tối đa 180 giây mà không bị crash trình duyệt. Client tuyệt đối không giả định toàn bộ hàng ngàn dòng BOM đều đồng thời tồn tại dưới dạng phần tử `div.aw-splm-tableRow` trong DOM.

- **AC-EXPAND-03 (Xử Lý Nhánh Bị Khóa Quyền Truy Cập - Partial Warning)**:
  - **Given** một số chi tiết con trong cây bị giới hạn quyền truy cập (Access Denied),
  - **When** quá trình bung cây chạm tới nhánh đó,
  - **Then** client ghi nhận cảnh báo mức WARNING trong log và tiếp tục mở rộng toàn bộ các nhánh hợp lệ còn lại mà không ngắt quy trình.

- **AC-EXPAND-04 (Bảo Toàn Thứ Tự Phân Cấp Cha-Con)**:
  - **Given** cây đã mở rộng xong,
  - **When** kiểm tra thuộc tính `aria-level` trên các dòng bảng,
  - **Then** các dòng con xuất hiện ngay dưới dòng cha tương ứng, bảo toàn tuyệt đối quan hệ thứ tự cây nguyên bản.

### 6.3. Input / Output Schemas

#### Pydantic Schema (Python)
```python
from pydantic import BaseModel, Field

class ExpandInputSchema(BaseModel):
    target_level: int = Field(default=7, ge=1, le=10, description="Độ sâu cần bung cây BOM (mặc định Level 7)")
    timeout_seconds: int = Field(default=180, ge=30, le=600, description="Thời gian chờ tối đa cho cây lớn")
    stability_cycles: int = Field(default=3, ge=2, le=10, description="Số chu kỳ số dòng không đổi để xác định nạp xong")
    stability_interval_seconds: float = Field(default=1.0, ge=0.5, le=5.0, description="Khoảng cách giữa các chu kỳ kiểm tra")

class ExpandOutputSchema(BaseModel):
    root_item_id: str = Field(..., description="Mã linh kiện của nút gốc")
    expanded_level: int = Field(..., description="Mức độ sâu thực tế đã yêu cầu")
    total_rows_loaded: int = Field(..., description="Tổng số dòng BOM đã nạp trên cây")
    max_depth_detected: int = Field(..., description="Độ sâu lớn nhất thực tế phát hiện trong DOM")
    duration_seconds: float = Field(..., description="Thời gian bung cây tính bằng giây")
    warning_branches: list[str] = Field(default_factory=list, description="Danh sách các nhánh gặp lỗi phân quyền nếu có")
```

### 6.4. Verified DOM Locators Table (TC2412 Standard)

| UI Element | Tag | CSS Selector | XPath Expression | Identifying Attributes | Fallback Selector | Strategy / Notes |
|:---|:---|:---|:---|:---|:---|:---|
| **BOM Tree Table** | `div` | `div.aw-splm-table#occTreeTable` | `//div[@id='occTreeTable' and contains(@class,'aw-splm-table')]` | `id="occTreeTable"`, `role="grid"` | `div.aw-splm-table` | Khung bảng ảo chính |
| **Root BOM Row** | `div` | `div.aw-splm-tableRow[data-indexnumber="0"]` | `//div[@id='occTreeTable']//div[contains(@class,'aw-splm-tableRow') and @data-indexnumber='0']` | `data-indexnumber="0"`, `aria-level="1"` | `div.aw-splm-tableRow:first-child` | Click để chọn nút gốc |
| **Root Cell Link** | `a` | `.aw-splm-tableCellText a.aw-uiwidgets-clickableTitle` | `//div[@id='occTreeTable']//div[@data-indexnumber='0']//a[contains(@class,'aw-uiwidgets-clickableTitle')]` | `title` khớp mã Item gốc | `a.aw-uiwidgets-clickableTitle` | Vị trí click nhận diện |
| **Expand Toolbar Button** | `button` | `button[command-id="Awb0Expand"]` | `//button[@command-id='Awb0Expand' or @button-id='Awb0Expand']` | `command-id="Awb0Expand"`, `aria-haspopup="true"` | `button.aw-commandId-Awb0Expand` | Mở dropdown menu |
| **Expand Below Item** | `div` | `div[command-id="Awb0ExpandBelow"]` | `//*[contains(text(),'Expand Below') or @command-id='Awb0ExpandBelow']` | `command-id="Awb0ExpandBelow"` | `div.aw-widgets-cellListItem[title*="Expand Below"]` | Mở dialog nhập level |
| **Level Input Field** | `input` | `input.aw-widgets-propertyVal[type="number"]` | `//input[@type='number' or contains(@aria-label,'Level')]` | `type="number"`, `aria-label*="Level"` | `input[type="number"]` | Xóa và điền `7` |
| **Confirm Expand Button** | `button` | `button.sw-button[type="button"]` | `//div[contains(@class,'aw-popup')]//button[contains(.,'Expand') or contains(.,'OK')]` | Text: `"Expand"` hoặc `"OK"` | `//button[text()='Expand']` | Bấm để bắt đầu bung |
| **Loading Progress Bar** | `div` | `div.aw-layout-progressBar` | `//div[contains(@class,'aw-layout-progressBar')]` | `role="progressbar"`, `aria-busy` | `div.aw-layout-progressBarContainer` | Chờ `aria-busy="false"` + `hide` |
| **Table Summary Indicator** | `div` | `div.aw-splm-tableSummary` | `//div[contains(@class,'aw-splm-tableSummary')]` | `class="aw-splm-tableSummary"` | `div.aw-splm-tableStatus` | Đọc số dòng đã nạp / trạng thái nạp (chống bẫy Virtual DOM) |

### 6.5. Edge Cases & Fail-Closed Guardrails
1. **Nút Expand bị mờ (`disabled="true"`)**: Xảy ra khi dòng gốc chưa được chọn. Client thực hiện tái chọn dòng gốc bằng cách click vào ô text của nút gốc và kiểm tra lại thuộc tính `aria-selected="true"`.
2. **Cây lớn kích hoạt DOM Virtualization (~30-50 nodes rendered)**: Active Workspace 2412 chỉ giữ các node trong viewport trên DOM cây `occTreeTable`. Khi kiểm tra số lượng dòng hoặc bung cây, tuyệt đối không dùng `len(find_elements(div.aw-splm-tableRow))` làm tiêu chí so khớp số lượng tổng. Thay vào đó, đọc chỉ số tổng từ `div.aw-splm-tableSummary` và theo dõi thanh tiến trình biến mất. Khi cần cuộn để kích hoạt render, sử dụng `execute_script("arguments[0].scrollIntoView(true);", elem)`.
3. **Cây BOM rỗng không có linh kiện con**: Nếu sau khi mở rộng số dòng vẫn bằng 1, client ghi nhận cảnh báo BOM đơn cấp và cho phép xuất tiếp.

### 6.6. Error Handling & Recovery Matrix

| Tình huống lỗi | Mã lỗi ngoại lệ | Hành động khôi phục | Nhật ký & Thông báo người dùng |
|:---|:---|:---|:---|
| Không chọn được nút gốc | `RootNodeSelectionError` | Thử lại click chọn nút gốc tối đa 3 lần qua tọa độ dòng | Log ERROR: `"Không thể chọn nút gốc của cây BOM trên TC2412."` |
| Quá hạn bung cây (> 180s) | `BOMExpandTimeoutError` | Chụp ảnh màn hình headless debug, ghi lại số dòng đã tải và ném lỗi | Log ERROR: `"Hết thời gian chờ bung cây BOM (180s). Đã tải được {rows} dòng."` |
| Hộp thoại Expand không mở | `ExpandMenuNotOpenError` | Bấm lại nút Expand trên thanh công cụ | Log WARN: `"Menu Expand chưa mở, đang thử bấm lại..."` |

---

## 7. PHASE 4: BOM-SELECT-01 – CHỌN TOÀN BỘ CÂY BOM & MỞ MENU EXCEL REPORT

### 7.1. User Story
```
As an Automation Runner,
I want to select all loaded BOM lines using the Select All command and open the Excel Report menu,
So that 100% of the expanded hierarchy is marked for export and the selection is preserved when the menu appears.
```

### 7.2. Acceptance Criteria (Given-When-Then)

- **AC-SELECT-01 (Select All 100% Dòng BOM - Virtual DOM Resilient)**:
  - **Given** cây BOM đã hoàn thành nạp đến cấp 7 và trạng thái loading đã tắt hoàn toàn,
  - **When** client click nút `Select All` (`button[command-id="Awp0SelectAll"]`),
  - **Then** Active Workspace cập nhật mô hình dữ liệu chọn. Do cơ chế DOM Virtualization chỉ render ~30-50 dòng trong viewport, client xác nhận toàn bộ cây được chọn bằng cách:
    1. Đọc bộ đếm chỉ số lựa chọn `div.aw-splm-tableSelectionCount` (hoặc nhãn `div.aw-splm-tableSummary` hiển thị chuỗi như `"<N> Selected"` với N > 0 và bằng tổng số dòng đã nạp);
    2. Kiểm tra ô checkbox chọn tất cả trên thanh tiêu đề bảng `div.aw-splm-tableHeaderCheckbox` có trạng thái `aria-checked="true"` hoặc `checked`;
    3. Và kiểm tra toàn bộ các dòng hiện đang được render trong viewport (`div.aw-splm-tableRow`) đều có thuộc tính `aria-selected="true"`.
    Client tuyệt đối không so sánh `len(rendered_dom_rows) == total_rows_loaded` vì các dòng nằm ngoài viewport không tồn tại trong DOM.

- **AC-SELECT-02 (Mở Menu Excel Report Duy Trì Selection)**:
  - **Given** toàn bộ các dòng BOM đang ở trạng thái được chọn,
  - **When** client click nút `Excel Round-trip / Excel Report` (`button[command-id="Arm0ExportImport"]`),
  - **Then** menu xổ xuống xuất hiện hiển thị các tùy chọn xuất nhập, đồng thời toàn bộ tập dòng BOM vẫn duy trì trạng thái `aria-selected="true"` mà không bị bỏ chọn.

- **AC-SELECT-03 (Bảo Vệ Không Chọn Dòng Nào - Fail-Closed)**:
  - **Given** vì sự cố DOM số dòng được chọn bằng 0,
  - **When** chuẩn bị mở menu Excel Report,
  - **Then** client dừng lại ngay lập tức và ném ngoại lệ `NoRowsSelectedError`, không tiếp tục bước xuất dữ liệu rỗng.

### 7.3. Input / Output Schemas

#### Pydantic Schema (Python)
```python
from pydantic import BaseModel, Field

class SelectInputSchema(BaseModel):
    expected_min_rows: int = Field(default=1, ge=1, description="Số lượng dòng tối thiểu kỳ vọng được chọn")
    timeout_seconds: int = Field(default=30, ge=5, le=60, description="Thời gian chờ tối đa cho thao tác chọn")

class SelectOutputSchema(BaseModel):
    total_rows_selected: int = Field(..., ge=1, description="Tổng số dòng được chọn (đọc từ div.aw-splm-tableSelectionCount hoặc summary indicator)")
    selection_verified: bool = Field(..., description="Khẳng định chọn thành công qua selection count indicator và header checkbox aria-checked='true'")
    menu_opened: bool = Field(..., description="Menu Excel Round-trip đã mở thành công")
    duration_seconds: float = Field(..., description="Thời gian thực hiện thao tác chọn")
```

### 7.4. Verified DOM Locators Table (TC2412 Standard)

| UI Element | Tag | CSS Selector | XPath Expression | Identifying Attributes | Fallback Selector | Strategy / Notes |
|:---|:---|:---|:---|:---|:---|:---|
| **Multi-Select Enable** | `button` | `button[command-id="Awp0EnableMultiSelect"]` | `//button[@command-id='Awp0EnableMultiSelect']` | `command-id="Awp0EnableMultiSelect"` | `button[aria-label="Selection Mode"]` | Kích hoạt nếu đang ở single-select |
| **Select All Button** | `button` | `button[command-id="Awp0SelectAll"]` | `//button[@command-id='Awp0SelectAll' or @button-id='Awp0SelectAll']` | `command-id="Awp0SelectAll"`, `aria-label="Select All"` | `button.aw-commandId-Awp0SelectAll` | Icon `cmdCheckmark` |
| **Table Selection Count** | `div` | `div.aw-splm-tableSelectionCount` | `//div[contains(@class,'aw-splm-tableSelectionCount')]` | `class="aw-splm-tableSelectionCount"` | `div.aw-splm-tableSummary` | Đọc số dòng đã chọn (e.g. `"2548 Selected"`) - Tiêu chuẩn Virtual DOM |
| **Table Header Checkbox** | `div` | `div.aw-splm-tableHeaderCheckbox, input[type="checkbox"].aw-splm-tableHeaderCheckbox` | `//div[contains(@class,'aw-splm-tableHeaderCheckbox')]` | `aria-checked="true"` | `input.aw-splm-tableHeaderCheckbox` | Khẳng định trạng thái Select All toàn bảng |
| **Selected Row Marker**| `div` | `div.aw-splm-tableRow[aria-selected="true"]` | `//div[contains(@class,'aw-splm-tableRow') and @aria-selected='true']` | `aria-selected="true"` | `.aw-splm-tableRowSelected` | Chỉ kiểm tra các dòng đang render trong viewport (~30-50 nodes) |
| **Excel Report Button**| `button` | `button[command-id="Arm0ExportImport"]` | `//button[@command-id='Arm0ExportImport' or @button-id='Arm0ExportImport']` | `command-id="Arm0ExportImport"`, `aria-label="Excel Round-trip"` | `button.aw-commandId-Arm0ExportImport` | Mở danh sách lệnh Excel |
| **Popup Command Menu** | `div` | `div.aw-popup-commandListContainer` | `//div[contains(@class,'aw-popup-commandListContainer')]` | `class="aw-popup-commandListContainer"` | `div.aw-popup.sw-popup` | Chứa Export và Import |

### 7.5. Edge Cases & Fail-Closed Guardrails
1. **Bẫy Virtual DOM khi đếm dòng chọn**: Trong Teamcenter Active Workspace 2412, `occTreeTable` áp dụng kỹ thuật virtual scrolling. Nếu cây BOM có 2.500 dòng, số phần tử `div.aw-splm-tableRow` tồn tại trong DOM tại một thời điểm chỉ từ 30-50 phần tử. Nếu kiểm tra bằng `len(driver.find_elements(By.CSS_SELECTOR, "div.aw-splm-tableRow[aria-selected='true']")) == 2500` sẽ luôn gây false negative và làm hỏng test. Tiêu chí xác thực chuẩn là đọc số dòng chọn từ `div.aw-splm-tableSelectionCount` (hoặc `div.aw-splm-tableSummary`) và kiểm tra `aria-checked="true"` trên checkbox header `div.aw-splm-tableHeaderCheckbox`.
2. **Click trượt nút Select All**: Sau khi click, luôn kiểm tra chỉ số `div.aw-splm-tableSelectionCount` hoặc checkbox header. Nếu chỉ số bằng 0 hoặc chưa check, thử lại với Explicit Wait 2s.
3. **Mất chọn khi popup mở**: Nếu việc click menu Excel làm mất trạng thái chọn trên mô hình dữ liệu, client lập tức đóng menu, bấm lại Select All rồi mới mở lại menu.

### 7.6. Error Handling & Recovery Matrix

| Tình huống lỗi | Mã lỗi ngoại lệ | Hành động khôi phục | Nhật ký & Thông báo người dùng |
|:---|:---|:---|:---|
| Không có dòng nào được chọn | `NoRowsSelectedError` | Thử lại bấm Select All tối đa 3 lần; nếu không được dừng lại | Log ERROR: `"Lỗi chọn dòng: Không có dòng BOM nào được đánh dấu chọn."` |
| Menu Excel không hiển thị | `ExcelMenuNotOpenError` | Kiểm tra lại nút `Arm0ExportImport` và click lại | Log WARN: `"Menu Excel Report chưa xuất hiện, đang thử lại..."` |

---

## 8. PHASE 5: BOM-EXPORT-OPEN-01 – KÍCH HOẠT EXPORT TO EXCEL & CHỐT CHẶN IMPORT

### 8.1. User Story
```
As a System Safety Officer and Data Operator,
I want the automation system to trigger the Export to Excel dialog with a strict Fail-Closed anti-import guard,
So that the export configuration panel opens reliably while absolutely guaranteeing that the dangerous Import Changes action is never clicked.
```

### 8.2. Acceptance Criteria (Given-When-Then)

- **AC-EXPORT-OPEN-01 (Mở Đúng Panel Export To Excel)**:
  - **Given** menu lệnh Excel Report đang hiển thị các mục chọn,
  - **When** client quét và click vào mục `"Export to Excel"` (`button[command-id="Awp0ExportToExcel"]`),
  - **Then** panel cấu hình xuất dạng flyout (`div.sw-popup.sw-dialog#ui-id-3` hoặc vùng chứa có tiêu đề `"Export To Excel"`) mở ra trong vòng 10 giây và tập dòng đã chọn được giữ nguyên.

- **AC-EXPORT-OPEN-02 (Fail-Closed: Chặn Đứng Tuyệt Đối "Import Changes")**:
  - **Given** menu hiển thị cả hai mục `"Export to Excel"` và `"Import Changes"` nằm cạnh nhau,
  - **When** thuật toán quét phần tử duyệt qua các mục menu,
  - **Then** bất kỳ phần tử nào có văn bản chứa chuỗi `"import"` (không phân biệt hoa thường) hoặc command-id chứa `"Import"` đều bị bỏ qua ngay lập tức; nếu xảy ra sự cố click nhầm vào Import, hệ thống ngắt toàn bộ tiến trình tức thì và ném ngoại lệ nghiêm trọng `ImportChangesForbiddenError`.

### 8.3. Thuật Toán Phòng Vệ Fail-Closed (Anti-Import Guard Algorithm)
```python
import re
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException

class ImportChangesForbiddenError(RuntimeError):
    """Ngoại lệ nghiêm trọng khi phát hiện nguy cơ kích hoạt Import Changes."""
    pass

def safe_trigger_export_to_excel(driver, timeout: int = 15):
    """Kích hoạt Export to Excel với cơ chế Fail-Closed chống nhầm lẫn Import."""
    # 1. Quét tất cả các phần tử trong danh sách lệnh menu popup
    candidates = driver.find_elements(By.CSS_SELECTOR, ".aw-command, .aw-widgets-cellListItem, button")
    target_btn = None
    
    for elem in candidates:
        text = elem.text.strip().lower()
        cmd_id = (elem.get_attribute("command-id") or elem.get_attribute("button-id") or "").lower()
        title = (elem.get_attribute("title") or "").lower()
        
        # CHỐT CHẶN 1: Nếu chứa chữ 'import' -> TUYỆT ĐỐI BỎ QUA
        if "import" in text or "import" in cmd_id or "import" in title:
            continue
            
        # CHỐT CHẶN 2: Nhận diện chính xác Export to Excel
        if ("export" in text and "excel" in text) or cmd_id == "awp0exporttoexcel" or "export to excel" in title:
            target_btn = elem
            break
            
    if target_btn is None:
        raise NoSuchElementException("Không tìm thấy mục menu 'Export to Excel' an toàn!")
        
    # Xác thực lần cuối trước khi click
    final_cmd = (target_btn.get_attribute("command-id") or "").lower()
    if "import" in final_cmd:
        raise ImportChangesForbiddenError("Phát hiện nguy cơ click nhầm vào Import Changes! Dừng khẩn cấp.")
        
    target_btn.click()
```

### 8.4. Input / Output Schemas

#### Pydantic Schema (Python)
```python
from pydantic import BaseModel, Field

class ExportOpenInputSchema(BaseModel):
    timeout_seconds: int = Field(default=15, ge=5, le=60, description="Thời gian chờ mở panel xuất Excel")

class ExportOpenOutputSchema(BaseModel):
    panel_opened: bool = Field(..., description="Panel cấu hình xuất đã mở thành công")
    panel_id: str = Field(..., description="ID định danh của panel (e.g. ui-id-3)")
    anti_import_guard_passed: bool = Field(..., description="Xác nhận vượt qua chốt chặn an toàn chống Import")
    duration_seconds: float = Field(..., description="Thời gian thực thi")
```

### 8.5. Verified DOM Locators Table (TC2412 Standard)

| UI Element | Tag | CSS Selector | XPath Expression | Identifying Attributes | Fallback Selector | Strategy / Notes |
|:---|:---|:---|:---|:---|:---|:---|
| **Export to Excel Item** | `button` | `button[command-id="Awp0ExportToExcel"]` | `//button[@command-id='Awp0ExportToExcel']` | `command-id="Awp0ExportToExcel"` | `div[title="Export To Excel"]` | Mục tiêu click an toàn |
| **Import Changes (CẤM)** | `button` | `button[command-id*="Import"]` | `//*[contains(text(),'Import Changes') or contains(@command-id,'Import')]` | Chứa text: `"Import"` | KHÔNG BAO GIỜ CLICK | Chốt chặn an toàn |
| **Export Dialog Container** | `div` | `div.sw-popup.sw-dialog#ui-id-3` | `//div[@id='ui-id-3' and contains(@class,'sw-popup')]` | `id="ui-id-3"` | `div.sw-right-dialog, div.aw-layout-popup` | Panel cấu hình flyout |
| **Dialog Caption Header** | `div` | `div.aw-panel-caption` | `//div[contains(@class,'aw-panel-caption') and contains(.,'Export To Excel')]` | text: `"Export To Excel"` | `h2:contains('Export To Excel')` | Kiểm tra tiêu đề panel |

### 8.6. Error Handling & Recovery Matrix

| Tình huống lỗi | Mã lỗi ngoại lệ | Hành động khôi phục | Nhật ký & Thông báo người dùng |
|:---|:---|:---|:---|
| Nghi ngờ click nhầm Import | `ImportChangesForbiddenError` | Ngắt toàn bộ WebDriver ngay lập tức, đóng browser | Log CRITICAL: `"BẢO VỆ AN TOÀN: Phát hiện phần tử Import Changes! Đã hủy tác vụ."` |
| Panel Export không mở | `ExportDialogTimeoutError` | Đóng popup, mở lại menu Excel Report và thử click lại 1 lần | Log WARN: `"Hộp thoại Export To Excel chưa mở sau 15s. Đang thử lại..."` |

---

## 9. PHASE 6: BOM-EXPORT-CONFIG-01 – CẤU HÌNH THUỘC TÍNH ITEM & 14 CỘT HIỂN THỊ

### 9.1. User Story
```
As a BOM Core Engine Integrator,
I want the automation client to select Property Source 'Item', purge incorrect default columns, and configure exactly the 14 canonical columns in strict order,
So that the exported Excel file conforms 100% to the data contract required by downstream reconciliation and pruning algorithms.
```

### 9.2. Acceptance Criteria (Given-When-Then)

- **AC-CONFIG-01 (Chọn Nguồn Thuộc Tính Item)**:
  - **Given** panel `Export To Excel` đã mở,
  - **When** client mở bộ chọn `Property Source` và chọn giá trị `"Item"`,
  - **Then** danh sách Available Columns cập nhật tương ứng với các thuộc tính của đối tượng Item.

- **AC-CONFIG-02 (Xóa Cột Mặc Định Không Phù Hợp)**:
  - **Given** danh sách `Displayed Columns` có chứa các cột mặc định sai của Teamcenter (`Object`, `Is Primary`, `Design Item`, `Owner`...),
  - **When** client quét danh sách hiện hành,
  - **Then** các cột không thuộc danh sách 14 cột bắt buộc đều bị gỡ bỏ thông qua nút Remove.

- **AC-CONFIG-03 (Bổ Sung Đầy Đủ 14 Cột Chuẩn & Lọc Thuộc Tính Tự Động)**:
  - **Given** các cột còn thiếu trong Displayed Columns và danh sách Available Columns chứa hàng trăm thuộc tính của Item,
  - **When** client sử dụng ô tìm kiếm/lọc `input.aw-uiwidgets-searchBox` bên trong panel Available Properties để lọc chính xác từng tên thuộc tính và nhấn nút `Add Properties` (`Awp0ExportSelectedColumnsAdd`),
  - **Then** tất cả các cột trong danh sách 14 cột được tìm thấy nhanh chóng, không bị trượt do cuộn trang ảo, và được thêm đầy đủ vào danh sách xuất.

- **AC-CONFIG-04 (Sắp Xếp Đúng 100% Thứ Tự Từ 1 Đến 14)**:
  - **Given** 14 cột đã có mặt trong Displayed Columns,
  - **When** client sử dụng các nút Move Up (`Awp0MoveUpExcelColumn`) và Move Down (`Awp0MoveDownExcelColumn`),
  - **Then** thứ tự các thẻ `li.aw-widgets-cellListItem` khớp chính xác 100% từ vị trí 1 đến 14 theo bảng quy định tại Mục 3:
    1. Home
    2. Level
    3. Item Type
    4. Item Id
    5. Has Children
    6. Quantity
    7. 1st Parts
    8. 2nd BOM Flag
    9. Occurrence Effectivities
    10. Item Revision Project List
    11. Item Name
    12. Notice No
    13. Revision
    14. Item Rev Status

- **AC-CONFIG-05 (Chống Trùng Lặp & Thiếu Cột - Fail-Closed)**:
  - **Given** cấu hình hoàn tất,
  - **When** bộ kiểm tra đối soát danh sách cột trước khi submit,
  - **Then** số lượng cột phải bằng đúng 14, không có cột trùng lặp; nếu sai lệch, ném ngoại lệ `ColumnConfigurationError` và không tiến hành bước xuất.

### 9.3. Input / Output Schemas

#### Pydantic Schema (Python)
```python
from pydantic import BaseModel, Field
from typing import List

CANONICAL_14_COLUMNS: List[str] = [
    "Home",
    "Level",
    "Item Type",
    "Item Id",
    "Has Children",
    "Quantity",
    "1st Parts",
    "2nd BOM Flag",
    "Occurrence Effectivities",
    "Item Revision Project List",
    "Item Name",
    "Notice No",
    "Revision",
    "Item Rev Status",
]

class ExportConfigInputSchema(BaseModel):
    property_source: str = Field(default="Item", description="Nguồn thuộc tính bắt buộc là Item")
    required_columns: List[str] = Field(default=CANONICAL_14_COLUMNS, description="Danh sách 14 cột chuẩn")
    timeout_seconds: int = Field(default=45, ge=15, le=120, description="Thời gian cấu hình tối đa")

class ExportConfigOutputSchema(BaseModel):
    configured_columns: List[str] = Field(..., description="Danh sách cột thực tế đã cấu hình")
    is_valid_14_columns: bool = Field(..., description="Khẳng định khớp đúng 14/14 cột")
    order_matches_canonical: bool = Field(..., description="Khẳng định đúng 100% thứ tự")
    duration_seconds: float = Field(..., description="Thời gian hoàn tất cấu hình")
```

### 9.4. Verified DOM Locators Table (TC2412 Standard)

| UI Element | Tag | CSS Selector | XPath Expression | Identifying Attributes | Fallback Selector | Strategy / Notes |
|:---|:---|:---|:---|:---|:---|:---|
| **Property Source Selector** | `div` | `.aw-property-source-dropdown, div.sw-lov-container` | `//*[contains(text(),'Property Source')]/following::div[contains(@class,'sw-lov-container')][1]` | `data-locator="Property Source"` | `//div[@aria-label='Property Source']` | Chọn mục `"Item"` |
| **Available Properties Filter** | `input` | `input.aw-uiwidgets-searchBox[placeholder*="Filter"], div.aw-panel-section input.aw-uiwidgets-searchBox` | `//div[contains(@class,'aw-panel-section')]//input[contains(@class,'aw-uiwidgets-searchBox') or @placeholder='Filter']` | `class="aw-uiwidgets-searchBox"`, `placeholder="Filter"` | `input[data-locator="Filter Properties"]` | Gõ tên thuộc tính để lọc nhanh từ hàng trăm mục xuống thuộc tính cần tìm |
| **Add Properties Button** | `button` | `button[command-id="Awp0ExportSelectedColumnsAdd"]` | `//button[@command-id='Awp0ExportSelectedColumnsAdd']` | `command-id="Awp0ExportSelectedColumnsAdd"` | `button[aria-label="Add Properties"]` | Thêm thuộc tính đã chọn |
| **Move Up Button** | `button` | `button[command-id="Awp0MoveUpExcelColumn"]` | `//button[@command-id='Awp0MoveUpExcelColumn']` | `command-id="Awp0MoveUpExcelColumn"` | `button[aria-label="Move Up"]` | Đẩy cột lên 1 vị trí |
| **Move Down Button** | `button` | `button[command-id="Awp0MoveDownExcelColumn"]` | `//button[@command-id='Awp0MoveDownExcelColumn']` | `command-id="Awp0MoveDownExcelColumn"` | `button[aria-label="Move Down"]` | Đẩy cột xuống 1 vị trí |
| **Displayed Columns List** | `details` | `details[caption="Selected Properties"] li.aw-widgets-cellListItem` | `//details[contains(@caption,'Selected Properties')]//li[contains(@class,'aw-widgets-cellListItem')]` | `class="aw-widgets-cellListItem"` | `ul.aw-widgets-cellListWidget` trong Selected | Đọc danh sách cột hiện có |
| **Remove Column Button** | `button` | `li.aw-widgets-cellListItem button[title="Remove"]` | `//li[contains(@class,'aw-widgets-cellListItem')]//button[contains(@title,'Remove') or contains(@aria-label,'Remove')]` | `title="Remove"` | `button[command-id*="Remove"]` | Xóa cột không hợp lệ |

### 9.5. Error Handling & Recovery Matrix

| Tình huống lỗi | Mã lỗi ngoại lệ | Hành động khôi phục | Nhật ký & Thông báo người dùng |
|:---|:---|:---|:---|
| Thiếu thuộc tính trong Available | `PropertyNotFoundError` | Áp dụng bảng ánh xạ alias (ví dụ `Item ID` <-> `Item Id`); nếu vẫn thiếu thì dừng | Log ERROR: `"Không tìm thấy thuộc tính '{col}' trong danh sách Available Columns."` |
| Sai thứ tự cột sau khi xếp | `ColumnOrderMismatchError` | Quét lại và thực hiện các thao tác Move Up / Down bù trừ | Log WARN: `"Thứ tự cột chưa khớp chuẩn, đang căn chỉnh lại vị trí..."` |
| Số cột không bằng 14 | `ColumnConfigurationError` | Dừng tiến trình, không bấm Export | Log ERROR: `"Cấu hình cột thất bại: Danh sách cột có {count}/14 cột."` |

---

## 10. PHASE 7: BOM-EXPORT-RUN-01 – XUẤT DỮ LIỆU, GIÁM SÁT TẢI VỀ & KIỂM TRA OPENPYXL

### 10.1. User Story
```
As a Systems Automation Engineer,
I want the client to trigger the export in background mode, monitor and safely download the resulting Excel file, and run openpyxl integrity validation (14 columns, row count > 0, Japanese MS Gothic Unicode),
So that I am 100% guaranteed to receive a healthy, complete, and uncorrupted file ready for production comparison.
```

### 10.2. Acceptance Criteria (Given-When-Then)

- **AC-RUN-01 (Bật Run in Background & Kích Hoạt Export)**:
  - **Given** 14 cột đã được cấu hình hoàn chỉnh trong panel `Export To Excel`,
  - **When** client tích chọn checkbox `Run in Background` (`input[name="runInBackground"]`) và bấm nút `Export` (`button.sw-button` có text `"Export"`),
  - **Then** tác vụ xuất được gửi tới máy chủ Teamcenter 2412 mà không gây nghẽn giao diện UI.

- **AC-RUN-02 (Bắt Thông Báo Hoàn Tất & Kích Hoạt Tải File An Toàn - Active Download Trigger)**:
  - **Given** tác vụ xuất đang chạy nền trên máy chủ Active Workspace 2412,
  - **When** máy chủ hoàn thành xuất dữ liệu và hiển thị thông báo hoàn tất (banner/toast `div.noty_message` chứa nội dung xuất thành công hoặc chuông thông báo `button[command-id="Awp0ShowAlertWithBubble"]` xuất hiện popup),
  - **Then** client phát hiện thông báo, thực hiện hành động **CLICK BẮT BUỘC** vào liên kết tải tệp (`a.aw-widgets-cellListCellTitle`, `a.aw-jswidgets-tableCellLink`, hoặc nút download trong popup/Reports tab `button[command-id="Awp0DownloadReport"]`) để kích hoạt luồng tải xuống vật lý của trình duyệt; sau đó giám sát thư mục tải về cho đến khi tệp `.xlsx` hoàn tất với dung lượng > 5 KB và tệp tạm `.crdownload` biến mất hoàn toàn trong thời gian tối đa 300 giây. Client tuyệt đối không chỉ thụ động chờ đợi thư mục tải về mà không kích hoạt click vào thông báo.

- **AC-RUN-03 (Kiểm Tra Tính Toàn Vẹn Bằng openpyxl - 14 Cột & Dòng > 0)**:
  - **Given** tệp `.xlsx` đã tải về thư mục làm việc,
  - **When** module kiểm tra tự động nạp tệp bằng thư viện `openpyxl` (`load_workbook(filename, data_only=True)`),
  - **Then** tệp mở thành công, số dòng dữ liệu > 1 (ít nhất 1 header + dữ liệu BOM), và 14 ô từ A1 đến N1 khớp chính xác 100% với danh sách 14 tên cột chuẩn.

- **AC-RUN-04 (Bảo Toàn Bảng Mã Unicode Tiếng Nhật & Font MS Gothic)**:
  - **Given** tệp Excel chứa dữ liệu tên linh kiện (Cột 11 - Item Name) và thư mục (Cột 1 - Home) bằng tiếng Nhật (Kanji, Katakana, Hiragana),
  - **When** module kiểm tra đọc giá trị ô,
  - **Then** toàn bộ chuỗi ký tự tiếng Nhật hiển thị nguyên vẹn, không có ký tự rác `Mojibake` (ví dụ: `???`, ``), và font chữ được bảo toàn theo định dạng `MS Gothic` hoặc font Unicode hệ thống.

- **AC-RUN-05 (Chống File Rỗng & File Lỗi HTML - Fail-Closed)**:
  - **Given** máy chủ Teamcenter gặp sự cố và trả về file rỗng 0 byte hoặc trang lỗi HTML đóng gói nhầm đuôi `.xlsx`,
  - **When** module kiểm tra phát hiện dung lượng < 5 KB hoặc `openpyxl` báo lỗi `InvalidFileException`,
  - **Then** client ném ngoại lệ `BOMCorruptFileError`, xóa tệp hỏng và thông báo lỗi chi tiết.

### 10.3. Quy Trình Kiểm Tra Tính Toàn Vẹn Bằng openpyxl (Integrity Verification Code)
```python
from pathlib import Path
import openpyxl
from openpyxl.reader.excel import InvalidFileException

class BOMValidationError(Exception):
    """Ngoại lệ khi file Excel xuất ra không đạt tiêu chuẩn nghiệp vụ."""
    pass

class BOMCorruptFileError(Exception):
    """Ngoại lệ khi file Excel bị hỏng hoặc dung lượng rỗng."""
    pass

EXPECTED_HEADERS = [
    "Home", "Level", "Item Type", "Item Id", "Has Children", "Quantity",
    "1st Parts", "2nd BOM Flag", "Occurrence Effectivities",
    "Item Revision Project List", "Item Name", "Notice No", "Revision",
    "Item Rev Status"
]

def verify_exported_excel(file_path: Path) -> dict:
    """Xác thực toàn diện tính toàn vẹn của tệp Excel tải về từ TC2412."""
    if not file_path.exists():
        raise FileNotFoundError(f"Không tìm thấy tệp: {file_path}")
        
    file_size = file_path.stat().st_size
    if file_size < 5120:  # < 5 KB
        raise BOMCorruptFileError(f"Tệp quá nhỏ ({file_size} bytes), nghi ngờ file rỗng hoặc lỗi HTML!")
        
    try:
        wb = openpyxl.load_workbook(filename=str(file_path), data_only=True)
    except InvalidFileException as e:
        raise BOMCorruptFileError(f"Tệp không phải định dạng Excel hợp lệ: {e}")
        
    ws = wb.active
    if ws is None:
        raise BOMCorruptFileError("Workbook không có active worksheet!")
        
    # 1. Đọc header hàng 1 (14 cột từ A1 đến N1)
    actual_headers = [str(ws.cell(row=1, column=col_idx).value or "").strip() for col_idx in range(1, 15)]
    
    for idx, (expected, actual) in enumerate(zip(EXPECTED_HEADERS, actual_headers), start=1):
        if expected.lower() != actual.lower():
            raise BOMValidationError(
                f"Lỗi tiêu đề tại Cột {idx}: Kỳ vọng '{expected}', thực tế nhận được '{actual}'!"
            )
            
    # 2. Đếm số dòng dữ liệu
    max_row = ws.max_row
    if max_row < 2:
        raise BOMValidationError(f"Tệp Excel không có dòng dữ liệu nào (max_row={max_row})!")
        
    # 3. Kiểm tra bảo toàn bảng mã Unicode tiếng Nhật (Cột 11 - Item Name)
    japanese_detected = False
    for r in range(2, min(max_row + 1, 50)):  # Kiểm tra tối đa 50 dòng đầu
        cell_val = str(ws.cell(row=r, column=11).value or "")
        if any('\u3000' <= char <= '\u303f' or   # Punctuation
               '\u3040' <= char <= '\u309f' or   # Hiragana
               '\u30a0' <= char <= '\u30ff' or   # Katakana
               '\u4e00' <= char <= '\u9faf'      # Kanji
               for char in cell_val):
            japanese_detected = True
            break
            
    wb.close()
    
    return {
        "file_path": str(file_path),
        "file_size_bytes": file_size,
        "total_rows": max_row,
        "column_count": len(actual_headers),
        "headers": actual_headers,
        "unicode_japanese_verified": japanese_detected,
    }
```

### 10.4. Input / Output Schemas

#### Pydantic Schema (Python)
```python
from pydantic import BaseModel, Field
from pathlib import Path
from typing import List

class ExportRunInputSchema(BaseModel):
    output_dir: Path = Field(..., description="Thư mục lưu trữ file tải về")
    item_id: str = Field(..., description="Mã Item ID dùng đặt tên file")
    run_in_background: bool = Field(default=True, description="Bật chế độ chạy ngầm trên máy chủ TC2412")
    download_timeout_seconds: int = Field(default=300, ge=30, le=900, description="Timeout tối đa cho quá trình tải")

class ExportRunOutputSchema(BaseModel):
    downloaded_file_path: Path = Field(..., description="Đường dẫn tuyệt đối của tệp Excel đã tải về")
    file_size_bytes: int = Field(..., description="Dung lượng tệp tính bằng byte")
    total_rows: int = Field(..., ge=2, description="Tổng số dòng tính cả header")
    column_headers: List[str] = Field(..., description="Danh sách 14 header đã đối soát")
    unicode_japanese_verified: bool = Field(..., description="Đã xác thực bảo toàn Unicode tiếng Nhật")
    duration_seconds: float = Field(..., description="Thời gian thực thi xuất và kiểm tra")
```

### 10.5. Verified DOM Locators Table (TC2412 Standard)

| UI Element | Tag | CSS Selector | XPath Expression | Identifying Attributes | Fallback Selector | Strategy / Notes |
|:---|:---|:---|:---|:---|:---|:---|
| **Run in Background Checkbox** | `input` | `input[type="checkbox"][name*="Background"]` | `//input[@type='checkbox' and (contains(@name,'Background') or contains(@data-locator,'Background'))]` | `name="runInBackground"` | `input[data-locator*="Background"]` | Tích chọn trước khi xuất |
| **Export Submit Button** | `button` | `form.sw-command-panel button.sw-button` | `//form[contains(@class,'sw-command-panel')]//button[contains(@class,'sw-button') and contains(.,'Export')]` | text: `"Export"`, class `sw-button` | `//button[text()='Export']` | Bấm để bắt đầu xuất |
| **Export Completion Toast** | `div` | `div.noty_message, div.aw-layout-popup` | `//div[contains(@class,'noty_message') or contains(@class,'aw-layout-popup')]` | `class="noty_message"` | `div.aw-layout-popup` | Toast popup xuất hiện khi job nền hoàn thành |
| **Notification Download Link** | `a` | `div.noty_message a, a.aw-widgets-cellListCellTitle` | `//div[contains(@class,'noty_message')]//a \| //a[contains(@class,'aw-widgets-cellListCellTitle')]` | `data-locator="aw-clickable-title"` | `a.aw-jswidgets-tableCellLink` | **CLICK BẮT BUỘC** để kích hoạt trình duyệt kéo tệp vật lý về đĩa |
| **Alerts Notification Bell** | `button` | `button[command-id="Awp0ShowAlertWithBubble"]` | `//button[@command-id='Awp0ShowAlertWithBubble']` | `command-id="Awp0ShowAlertWithBubble"` | `button[aria-label="Alerts"]` | Báo hoàn thành job ngầm (dự phòng) |
| **Reports Tab Link** | `a` | `a[data-locator="tab-tc_xrt_Rb0Reports"]` | `//a[@data-locator='tab-tc_xrt_Rb0Reports']` | `data-locator="tab-tc_xrt_Rb0Reports"` | `button[command-id="Awp0GoReports"]` | Nơi lưu trữ báo cáo xuất (dự phòng) |
| **Report Download Button** | `button` | `button[command-id="Awp0DownloadReport"]` | `//button[@command-id='Awp0DownloadReport']` | `command-id="Awp0DownloadReport"` | `button.aw-commandId-Awp0DownloadReport` | Nút tải tệp trong danh sách Reports nếu toast biến mất |

### 10.6. Error Handling & Recovery Matrix

| Tình huống lỗi | Mã lỗi ngoại lệ | Hành động khôi phục | Nhật ký & Thông báo người dùng |
|:---|:---|:---|:---|
| Quá hạn tải tệp (> 300s) | `ExportDownloadTimeoutError` | Kiểm tra toast notification hoặc vào tab Reports để click trực tiếp `Awp0DownloadReport` | Log ERROR: `"Hết thời gian chờ tải tệp Excel (300s). Kiểm tra toast hoặc tab Reports."` |
| Tệp bị kẹt đuôi `.crdownload` | `IncompleteDownloadError` | Chờ thêm tối đa 30s trước khi hủy tiến trình | Log WARN: `"Tệp đang trong quá trình ghi dữ liệu (.crdownload), đang chờ hoàn tất..."` |
| Sai tiêu đề cột trong Excel | `BOMValidationError` | Dừng quy trình, không chuyển sang Lõi so sánh | Log ERROR: `"Dữ liệu tệp Excel không khớp định dạng 14 cột tiêu chuẩn!"` |
| Tệp rỗng hoặc lỗi mã hóa | `BOMCorruptFileError` | Xóa tệp hỏng, báo lỗi để thực hiện lại | Log ERROR: `"Tệp Excel tải về bị lỗi hoặc rỗng. Đã xóa tệp tạm hỏng."` |
| Tệp bị kẹt đuôi `.crdownload` | `IncompleteDownloadError` | Chờ thêm tối đa 30s trước khi hủy tiến trình | Log WARN: `"Tệp đang trong quá trình ghi dữ liệu (.crdownload), đang chờ hoàn tất..."` |
| Sai tiêu đề cột trong Excel | `BOMValidationError` | Dừng quy trình, không chuyển sang Lõi so sánh | Log ERROR: `"Dữ liệu tệp Excel không khớp định dạng 14 cột tiêu chuẩn!"` |
| Tệp rỗng hoặc lỗi mã hóa | `BOMCorruptFileError` | Xóa tệp hỏng, báo lỗi để thực hiện lại | Log ERROR: `"Tệp Excel tải về bị lỗi hoặc rỗng. Đã xóa tệp tạm hỏng."` |

---

## 11. PHASE 8: REP-01 (R8) – TRÌNH BÁO CÁO TIẾN ĐỘ ĐỊNH KỲ (2 PHÚT/CHU KỲ)

### 11.1. User Story
```
As an Operations Monitor and Parent Orchestration Agent,
I want the automation execution to report structured progress updates on a 2-minute periodic cycle and at key phase boundaries in the format 'Phase [X/8] - [Name] - [%] - [Elapsed]',
So that long-running operations (> 10 minutes) provide continuous visibility and prevent premature timeout kills.
```

### 11.2. Acceptance Criteria (Given-When-Then)

- **AC-REP-01 (Định Dạng Báo Cáo Chuẩn Hóa)**:
  - **Given** quy trình tự động hóa đang chạy,
  - **When** một phase hoàn thành hoặc chu kỳ 2 phút (120 giây) chạm mốc,
  - **Then** thông báo tiến độ được phát ra tuân thủ chính xác định dạng quy định tại `PROJECT.md:88-91`:
    `Phase [X/8] - [Tên phân hệ] - [Tỷ lệ hoàn thành %] - [Thời gian: mm:ss elapsed]`

- **AC-REP-02 (Chu Kỳ Báo Cáo Không Chặn 120 Giây - Non-blocking Heartbeat)**:
  - **Given** một phase xử lý lâu (như Phase 3 bung cây Level 7 hoặc Phase 7 chạy background export mất nhiều phút),
  - **When** đồng hồ đạt mốc 120 giây mà phase chưa đổi,
  - **Then** một thông báo tiến độ heartbeat định kỳ được gửi ra với thời gian trôi qua cập nhật mà không làm gián đoạn luồng WebDriver đang chạy.

- **AC-REP-03 (Phân Bổ 8 Giai Đoạn Tuyến Tính)**:
  - **Given** 7 phân hệ nghiệp vụ và giai đoạn hoàn tất toàn trình,
  - **When** các phase lần lượt được thực thi,
  - **Then** tỷ lệ phần trăm tương ứng với từng pha tuân theo bảng tỷ trọng chuẩn:

| Giai đoạn | Mã phân hệ | Tên phân hệ (Phase Name) | % Tiến độ | Mô tả nghiệp vụ |
|:---:|:---:|:---|:---:|:---|
| **[1/8]** | `AUTH-01` | Authentication & Keyring Persistence | 12.5% | Đăng nhập ngầm TC2412, xác thực session cookie |
| **[2/8]** | `BOM-SEARCH-01` | Search & Navigation | 25.0% | Tìm Item ID, đối soát Rev, vào tab Content |
| **[3/8]** | `BOM-EXPAND-01` | Deep Tree Expansion (Level 7) | 37.5% | Chọn root, bung cây phân cấp đến Level 7 |
| **[4/8]** | `BOM-SELECT-01` | Select All Rows & Open Menu | 50.0% | Chọn tất cả các dòng, mở menu Excel Report |
| **[5/8]** | `BOM-EXPORT-OPEN-01` | Safe Export Trigger (Anti-Import) | 62.5% | Mở panel Export, kích hoạt chốt chặn an toàn |
| **[6/8]** | `BOM-EXPORT-CONFIG-01` | 14-Column Configuration | 75.0% | Nguồn Item, xóa cột thừa, xếp 14 cột chuẩn |
| **[7/8]** | `BOM-EXPORT-RUN-01` | Run Export & File Verification | 87.5% | Xuất nền, tải file, kiểm tra openpyxl & Unicode |
| **[8/8]** | `COMPLETE` | Pipeline Finalized | 100.0% | Hoàn tất toàn trình, bàn giao dữ liệu cho Core |

### 11.3. Cấu Trúc Khối Báo Cáo Tiến Độ (Implementation Contract)
```python
import time
import threading
from typing import Callable, Optional

class ProgressReporter:
    """Bộ báo cáo tiến độ định kỳ chu kỳ 2 phút (120 giây) theo chuẩn [X/8].
    
    Tuân thủ hợp đồng tại PROJECT.md:88-91 và PROJECT.md:82:
    - Cung cấp phương thức update(phase_index, total_phases, phase_name, percent) -> str
    - Cung cấp cơ chế callback linh hoạt hỗ trợ 5 tham số:
      (phase_index: int, total_phases: int, phase_name: str, percent: float, elapsed: float) -> None
      và callback dạng chuỗi văn bản cho logging/UI.
    """
    
    PHASE_NAMES = {
        1: "Authentication & Keyring Persistence",
        2: "Search & Navigation",
        3: "Deep Tree Expansion (Level 7)",
        4: "Select All Rows & Open Menu",
        5: "Safe Export Trigger (Anti-Import)",
        6: "14-Column Configuration",
        7: "Run Export & File Verification",
        8: "Pipeline Finalized"
    }

    def __init__(
        self,
        callback: Optional[Callable[[int, int, str, float, float], None]] = None,
        string_callback: Optional[Callable[[str], None]] = None,
        interval_seconds: float = 120.0
    ):
        self.callback = callback
        self.string_callback = string_callback or print
        self.interval_seconds = interval_seconds
        self.start_time = time.time()
        self.current_phase = 1
        self.total_phases = 8
        self.running = False
        self._thread: Optional[threading.Thread] = None

    def start(self):
        """Khởi chạy luồng timer phát heartbeat định kỳ 2 phút."""
        self.running = True
        self._thread = threading.Thread(target=self._run_timer, daemon=True)
        self._thread.start()

    def stop(self):
        """Dừng luồng timer."""
        self.running = False

    def update(self, phase_index: int, total_phases: int, phase_name: str, percent: float) -> str:
        """Cập nhật tiến độ, kích hoạt callbacks và trả về chuỗi định dạng chuẩn.
        
        Tuân thủ PROJECT.md:88-91:
        Returns formatted string: 'Phase [X/8] - [Name] - [XX.X%] - [MM:SS elapsed]'
        """
        self.current_phase = min(max(phase_index, 1), total_phases)
        self.total_phases = total_phases
        elapsed = time.time() - self.start_time
        elapsed_int = int(elapsed)
        mm, ss = divmod(elapsed_int, 60)
        
        # Format string chuẩn theo PROJECT.md:90
        formatted_str = f"Phase [{self.current_phase}/{self.total_phases}] - {phase_name} - [{percent:.1f}%] - [{mm:02d}:{ss:02d} elapsed]"
        
        # Bắn callback 5 tham số nếu được đăng ký (tương thích TC2412AutomationClient và adapters.py)
        if self.callback:
            try:
                self.callback(self.current_phase, self.total_phases, phase_name, percent, elapsed)
            except Exception:
                pass
                
        # Bắn callback chuỗi (cho stdout / logging)
        if self.string_callback:
            try:
                self.string_callback(formatted_str)
            except Exception:
                pass
                
        return formatted_str

    def update_phase(self, phase_index: int) -> str:
        """Hàm tiện ích chuyển đổi theo số phase mặc định [1..8]."""
        idx = min(max(phase_index, 1), 8)
        name = self.PHASE_NAMES.get(idx, "Unknown")
        pct = (idx / 8.0) * 100.0
        return self.update(idx, 8, name, pct)

    def _emit_heartbeat(self):
        """Phát heartbeat chu kỳ 120s với thời gian trôi qua mới nhất."""
        idx = self.current_phase
        name = self.PHASE_NAMES.get(idx, "Unknown")
        pct = (idx / float(self.total_phases)) * 100.0
        self.update(idx, self.total_phases, name, pct)

    def _run_timer(self):
        while self.running:
            time.sleep(self.interval_seconds)
            if self.running:
                self._emit_heartbeat()
```

---

## 12. MA TRẬN TRUY XUẤT NGUỒN GỐC & KIỂM THỬ TOÀN DIỆN (TRACEABILITY MATRIX TIER 1 - TIER 5)

Ma trận dưới đây liên kết toàn bộ 9 mã yêu cầu nghiệp vụ từ các chỉ thị người dùng (`ORIGINAL_REQUEST.md`), tài liệu slide PPTX, và kiến trúc hệ thống (`PROJECT.md`) xuyên suốt 5 tầng kiểm thử tự động (Tiers 1-5):

### 12.1. Ma Trận Truy Xuất Tổng Thể (Master 5-Tier Traceability Matrix)

| Mã Yêu Cầu | Nguồn Chỉ Thị / Slide | Phân Hệ Đặc Tả | File Triển Khai | Tier 1: Feature Isolation | Tier 2: Boundary Conditions | Tier 3: Cross-Phase Integration | Tier 4: Real-World Scenarios | Tier 5: Adversarial Hardening |
|:---|:---|:---|:---|:---|:---|:---|:---|:---|
| **REQ-AUTH-01** | Slide 1, R2, 04:35:40Z | `AUTH-01` (Phase 1) | `src/security/credentials.py`, `src/automation/tc2412/client.py` | `test_auth.py` (Mock login, form fill, DPAPI CRUD) | Thử lại 3 lần timeout, tài khoản bị khóa, rỗng mật khẩu | Phiên đăng nhập hợp lệ bàn giao session cookie cho Phase 2 | Tự động đăng nhập với tài khoản nhà máy `vn_pe03` | Chống rò rỉ plaintext trong memory/log/dump, chống trộm key DPAPI |
| **REQ-SEARCH-01** | Slide 2, R3, 04:35:40Z | `BOM-SEARCH-01` (Phase 2) | `client.py`, `selectors.py` | `test_search.py` (Input part number, open Content tab) | Mã không tồn tại (`ItemNotFoundError`), timeout 30s, đa Rev | Nhận session từ Phase 1, nạp tab Content chuyển giao cho Phase 3 | Tìm mã BOM thực tế (e.g. `110C103NL0`, `110C0Z3LV1`) | Chống lỗi search lag, phục hồi qua hash route trực tiếp |
| **REQ-EXPAND-01** | Slide 3, R4, 04:35:40Z | `BOM-EXPAND-01` (Phase 3) | `client.py`, `session.py` | `test_expand.py` (Click root, set level 7, verify progressBar) | Giới hạn Level 7, cây BOM đơn cấp (0 linh kiện con), timeout 180s | Cây bung xong sẵn sàng dữ liệu cho Phase 4 Select All | Bung cây lớn (> 2.500 dòng), bảo toàn thứ tự quan hệ cha con | Kháng bẫy Virtual DOM (~30-50 nodes viewport), dynamic wait ổn định |
| **REQ-SELECT-01** | Slide 4, R5, 04:35:40Z | `BOM-SELECT-01` (Phase 4) | `client.py` | `test_select.py` (Select All, mở menu Excel Report) | Bảng 0 dòng được chọn (`NoRowsSelectedError`), timeout 30s | Duy trì 100% selection khi mở menu Excel sang Phase 5 | Đối soát chọn toàn bộ cây phân cấp lắp ráp phức tạp | Đọc chỉ số `aw-splm-tableSelectionCount`, chống false negative DOM ảo |
| **REQ-EXPORT-01** | Slide 5, R5, 04:35:40Z | `BOM-EXPORT-OPEN-01` (Phase 5) | `client.py` | `test_export_open.py` (Quét menu, mở panel `Export To Excel`) | Menu thiếu mục Export, nút bị disable, timeout 15s | Giữ nguyên selection của Phase 4 khi panel cấu hình mở ra | Mở đúng panel flyout `#ui-id-3` trên màn hình chuẩn 1080p | **Fail-Closed Anti-Import Guard**: Chặn 100% click vào `Import Changes` |
| **REQ-CONFIG-01** | Slide 6, R6, 04:35:40Z | `BOM-EXPORT-CONFIG-01` (Phase 6) | `client.py` | `test_config.py` (Nguồn Item, xếp 14 cột theo thứ tự chuẩn) | Thiếu cột (13/14), thừa cột (15/14), sai thứ tự cột, thiếu property | Cấu hình hoàn tất chuyển sang Phase 7 thực thi lệnh xuất | Lọc qua search box `aw-uiwidgets-searchBox` giữa hàng trăm thuộc tính | Chống trùng cột, chống lưu cấu hình hỏng, chặn submit sai format |
| **REQ-RUN-01** | Slide 7, R7, 04:35:40Z | `BOM-EXPORT-RUN-01` (Phase 7) | `client.py` | `test_run.py` (Run in Background, download poll, openpyxl) | File < 5KB, file rỗng 0-byte, kẹt `.crdownload`, timeout 300s | Xuất file và chuyển giao đường dẫn `Path` cho Core Adapter | Kiểm tra bảo toàn font tiếng Nhật `MS Gothic`, Kanji/Katakana | **Active Download Trigger**: Bắt toast `noty_message` & click tải file |
| **REQ-REP-01** | R8, 04:35:40Z | `REP-01 (R8)` (Phase 8) | `src/automation/tc2412/progress_reporter.py` | `test_progress_reporter.py` (Định dạng `[X/8]`, callback 5 args) | Sai lệch thời gian trôi qua, phân bổ phần trăm [1..8], dừng sạch | Cập nhật chuyển phase tức thì xuyên suốt Phase 1 đến Phase 7 | Heartbeat định kỳ 120s phát liên tục trong thời gian bung cây/tải tệp | Threading an toàn, không block main WebDriver, bắt ngoại lệ callback |
| **REQ-ADAPT-01** | R5, PROJECT.md | Core Adapter Bridge | `src/core/adapters.py`, `src/core/model_pruner.py` | `test_adapter.py` (Cắm rút PLMProvider, hỗ trợ `part_rev`) | Xử lý `part_rev=None` (chọn latest) vs chỉ định cụ thể (e.g. `'01'`) | Tích hợp toàn trình `TC2412AutomationClient` -> `TeamcenterSeleniumAdapter` | Đối soát dữ liệu với `BOM_*.xlsm`, chạy hàm `prune_excel_file` | Kháng lỗi sai định dạng tệp, chuyển giao an toàn cho lõi so sánh BOM |

---

### 12.2. Chi Tiết Các Tầng Kiểm Thử (Detailed Test Tier Catalog)

#### Tier 1: Feature Isolation (Độc Lập & Cô Lập Phân Hệ)
- **Mục tiêu**: Kiểm thử tính đúng đắn của từng phase riêng lẻ bằng Selenium WebDriver mock hoặc headless browser điều khiển DOM fixtures cục bộ.
- **Thư mục mã kiểm thử**: `tests/tier1_features/`
  - `test_auth.py`: Kiểm thử luồng điền form đăng nhập, tương tác nút bấm, lưu trữ credential.
  - `test_search.py`: Kiểm thử nhập liệu tìm kiếm, chọn kết quả, chuyển tab Content.
  - `test_expand.py`: Kiểm thử click root node, mở menu Expand Below, điền cấp độ 7.
  - `test_select.py`: Kiểm thử click Select All, kiểm tra thuộc tính phần tử được chọn.
  - `test_export_open.py`: Kiểm thử quét các phần tử menu, nhận diện chính xác Export to Excel.
  - `test_config.py`: Kiểm thử chọn Property Source, thêm/xóa cột, sắp xếp thứ tự 1-14.
  - `test_run.py`: Kiểm thử tick Run in Background, bấm Export, mock tải tệp thành công.
  - `test_progress_reporter.py`: Kiểm thử phương thức `update(...) -> str` và phát chuỗi định dạng chuẩn `Phase [X/8] - [Name] - [XX.X%] - [MM:SS elapsed]`.

#### Tier 2: Boundary Conditions (Điều Kiện Biên, Ngưỡng Timeout & Xử Lý Lỗi)
- **Mục tiêu**: Đảm bảo hệ thống không bị crash hoặc treo khi gặp dữ liệu bất thường hoặc quá hạn thời gian.
- **Thư mục mã kiểm thử**: `tests/tier2_boundaries/`
  - `test_auth_retry_boundary.py`: Kiểm thử thử lại tối đa 3 lần khi mạng chập chờn; dừng ngay khi gặp lỗi sai mật khẩu tránh khóa tài khoản AD.
  - `test_search_not_found.py`: Kiểm thử mã Item không tồn tại trả về đúng `ItemNotFoundError` trong 30s.
  - `test_expand_depth_limits.py`: Kiểm thử độ sâu biên (Level 1, Level 7, Level > 7) và cây BOM đơn cấp 0 con.
  - `test_select_zero_rows.py`: Kiểm thử phát hiện không có dòng nào được chọn ném `NoRowsSelectedError`.
  - `test_config_column_boundaries.py`: Kiểm thử khi danh sách cột thiếu (13 cột) hoặc thừa (15 cột) ném `ColumnConfigurationError`.
  - `test_file_integrity_boundaries.py`: Kiểm thử file rỗng 0 byte, file lỗi HTML, file < 5KB ném `BOMCorruptFileError`.

#### Tier 3: Combinations & Integration Pipeline (Tích Hợp Chuỗi Toàn Trình)
- **Mục tiêu**: Kiểm thử luồng dữ liệu liên tục xuyên suốt từ Phase 1 đến Phase 7, kết nối qua `TC2412AutomationClient.download_bom_full`.
- **Thư mục mã kiểm thử**: `tests/tier3_combinations/`
  - `test_full_pipeline_flow.py`: Kiểm thử chạy tuần tự 7 phase trong một session WebDriver duy nhất, xác nhận cookie phiên được duy trì xuyên suốt.
  - `test_adapter_integration.py`: Kiểm thử cắm rút `TeamcenterSeleniumAdapter` vào `PLMProvider`, gọi `download_bom_full` và trả về `Path` hợp lệ cho Core.
  - `test_reporter_pipeline_binding.py`: Kiểm thử `ProgressReporter` nhận callback 5 tham số chính xác tại mỗi ranh giới phase trong quá trình thực thi pipeline.

#### Tier 4: Real-World Scenarios (Dữ Liệu Thực Tế & Toàn Vẹn Doanh Nghiệp)
- **Mục tiêu**: Xác thực chất lượng dữ liệu đầu ra trên các file BOM thực tế của nhà máy.
- **Thư mục mã kiểm thử**: `tests/tier4_real_world/`
  - `test_canonical_14_columns.py`: Xác thực file Excel đầu ra có đúng 14 cột theo thứ tự bất biến từ A đến N.
  - `test_japanese_unicode_preservation.py`: Xác thực các ký tự tiếng Nhật (Kanji, Katakana, Hiragana) tại Cột 11 (Item Name) hiển thị toàn vẹn, không bị lỗi font hay Mojibake.
  - `test_large_bom_handling.py`: Kiểm thử tải cụm máy lớn (> 2.500 dòng) với chế độ Run in Background hoàn tất ổn định.

#### Tier 5: Adversarial Hardening (Phòng Vệ Chủ Động & An Toàn Nghiêm Ngặt)
- **Mục tiêu**: Thử thách hệ thống trước các tình huống tấn công, bất thường giao diện hoặc lỗi hạ tầng nguy hiểm.
- **Thư mục mã kiểm thử**: `tests/tier5_adversarial/`
  - `test_anti_import_fail_closed.py`: Cố tình tiêm các phần tử menu chứa chuỗi "Import Changes", đổi chỗ vị trí với Export để khẳng định chốt chặn `ImportChangesForbiddenError` luôn kích hoạt 100% và ngắt browser ngay lập tức.
  - `test_credential_security_leak.py`: Quét bộ nhớ, log files, stdout, crash traces để đảm bảo không xuất hiện plaintext password.
  - `test_virtual_dom_scrolling_resilience.py`: Giả lập cây BOM 5.000 dòng trong viewport 36 dòng, khẳng định hệ thống đếm đúng số dòng qua `aw-splm-tableSelectionCount` mà không gây crash trình duyệt.
  - `test_background_download_trigger_click.py`: Giả lập thông báo toast `div.noty_message` và kiểm tra client tự động gửi lệnh click trigger để kéo file về máy, không bị treo chờ thụ động.

---

## 13. CHUẨN HÓA DỮ LIỆU CANONICAL & CẦU NỐI SHEET PLM BẢO TOÀN CÔNG THỨC EXCEL

### 13.1. Phân Tích Tử Huyệt Công Thức Excel Trong File BOM Tổng Hợp
Trong hệ thống so sánh BOM tự động (`pm_sosanhbom`), tệp kết quả tổng hợp (ví dụ: `BOM_110C0Z3LV1.xlsm` kế thừa từ `form_ssbom.xlsm`) chứa mạng lưới công thức Excel liên kết chéo cực kỳ nhạy cảm giữa các sheet `Tongket`, `CTTT` và `PLM`:
1. **Sheet `Tongket!C5` & `CTTT!A1`**:
   - Chứa công thức trực tiếp `=PLM!C2` để lấy Mã máy (Top-level Model Part Number).
   - **Tử huyệt**: Bắt buộc ô `C2` trên sheet `PLM` phải chứa Mã linh kiện của cụm gốc (Item ID). Nếu Mã linh kiện bị dịch chuyển sang cột khác (ví dụ Cột A hay D), toàn bộ báo cáo tổng kết và chỉ thị thao tác sẽ nhận giá trị rỗng hoặc sai lệch hoàn toàn.
2. **Sheet `CTTT` (Chỉ Thị Thao Tác)**:
   - Chứa công thức tra cứu `=VLOOKUP(C3, PLM!C:L, 10, 0)`:
     - Điểm neo tra cứu bắt đầu từ **Cột C** (`Item Id`).
     - Chỉ số cột trả về là **10** (tức là Cột thứ 10 tính từ Cột C: C->D->E->F->G->H->I->J->K->**L**).
     - **Tử huyệt**: **Cột L** trên sheet `PLM` bắt buộc phải là **`Revision`**! Nếu Revision nằm sai cột, hàm VLOOKUP trả về `#N/A` hoặc sai phiên bản linh kiện.
   - Chứa công thức tra cứu `=VLOOKUP(C3, PLM!T:U, 2, 0)`:
     - Tra cứu từ bảng Pivot Table / tổng hợp dữ liệu tại vùng **Cột T:U**.
3. **Sheet `PLM`**:
   - Cột R chứa công thức `=IF(C2="","",C2)` (PART CODE).
   - Cột S chứa công thức `=IF(E2="","",E2)` (Q.TY - Số lượng).
   - **Tử huyệt**: Cột C bắt buộc là `Item Id` và Cột E bắt buộc là `Quantity`!

### 13.2. Cơ Chế Canonical Normalizer (Đọc 14 Cột Chuẩn & 24 Cột TC2412 Mới)
Trên thực tế sản xuất tại nhà máy, khi người dùng hoặc hệ thống xuất dữ liệu từ Teamcenter 2412 mới (ví dụ tệp `T10C423NL0 Mới.xlsm`), Teamcenter có thể xuất ra file gồm **24 cột** với các tên tiêu đề khác biệt. Bộ chuẩn hóa `TC2412CanonicalNormalizer` có nhiệm vụ tự động nhận diện và ánh xạ đồng nhất về mô hình Canonical Model:

| Thuộc tính Canonical | Vị trí trong File 14 Cột Chuẩn (Mục 3) | Vị trí & Tên Header trong File 24 Cột TC2412 Mới | Ý nghĩa nghiệp vụ |
|:---|:---:|:---|:---|
| **level** | Cột B (`Level`) | **Col B** (`Level`) | Tầng phân cấp BOM (0, 1, 2...) |
| **item_type** | Cột C (`Item Type`) | **Col C** (`Item Type`) | Loại đối tượng chi tiết |
| **item_id** | Cột D (`Item Id`) | **Col D (`Name`)** | **Mã linh kiện (Part Number)** |
| **has_children** | Cột E (`Has Children`) | Col E (`Has Children`) | Cờ có chi tiết con (`true`/`false`) |
| **quantity** | Cột F (`Quantity`) | Col F (`Quantity`) | Số lượng chi tiết cấu thành cụm |
| **item_name** | Cột K (`Item Name`) | **Col H (`Parts Text`)** | **Tên linh kiện tiếng Nhật (Kanji/MS Gothic)** |
| **revision** | Cột M (`Revision`) | **Col J (`Revision`)** | **Ký hiệu phiên bản chi tiết (Rev)** |
| **item_rev_status** | Cột N (`Item Rev Status`) | **Col K (`Release Status`)** | **Trạng thái phát hành chi tiết** |

```python
class TC2412CanonicalNormalizer:
    """Bộ chuẩn hóa dữ liệu BOM đầu vào, hỗ trợ cả tệp 14 cột chuẩn và 24 cột TC2412 mới."""
    
    @classmethod
    def normalize_row(cls, raw_row: dict) -> dict:
        """Tự động ánh xạ header tệp xuất thành các trường dữ liệu chuẩn hóa (Canonical Dict)."""
        # Nếu tệp xuất là 24 cột TC2412 mới
        if "Name" in raw_row and "Parts Text" in raw_row:
            return {
                "level": raw_row.get("Level", 0),
                "item_type": raw_row.get("Item Type", ""),
                "item_id": raw_row.get("Name", ""),  # Col D Name -> item_id
                "has_children": raw_row.get("Has Children", False),
                "quantity": raw_row.get("Quantity", 1),
                "item_name": raw_row.get("Parts Text", ""),  # Col H Parts Text -> item_name
                "revision": raw_row.get("Revision", ""),  # Col J Revision -> revision
                "item_rev_status": raw_row.get("Release Status", ""),  # Col K Release Status
                "occurrence_effectivity": raw_row.get("Occurrence Effectivities", ""),
                "project_list": raw_row.get("Item Revision Project List", ""),
            }
        # Nếu tệp xuất là 14 cột chuẩn
        return {
            "level": raw_row.get("Level", 0),
            "item_type": raw_row.get("Item Type", ""),
            "item_id": raw_row.get("Item Id", ""),
            "has_children": raw_row.get("Has Children", False),
            "quantity": raw_row.get("Quantity", 1),
            "item_name": raw_row.get("Item Name", ""),
            "revision": raw_row.get("Revision", ""),
            "item_rev_status": raw_row.get("Item Rev Status", ""),
            "occurrence_effectivity": raw_row.get("Occurrence Effectivities", ""),
            "project_list": raw_row.get("Item Revision Project List", ""),
        }
```

### 13.3. Quy Tắc Ánh Xạ Sheet 'PLM' Trong File BOM Tổng Hợp (Preserving Excel Formulas)
Khi ghi dữ liệu đã tải về vào Sheet `PLM` của tệp `BOM_*.xlsm`, module Adapter và Writer **BẮT BUỘC** ghi dữ liệu vào các cột chính xác theo bảng dưới đây nhằm bảo toàn 100% tính toàn vẹn của các công thức Excel phụ thuộc:

| Vị trí Cột Sheet 'PLM' | Tên Trường Dữ Liệu | Nguồn Trích Xuất Canonical | Công Thức Excel Phụ Thuộc Trong Workbook | Mức Độ Nghiêm Ngặt |
|:---:|:---|:---|:---|:---|
| **Cột A** | `Level` | `level` (Col B TC2412) | Duyệt cây phân cấp Lõi Python & bộ lọc hiển thị VBA | BẮT BUỘC |
| **Cột B** | `Item Type` | `item_type` (Col C TC2412) | Phân loại loại chi tiết | BẮT BUỘC |
| **Cột C** | **`Item Id`** | **`item_id` (Col D 'Name' TC2412)** | **`=PLM!C2` (tại `Tongket!C5` & `CTTT!A1`), điểm bắt đầu `VLOOKUP(C3, PLM!C:L, ...)`** | **TỬ HUYỆT (CRITICAL)** |
| **Cột D** | `Has Children` | `has_children` | Kiểm tra cụm chi tiết có linh kiện con | BẮT BUỘC |
| **Cột E** | **`Quantity`** | **`quantity`** | **`=IF(E2="","",E2)` (tại `PLM!S2`)** | **TỬ HUYỆT (CRITICAL)** |
| **Cột F..I**| Thuộc tính bổ trợ | `1st Parts`, `2nd BOM Flag`, `Effectivities`, `Project List` | Lọc Model máy Virgo/Libra/Iris | BẮT BUỘC |
| **Cột J** | **`Item Name`** | **`item_name` (Col H 'Parts Text' TC2412)** | Tên linh kiện tiếng Nhật font `MS Gothic` | BẮT BUỘC |
| **Cột K** | `Notice No` | `ec_notice_no` | Số thông báo kỹ thuật ECN | BẮT BUỘC |
| **Cột L** | **`Revision`** | **`revision` (Col J 'Revision' TC2412)** | **`=VLOOKUP(C3, PLM!C:L, 10, 0)` (lấy Revision tại Cột L - Cột thứ 10)** | **TỬ HUYỆT (CRITICAL)** |
| **Cột M** | `Item Rev Status`| `item_rev_status` (Col K 'Release Status' TC2412) | Trạng thái phát hành linh kiện | BẮT BUỘC |
| **Cột R** | **PART CODE** | Công thức tự động `=IF(C2="","",C2)` | Cột phụ trợ phục vụ đối soát CTTT | BẮT BUỘC |
| **Cột S** | **Q.TY** | Công thức tự động `=IF(E2="","",E2)` | Cột phụ trợ số lượng | BẮT BUỘC |
| **Cột T:U**| **Pivot Table**| Dữ liệu tổng hợp linh kiện | **`=VLOOKUP(C3, PLM!T:U, 2, 0)` (tại sheet CTTT)** | **TỬ HUYỆT (CRITICAL)** |

### 13.4. Cam Kết Bất Biến (Invariant Integrity Mandate)
1. **Không Dịch Chuyển Cột C**: Ô `C2` luôn luôn là nơi lưu trữ Mã máy của cụm gốc.
2. **Không Dịch Chuyển Cột L**: Revision luôn luôn nằm tại Cột L để chỉ số cột 10 trong `VLOOKUP(..., PLM!C:L, 10, 0)` không bao giờ bị lệch.
3. **Bảo Tồn Vùng T:U**: Khu vực cột T và U không bao giờ bị ghi đè bởi dữ liệu text thô mà được dành riêng cho Pivot Table / công thức tra cứu.
4. **Không Phá Vỡ VBA / Macros Cũ**: Đảm bảo tệp sau khi cập nhật Sheet `PLM` có thể mở lại bình thường trong Microsoft Excel mà không phát sinh cảnh báo `"Circular Reference"` hoặc `#REF!`.

---
*Tài liệu này là đặc tả kỹ thuật tối cao cho Milestone M1 của dự án `pm_sosanhbom`. Mọi thay đổi về cấu trúc giao diện hoặc hành vi của các phân hệ M2..M5 phải được đối chiếu và cập nhật đồng bộ với văn bản này.*
