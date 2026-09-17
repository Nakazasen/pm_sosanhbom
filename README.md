# Hệ Thống So Sánh BOM Tự Động (SSBOM Manager v1.0.0)

Phần mềm tự động hóa đối soát danh mục linh kiện (BOM - Bill of Materials) đa cấp độ, hiện đại hóa toàn diện từ hệ thống cũ chạy bằng Excel VBA macros (`form_ssbom.xlsm`, `tonghop_*.xlsm`, `Hamtimlinhkienthuoc_UNIT_naotren_BOM.xlsx`) và VBScript (`tudongdangnhapR3.vbs`).

Hệ thống được thiết kế bền vững, mở rộng lâu dài theo kiến trúc **Adapter Pattern**, tích hợp cổng thông tin kỹ thuật **Siemens Teamcenter Active Workspace (TC14)** và hệ thống quản trị nguồn lực doanh nghiệp **SAP R3 ERP**.

---

## 🌟 Tính Năng Nổi Bật

1. **Siemens Teamcenter TC14 Web Automation**:
   - Tự động hóa trình duyệt qua Selenium WebDriver (Microsoft Edge headless) kết nối cổng Web `http://tcmp3gwb:3000/` (tài khoản `vn_pe03/vn_pe03`).
   - Tự động tìm kiếm mã sản phẩm, xử lý bung phân cấp BOM hoặc tải BOM Full 14 cột qua Chrome DevTools Protocol (CDP).
   - Cơ chế tự phát hiện crash và khôi phục phiên thông minh (`is_session_alive`).

2. **SAP R3 CS12 Automation**:
   - Tự động hóa qua giao diện `win32com.client` kết nối SAP GUI Scripting 770 (`P1J(ERP60-AWS)-VN`, Plant 2200, Usage pp01, Alt 01).
   - Tự động xử lý hộp thoại đăng nhập đa phiên `radMULTI_LOGON_OPT2` và xuất báo cáo đa tầng CS12.

3. **Core Engine Xử Lý Thuật Toán Độc Lập**:
   - **Lọc ngày hiệu lực**: Bộ lọc 2 lượt (Dual-Pass Date Filter) xử lý chuỗi "to", "UP" so với ngày đối soát, an toàn tuyệt đối trước đồ thị vòng lặp chu trình (cycle-safe).
   - **Bóc tách Model máy**: Cắt tỉa cây BOM theo mã Model (Virgo, Libra2, Iris2024...) và lọc các cụm không cần bung.
   - **Xác định Unit $O(N)$**: Thuật toán duyệt ngăn xếp bộ nhớ (26ms / 10.000 dòng) thay thế hoàn toàn bảng tra cứu tĩnh 5.619 dòng `Hamtimlinhkienthuoc_UNIT_naotren_BOM.xlsx`.
   - **Đối soát chéo 3 chiều**: So sánh CTTT (Chỉ thị thao tác) vs PLM vs R3, đối chiếu số lượng thập phân (hỗ trợ dấu phẩy `1,5`), chuẩn hóa Rev (`01` vs `1`, `Rev.01`, `/A`).
   - **Phán định MSI**: Động cơ quyết định 9 nhánh chuẩn hóa và kiểm tra mã cố định (fix serial tool).

4. **Giao Diện Trực Quan Non-Tech & Đa Ngôn Ngữ (i18n)**:
   - **3-Step Wizard**: Tối giản, trực quan, siêu dễ hiểu cho công nhân và kỹ sư sản xuất phổ thông.
   - **Leader / Member Workspace**: Phân hệ nâng cao cho phép quản lý station, nộp báo cáo cục bộ và xuất báo cáo tổng hợp.
   - **Đa ngôn ngữ**: Chuyển đổi tức thì giữa **Tiếng Việt**, **日本語 (Tiếng Nhật)**, và **中文 (Tiếng Trung)**.

5. **Đóng Gói & Tự Động Cập Nhật Chuẩn MP2027**:
   - Cài đặt không cần quyền Admin (`PrivilegesRequired=lowest`) qua Inno Setup.
   - Cơ chế tự động cập nhật mạng LAN (`HASH_ONLY_LAN`) nguyên tử qua file `.mpupdate` và `latest.json`, có khả năng rollback an toàn qua `previous.json`.
   - Bản thực thi độc lập **Portable (.EXE)** dung lượng 23.8 MB, chạy ngay không cần cài đặt Python.

---

## 🚀 Hướng Dẫn Khởi Động

### Cách 1: Khởi động nhanh 1 chạm (Khuyên dùng)
Nhấp đúp chuột vào file:
```cmd
Khoi_Dong_SSBOM.bat
```
*(Hoặc gõ `Run_SSBOM.bat`)*

### Cách 2: Chạy bản Portable (.EXE độc lập không cần Python)
Nhấp đúp chuột vào file:
```cmd
Khoi_Dong_Portable.bat
```
*(Hoặc mở trực tiếp `dist\SSBOM_Portable\SSBOM_Portable.exe`)*

### Cách 3: Khởi động qua Launcher Python
```powershell
python SSBOM_Launcher.py
```

### Cách 4: Mở giao diện Chuyên sâu Leader / Member Workspace
```powershell
python -m src.gui.app
```

---

## 🧪 Kiểm Thử Tự Động (Test Suite)

Hệ thống sở hữu bộ kiểm thử tự động toàn diện đạt tỷ lệ vượt qua **100% (464 / 464 tests PASS)**:

- **Tier 1: Feature Isolation (F01 - F28)**: 140 tests
- **Tier 2: Boundary & Challenger Probes**: 43 tests
- **Tier 3: Multi-Module Combinations**: 14 tests
- **Tier 4: Factory Ground Truth (Khớp 1:1 dữ liệu VBA cũ)**: 10 tests
- **Tier 5: Adversarial Hardening & Stress (50.000 nodes)**: 59 tests
- **Unit & Architecture Tests**: 198 tests

Chạy toàn bộ test suite:
```powershell
pytest tests/ -q
```

---

## 📁 Cấu Trúc Dự Án (Flat Production Layout)

```
pm_sosanhbom/
├── Khoi_Dong_SSBOM.bat        # Batch khởi động thông minh 1-chạm
├── Khoi_Dong_Portable.bat     # Batch mở trực tiếp bản Portable EXE
├── Run_SSBOM.bat              # Alias tiếng Anh
├── SSBOM_Launcher.py          # Bộ khởi chạy MP2027 có auto-update & rollback
├── current.json               # Pointer phiên bản đang kích hoạt
├── apps/                      # Các phiên bản ứng dụng (1.0.0, ...)
├── locales/                   # Tệp từ điển đa ngôn ngữ (vi.json, ja.json, zh.json)
├── src/                       # Mã nguồn hệ thống
│   ├── automation/            # Web TC14 (Selenium) & SAP R3 CS12 (win32com)
│   ├── core/                  # Engine so sánh BOM, parser cây, unit resolver
│   ├── gui/                   # Giao diện đồ họa PyQt6 (Wizard, Leader, Member)
│   ├── reporting/             # Xuất báo cáo Excel và gửi email Outlook
│   └── services/              # Dịch vụ điều phối nghiệp vụ
├── tests/                     # Toàn bộ test suite 5 tầng (464 tests)
├── scripts/                   # Script đóng gói package_app.py
└── installer/                 # Cấu hình Inno Setup (SSBOM_Manager.iss)
```

---

## 📄 Bản Quyền & Giấy Phép
Dự án được bảo vệ và phát hành theo giấy phép đính kèm trong tệp [LICENSE](LICENSE).
