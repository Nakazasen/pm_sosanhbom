# Original User Request

## 2026-09-17T02:47:26Z

Viết lại toàn bộ hệ thống "Chương trình so sánh BOM tự động" (thay thế hoàn toàn Excel VBA, VBScript và quy trình thủ công) thành phần mềm tự động hóa hiện đại bằng Python, hoạt động bền vững lâu dài, tích hợp web Siemens Teamcenter Active Workspace (TC14) và SAP R3 ERP.

Working directory: D:\Sandbox\pm_sosanhbom
Integrity mode: development

## Requirements

### R1. Web Automation: Teamcenter Active Workspace (TC14)
- Tự động đăng nhập vào cổng web http://tcmp3gwb:3000/ với tài khoản vn_pe03 / vn_pe03 (sử dụng Selenium / Edge / Chrome WebDriver đã được kiểm chứng hoạt động tốt).
- Tự động tìm kiếm mã hàng, xử lý bung cây phân cấp BOM hoặc tải BOM Full theo quy tắc xuất dữ liệu PLM.
- Có cơ chế xử lý lỗi mạng, timeout, hoặc session hết hạn.

### R2. SAP R3 Automation: Download BOM Multilevel (CS12)
- Tự động hóa qua Python win32com.client kết nối SAP GUI Scripting (saplogon.exe 770, hệ thống P1J(ERP60-AWS)-VN).
- Tự động đăng nhập, gọi transaction CS12 với Plant 2200, BOM Usage pp01, Alternative 01 và ngày hiệu lực chỉ định.
- Tự động lưu và phân loại file xuất .xls vào đúng thư mục hướng xuất.

### R3. Core Comparison & Tree Algorithm Engine
- Kế thừa và chuẩn hóa toàn bộ logic nghiệp vụ:
  - Bộ lọc cây BOM 6 tầng (Level 1..6) theo ngày hiệu lực ("to", "UP", so sánh với ngày hiện tại).
  - Bóc tách theo từng Model máy (Virgo, Libra2, Iris2024...) và lọc bỏ các cụm linh kiện không cần bung.
  - Thuật toán xác định linh kiện thuộc Unit nào (thay thế bảng 5.619 dòng Hamtimlinhkienthuoc_UNIT_naotren_BOM.xlsx bằng giải thuật duyệt cây nhanh).
  - Đối soát chéo CTTT (Chỉ thị thao tác) vs PLM vs R3, nhận diện sai khác số lượng, mã đổi, thêm bớt.
  - Phán định dữ liệu quản lý MSI và mã cố định 3 ký tự (từ fix_serial tool).

### R4. Giao diện Người dùng (Desktop GUI / Modern Architecture)
- Xây dựng giao diện trực quan bằng PyQt6 (hoặc CustomTkinter):
  - Phân hệ cho Leader: Tạo thư mục hướng xuất, kiểm tra tiến độ nộp danh sách của thành viên, chạy tổng hợp BOM tự động, xuất báo cáo tổng hợp, gửi email thông báo qua Outlook.
  - Phân hệ cho Thành viên: Nhập danh sách linh kiện phụ trách (CTTT, MSI, Label 7980/7990), tự đối soát sơ bộ PLM/R3 của cụm mình và gửi xác nhận OK.

### R5. Đóng Gói & Tự Động Cập Nhật Chuẩn HASH_ONLY_LAN (Theo MP2027 Standard)
- Tuân thủ 100% tài liệu hướng dẫn `D:\Sandbox\MP2027\huongdansetup_autoupdate.md`:
  - **Kiến trúc 2 tầng**: `<App>_Launcher.exe` + `current.json` ở root; `apps/<version>/` chứa onedir app + `manifest.json`.
  - **Inno Setup**: Setup cài đặt per-user `{localappdata}` (`PrivilegesRequired=lowest`, `lzma2`).
  - **Auto-Update LAN**: Sinh gói `.mpupdate` (ZIP chứa manifest + files) và catalog `latest.json`, publish atomically qua file tạm `.part` lên thư mục mạng LAN.
  - **An toàn & Rollback**: Safe extraction chống zip-slip, health-check `--health-check` trước khi kích hoạt, backup SQLite tự động và rollback atomic qua `previous.json`.

## Acceptance Criteria

### Verification & Testing
- [x] Đã bóc tách 100% mã nguồn VBA hiện tại (locbomfull, DownloadAutoR3, msi, md_TaoFileSSB, capnhat_PLM_R3).
- [x] Đã kiểm thử kết nối cổng Web http://tcmp3gwb:3000/ (Status 200 OK, xác định kiến trúc React SPA).
- [x] Đã kiểm thử khởi chạy thành công cả Edge WebDriver và Chrome WebDriver.
- [ ] Chạy thử nghiệm trích xuất dữ liệu đăng nhập từ web http://tcmp3gwb:3000/.
- [ ] Xây dựng bộ test so sánh kết quả logic Python với file mẫu form_ssbom.xlsm để đảm bảo kết quả đối soát khớp 100%.
- [ ] Đóng gói chương trình theo đúng chuẩn MP2027 (`<App>_Launcher.exe`, Inno Setup installer, bundle `apps/<version>/` và script `package_app.py` sinh `.mpupdate` + `latest.json`).
