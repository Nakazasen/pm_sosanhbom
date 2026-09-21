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

## 2026-09-18T04:32:25Z

Lên đặc tả kỹ thuật theo chuẩn Spec-Kit và lập trình hoàn chỉnh tính năng Tự động tải BOM PLM từ hệ thống Siemens Teamcenter Active Workspace (TC14) sang tệp Excel theo 7 phân hệ trong tài liệu hướng dẫn tai_lieu_huong_dan_download_BOM.pptx.

Working directory: D:\Sandbox\pm_sosanhbom
Integrity mode: development

## Source Specification Context (Tài liệu gốc PPTX)
- Đường dẫn: \\fstvn01\Data\10_Production Engineering Department(製造技術部)\02.製造技術課\PE Dept\15. FORM（BIEU MAU）-形式\Form_VBA\Form_VBA_SS_BOM\tai_lieu_huong_dan_download_BOM.pptx
- Gồm 7 phân hệ nghiệp vụ chính (Phases 1-7):
  1. AUTH-01: Đăng nhập và lưu thông tin tài khoản an toàn (Hỗ trợ Silent Auto-login, có quản lý cập nhật/xóa thông tin).
  2. BOM-SEARCH-01: Tìm Item theo mã và điều hướng vào tab cấu trúc Content.
  3. BOM-EXPAND-01: Mở rộng phân cấp BOM tự động đến mức sâu tối đa (Level 7) ở chế độ ngầm (Headless Mode).
  4. BOM-SELECT-01: Chọn toàn bộ BOM (Select All) và mở menu lệnh Excel Report.
  5. BOM-EXPORT-OPEN-01: Kích hoạt tác vụ "Export to Excel" (bảo vệ chống chọn nhầm "Import Changes").
  6. BOM-EXPORT-CONFIG-01: Cấu hình nguồn thuộc tính (Item) và thứ tự hiển thị chuẩn xác 14 cột Displayed Columns.
  7. BOM-EXPORT-RUN-01: Kích hoạt xuất dữ liệu (Run in Background), kiểm soát tiến trình và tải file an toàn.
  8. SPEC-KIT: Biên soạn bộ tài liệu đặc tả chuẩn Spec-Kit (specs/SPEC_PLM_AUTO_DOWNLOAD.md).
  9. REPORTING: Cơ chế cập nhật và báo cáo tiến độ định kỳ (2 phút/lần) trên tổng số phase thực hiện.

## Requirements

### R1. Spec-Kit Specification (specs/SPEC_PLM_AUTO_DOWNLOAD.md)
- Soạn thảo tài liệu đặc tả đầy đủ theo phương pháp Spec-Kit: User Stories, Acceptance Criteria (Given-When-Then), Input/Output Schema, Edge Cases, Error Handling Scenarios cho cả 7 phase nghiệp vụ.

### R2. Teamcenter Authentication & Credential Persistence (AUTH-01)
- Triển khai xác thực tự động với cổng Teamcenter Active Workspace.
- Cơ chế lưu trữ thông tin đăng nhập an toàn (mã hóa chuẩn hệ điều hành / DPAPI hoặc keyring an toàn).
- Mặc định tự động đăng nhập ngầm (Auto-login) khi đã lưu thông tin hợp lệ; cung cấp giao diện/lệnh để cập nhật hoặc xóa thông tin tài khoản khi cần.

### R3. Headless Search & Navigation Pipeline (BOM-SEARCH-01)
- Tìm kiếm mã Item trên ô tìm kiếm toàn cục, xác thực mã và revision được phát hiện.
- Tự động chuyển hướng chính xác vào tab Content để nạp cây cấu trúc BOM.
- Xử lý các tình huống biên: Mã linh kiện không tồn tại, phiên làm việc hết hạn, mạng chập chờn.

### R4. Deep BOM Tree Expansion (BOM-EXPAND-01)
- Tự động chọn nút gốc (Root node) của cấu trúc BOM.
- Kích hoạt menu Expand -> "Expand Below" và bung tự động đến độ sâu tối đa (Level 7) nhằm đảm bảo không sót linh kiện con.
- Vận hành ổn định ở chế độ không đầu (Headless Mode), có cơ chế chờ linh hoạt (Dynamic Explicit Waits) tránh lỗi gián đoạn do DOM chưa nạp xong.

### R5. BOM Selection & Safe Export Trigger (BOM-SELECT-01 & BOM-EXPORT-OPEN-01)
- Thực thi "Select All" toàn bộ tập dòng BOM đã mở rộng.
- Mở menu Excel Report và kích hoạt chính xác tác vụ "Export to Excel" (tuyệt đối không chọn "Import Changes").
- Duy trì nguyên vẹn tập dòng được chọn khi panel Export mở ra.

### R6. Exact 14-Column Configuration (BOM-EXPORT-CONFIG-01)
- Lựa chọn nguồn thuộc tính "Item".
- Cấu hình chuẩn xác danh sách Displayed Columns theo đúng thứ tự 14 cột tiêu chuẩn của BOM PLM:
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
- Ngăn chặn lỗi trùng cột, thiếu cột hoặc sai thứ tự cột.

### R7. Resilient Download Pipeline & Integrity Verification (BOM-EXPORT-RUN-01)
- Kích hoạt lệnh Export (hỗ trợ tác vụ nền nếu cây BOM lớn).
- Giám sát trạng thái sinh tệp, tự động phát hiện khi tệp sẵn sàng và tải về thư mục làm việc an toàn.
- Xác thực tệp tải về: Đảm bảo mở được bằng thư viện Excel, bảo toàn bảng mã Unicode tiếng Nhật (MS Gothic), đầy đủ số dòng và đúng 14 cột.

### R8. Scheduled Periodic Progress Reporter
- Cung cấp cơ chế thông báo tiến độ định kỳ (chu kỳ 2 phút / 1 lần) thể hiện: Phase [X/8] - [Tên phân hệ] - [Tỷ lệ hoàn thành %] - [Thời gian trôi qua].

## Acceptance Criteria

### Spec-Kit & Verification
- [ ] Tệp specs/SPEC_PLM_AUTO_DOWNLOAD.md được khởi tạo hoàn chỉnh theo định dạng Spec-Kit, bao phủ 100% 7 phân hệ từ PPTX.
- [ ] Bộ kiểm thử tự động (Unit test / Integration test / Mock browser tests) bao phủ 7 phân hệ và đạt 100% Passed.
- [ ] Bộ điều khiển trình duyệt Headless WebDriver (Edge / Chrome) khởi chạy tin cậy trên Windows.
- [ ] Cấu hình cột Displayed Columns khớp 14/14 cột theo đúng thứ tự và tên chuẩn.
- [ ] Cơ chế lưu trữ tài khoản đảm bảo an toàn, hỗ trợ Auto-login và cho phép cập nhật/xóa.
- [ ] Trình báo cáo tiến độ tự động cập nhật đúng chu kỳ 2 phút / lần.

## 2026-09-18T04:33:53Z

[CHỈ THỊ CỦA NGƯỜI DÙNG]
Người dùng đính chính: Phiên bản Teamcenter của hệ thống là Teamcenter Version 2412 (TC2412), KHÔNG PHẢI Teamcenter 14.
Yêu cầu:
1. Cập nhật Spec-Kit `specs/SPEC_PLM_AUTO_DOWNLOAD.md` chuẩn hóa theo Siemens Teamcenter Version 2412 (Active Workspace).
2. Cập nhật tất cả module tự động hóa (automation client, selectors, DOM locators) và test suites để tương thích tuyệt đối với Teamcenter 2412.
3. Báo cáo tiến độ cập nhật thông tin này trong các chu kỳ báo cáo tiếp theo.

## 2026-09-18T04:35:40Z

[CẬP NHẬT CHI TIẾT TỪ TÀI LIỆU PPTX VÀ CHỈ THỊ NGƯỜI DÙNG]
1. Đã bóc tách 100% nội dung speaker notes gốc từ 7 slide của file:
   \\fstvn01\...\tai_lieu_huong_dan_download_BOM.pptx
   và 10 hình ảnh chụp màn hình UI thực tế tại:
   D:\Sandbox\pm_sosanhbom\scratch\pptx_inspect\extracted\ (image1.png -> image10.png)
   Nội dung đặc tả thô đã được nạp sẵn tại:
   D:\Sandbox\pm_sosanhbom\specs\SPEC_PLM_AUTO_DOWNLOAD.md

2. Chốt câu trả lời cho các 'open_questions' trong từng slide:
   - AUTH-01: Silent Auto-login bằng DPAPI/keyring an toàn; có chức năng đổi/xóa tài khoản.
   - BOM-SEARCH-01: Tìm chính xác Item ID, điều hướng vào tab Content.
   - BOM-EXPAND-01: Bung tự động độ sâu tối đa (Level 7) ở chế độ Headless WebDriver.
   - BOM-SELECT-01: Select All toàn bộ các dòng BOM đã bung.
   - BOM-EXPORT-OPEN-01: Kích hoạt Export to Excel (ngăn chặn Import Changes).
   - BOM-EXPORT-CONFIG-01: Nguồn thuộc tính "Item", đúng 14 cột chuẩn:
     Home, Level, Item Type, Item Id, Has Children, Quantity, 1st Parts, 2nd BOM Flag, Occurrence Effectivities, Item Revision Project List, Item Name, Notice No, Revision, Item Rev Status.
   - BOM-EXPORT-RUN-01: Run in Background, giám sát tải tệp về, kiểm tra font tiếng Nhật MS Gothic và 14 cột.
   - Target System: Siemens Teamcenter Version 2412 (TC2412).

Yêu cầu: Orchestrator và các agent triển khai bám sát 100% tài liệu này để hoàn thành Spec-Kit và mã nguồn.

## 2026-09-18T04:57:28Z

[CHỈ THỊ CẤP BÁCH & PHÁT HIỆN FORENSIC TỪ NGƯỜI DÙNG]
Người dùng cung cấp 2 tệp mẫu đối chiếu:
1. PLM cũ (TC14): `PLM_1102Z53KR0 Cũ.xlsx` (14 cột: Col C=Item Type, Col D=Item Id, Col K=Item Name, Col M=Revision).
2. PLM mới (TC2412): `T10C423NL0 Mới.xlsm` (24 cột: Col B=Level, Col C=Item Type, Col D=Name [Mã LK], Col H=Parts Text [Tên LK], Col J=Revision, Col K=Release Status).
3. File BOM chung: `BOM_110C0Z3LV1.xlsm` Sheet `PLM`.

PHÁT HIỆN TỬ HUYỆT CÔNG THỨC EXCEL TRONG FILE BOM CHUNG:
- Sheet `Tongket!C5` & `CTTT!A1` có công thức: `=PLM!C2` (yêu cầu Mã máy phải ở Cột C Sheet PLM).
- Sheet `CTTT` có hàng loạt công thức: `=VLOOKUP(C3, PLM!C:L, 10, 0)` (tra cứu mã ở Cột C, lấy Revision ở Cột L của Sheet PLM).
- Sheet `CTTT` có công thức: `=VLOOKUP(C3, PLM!T:U, 2, 0)` (tra cứu tổng số lượng trong bảng Pivot ở cột T:U).
- Sheet `PLM` cột R có `=IF(C2="","",C2)`, cột S có `=IF(E2="","",E2)`.

YÊU CẦU THIẾT KẾ BẮT BUỘC CHO MILESTONE M3 & M4 (TC2412 CANONICAL NORMALIZER & SHEET PLM BRIDGE):
1. Chuẩn hóa 100% đầu vào theo Teamcenter Version 2412 mới (đọc 24 cột của TC2412, map D='Name' thành item_id, H='Parts Text' thành item_name, B='Level', K='Release Status').
2. Khi ghi/xuất ra Sheet 'PLM' của file BOM chung (`BOM_*.xlsm`), BẮT BUỘC ánh xạ chính xác vào đúng vị trí cột truyền thống:
   - Cột A: Level
   - Cột B: Item Type
   - Cột C: Item Id (lấy từ Name của TC2412) -> Đảm bảo =PLM!C2 và VLOOKUP(..., PLM!C:L, ...) chạy đúng 100%!
   - Cột D: Has Children
   - Cột E: Quantity
   - Cột J: Item Name (lấy từ Parts Text của TC2412)
   - Cột L: Revision
   - Cột M: Item Rev Status (lấy từ Release Status của TC2412)
   - Cột R: PART CODE, Cột S: Q.TY, Cột T:U: Pivot Table.
3. Không làm gãy bất kỳ công thức VLOOKUP hay liên kết nào của các thành viên trong các sheet Tongket, CTTT!
Yêu cầu Orchestrator và Worker đưa cơ chế Bridge này vào Spec-Kit và mã nguồn ngay lập tức.

## 2026-09-19T09:43:14Z

Nâng cấp và hoàn thiện 100% hệ thống Phần mềm So Sánh BOM Tự Động (Kyocera Desktop App) bằng Python/PyQt6, tái cấu trúc toàn diện UX/UI theo luồng nghiệp vụ chuẩn từng bước và phục hồi đầy đủ tất cả các tính năng từ mã nguồn VBA gốc (tonghop_new, formnguoidung, form_ssbom) và tài liệu hướng dẫn chuẩn Chương trình so sánh BOM tự động.pptx.

Working directory: D:\Sandbox\pm_sosanhbom
Integrity mode: development

## 1. Context & Business Domain (Quy trình chuẩn Kyocera)
Hệ thống vận hành theo chu trình khép kín giữa Trưởng nhóm (Leader) và Kỹ sư phụ trách (Member):
- Leader Workflow:
  - Giai đoạn: maT (cho DMT / PMT) và ma1 (từ PP trở đi).
  - Quản lý nhân sự theo phòng ban (tenphong_pt): Cơ 1, Cơ 2, Cơ 3.
  - Lập dự án & Phân công mã máy/hướng xuất cho từng kỹ sư (Sheet Lichsu). Tự sinh thư mục máy và sinh file riêng mang tên kỹ sư.
  - Tải tự động BOM R3 trên SAP (hỗ trợ nhập 1 ngày chung hoặc ngày riêng cho từng mã).
  - Lọc BOM Full PLM Teamcenter TC24 theo bộ lọc cấp bậc Level 1..6 và cấu hình BolocBom (xử lý ngày hiệu lực from...to..., loại bỏ cụm không bung).
  - Quét trạng thái nộp bài (kiểm tra ô Q2 = OK).
  - Tổng hợp dữ liệu thành viên (CTTT, MSI, Label 7980/7990), tạo thư mục lưu trữ phutrach.
  - Tạo file So sánh BOM tổng (form_ssbom), Refresh Pivot Table, lọc dòng sai khác.
  - Kế thừa nội dung giải trình (PLM_old sang PLM mới qua ham_match_index_mix).
  - Quản lý Master List JIG (file 'List JIG thay doi, khi bo sung ma hang.xlsx') & Đánh giá xác nhận 4M với KTSX.
  - Gửi Mail tự động qua Outlook 2 luồng riêng biệt: Luồng 1 (nhắc thành viên nộp bài + 18 điểm kiểm tra trước sản xuất), Luồng 2 (gửi Quản lý kiểm tra xác nhận).
- Member Workflow:
  - Tự động nhận diện đúng tên kỹ sư và mã máy được Leader phân công.
  - Nhập 3 danh mục: CTTT, MSI, Nhãn 7980/7990.
  - Đối chiếu MSI với Master FIX_SERIAL_DLTOOL_VER010.xls (Sheet UNIT, MACHINE).
  - Tự nạp BOM PLM & R3 để đối soát sơ bộ tại chỗ (Self-check) và giải trình sai khác.
  - Bấm xác nhận nộp bài (Đóng dấu Q2 = OK).

## 2. Requirements (Speckit Functional Breakdown)
### R1. Tái cấu trúc giao diện Leader Workspace thành Luồng Wizard 4 bước tuần tự
- Bước 1 (Lập Dự Án & Phân Công): Chọn Model, Giai đoạn (maT/ma1). Nhập danh sách mã BOM. Bảng phân công nhân sự theo phòng ban Cơ 1, Cơ 2, Cơ 3 (tương đương Sheet Lichsu). Nút bấm tự động tạo thư mục mã hàng và khởi tạo file/gói nộp mang tên từng kỹ sư.
- Bước 2 (Tải & Xử Lý Nguồn Dữ Liệu): Tích hợp hộp thoại tải BOM PLM TC24 & SAP R3 (hỗ trợ ngày chung / ngày riêng). Tự động phân chia file về từng thư mục mã máy.
- Bước 3 (Theo Dõi & Tổng Hợp): Bảng quét trạng thái nộp bài theo thời gian thực (Đã nộp OK / Chưa nộp). Chỉ cho phép bấm nút Tổng Hợp Dữ Liệu khi toàn bộ thành viên đã xác nhận OK. Tự động gom CTTT, MSI, 7980/7990 và cất file thành viên vào thư mục phutrach.
- Bước 4 (So Sánh BOM Tổng & Gửi Báo Cáo): Tạo file form_ssbom, cập nhật Pivot Table. Quản lý Master JIG & đánh giá 4M. Xem trước và gửi Email Outlook 2 tầng (Gửi thành viên & Gửi quản lý).

### R2. Bộ lọc Cấu trúc BOM PLM TC24 Full (BOM Filter Engine Level 1..6)
- Xử lý tệp BOM PLM Full xuất từ Teamcenter: Quét cột hiệu lực I (from ... to ...), tự động loại bỏ linh kiện đã hết hạn so với thời điểm hiện tại và xóa sạch toàn bộ các tầng con bên dưới (Level BOM 1 đến 6).
- Áp dụng cấu hình bộ lọc theo từng model máy (BolocBom): Hỗ trợ khớp chính xác Full_name, khớp tương đối Part_name, và danh sách mã cụ thể không bung.
- Xóa đệ quy toàn bộ linh kiện con khi linh kiện cha bị chặn.
- Tự động tạo thư mục sao lưu backupTC14full.

### R3. Cơ chế Kế thừa Giải trình khi Cập nhật BOM Mới (PLM_old sang PLM mới)
- Khi cập nhật dữ liệu PLM hoặc R3 mới: Tự động lưu trữ Sheet hiện tại thành PLM_old, nạp dữ liệu mới vào Sheet PLM, tự động tra cứu và bảo lưu toàn bộ nội dung Giải thích, Phụ trách, Quản lý check từ PLM_old sang PLM mới cho các linh kiện không đổi, tự động chuyển file cũ vào thư mục capnhat\old\ và refresh Pivot Table.

### R4. Module Đối soát MSI Chuyên sâu với Master FIX_SERIAL_DLTOOL
- Tra cứu đối soát tự động với file Master FIX_SERIAL_DLTOOL_VER010.xls: Sheet UNIT (3 ký tự cố định MSI và LABEL_COMMENT SERVICE), Sheet MACHINE (3 ký tự cố định cho HONTAI), đối chiếu với Sheet PLM và phán định OK / NG, tô màu tương phản trực quan.

### R5. Module Quản lý Danh mục List JIG Master & Đánh giá 4M
- Liên kết với file Master 'List JIG thay doi, khi bo sung ma hang.xlsx', tự động lọc và nạp danh sách JIG tương ứng theo loại máy vào Sheet List JIG, giữ nguyên định dạng và công thức, tích hợp ô tích chọn / xác nhận đánh giá hạng mục 4M với bộ phận Kỹ thuật Sản xuất (KTSX).

### R6. Hoàn thiện Member Workspace theo chuẩn Form Người Dùng
- Giao diện thành viên tự động tải thông tin phân công (Tên kỹ sư, mã máy, phòng ban), loại bỏ ô text tự gõ mẫu.
- Cho phép nhập và kiểm tra 3 bảng: CTTT, MSI, Label 7980/7990.
- Tự động nạp BOM PLM & R3 trong thư mục máy để kỹ sư tự đối soát sơ bộ (Self-check).
- Nút Xác nhận Nộp (Đóng dấu Q2 = OK) để báo cho Leader biết bài nộp đã hoàn tất.

## 3. Acceptance Criteria
- Màn hình Leader hiển thị rõ ràng 4 tab/bước theo đúng quy trình từ trái sang phải.
- Bảng danh sách nhân sự Cơ 1, Cơ 2, Cơ 3 chuẩn và bảng phân công ma trận (Sheet Lichsu), bấm 1 nút là tự sinh thư mục máy + file phân công.
- Quét trạng thái nộp bài chính xác theo cờ Q2 = OK, hiển thị rõ ai chưa nộp.
- Tổng hợp đủ cả 3 phần: CTTT, MSI, Label 7980/7990 vào file tổng.
- Lọc chính xác các dòng hết hiệu lực và xóa đúng các cấp con (Level 1..6) theo thuật toán VBA gốc.
- Khi cập nhật PLM mới, toàn bộ giải thích cũ của các linh kiện không đổi được giữ nguyên vẹn 100%.
- Đối soát MSI khớp chính xác 3 ký tự cố định từ FIX_SERIAL_DLTOOL_VER010.xls.
- Nạp đúng bảng JIG và có trường xác nhận đánh giá 4M.
- Toàn bộ test suite tự động (Pytest) đạt 100% PASS.
- Chạy lệnh package_app.py thành công, các file .bat khởi chạy ứng dụng mượt mà không lỗi.

## 2026-09-21T01:40:20Z

Tái cấu trúc và nâng cấp toàn diện hệ thống giao diện người dùng (PyQt6 Desktop) của công cụ Quản Lý & So Sánh BOM (pm_sosanhbom) theo phong cách Data-Dense Enterprise Dashboard chuyên biệt cho kỹ sư Khối Sản Xuất Kỹ Thuật (PE Dept - Manufacturing & Production Engineering), tuân thủ nghiêm ngặt bộ quy chuẩn ui-ux-pro-max và tài liệu đặc tả chuẩn hóa specs/SPEC_UI_UX_ENTERPRISE_DASHBOARD.md.

Working directory: d:\Sandbox\pm_sosanhbom
Integrity mode: development
Authoritative Spec: specs/SPEC_UI_UX_ENTERPRISE_DASHBOARD.md

## 1. Bối cảnh & Hiện trạng Đặc tả (Spec-Kit Baseline)
- Đã xác thực thư mục specs/: Trước đó chỉ có SPEC_PLM_AUTO_DOWNLOAD.md (chuyên trách tải tự động từ TC2412).
- Đã tạo mới tài liệu đặc tả chuẩn hóa: specs/SPEC_UI_UX_ENTERPRISE_DASHBOARD.md bao gồm 7 chương đặc tả chi tiết về Design Tokens, bảng màu WCAG 4.5:1 / 7:1, kích thước Typography, độ cao dòng bảng 32px, viền bảng sắc nét (#CBD5E1 / #2A374A), và ma trận kiểm thử tự động.

## 2. Các yêu cầu cốt lõi (Requirements)

### R1. Kiến trúc Quản lý Giao diện Tập trung (ThemeManager & Design Tokens)
- Xây dựng hệ thống quản lý theme tại src/gui/styles/:
  - tokens.py: Tập trung định nghĩa bảng mã màu, font chữ, độ cao dòng bảng, kích thước viền và padding chuẩn Data-Dense.
  - theme_manager.py: Điều phối chuyển đổi tức thời giữa Light Theme (Slate Industrial) và Dark Theme (Industrial Dark Mode) không cần khởi động lại app.
  - light_theme.qss & dark_theme.qss: Bộ stylesheet tối ưu cho PyQt6, không có lỗi phân giải cú pháp.

### R2. Nâng cấp Toàn diện các Màn hình theo Phong cách Data-Dense
- Leader View (src/gui/leader_view.py):
  - Bổ sung 4 thẻ KPI Cards tinh gọn: Tổng số Model, Model đã khớp BOM, Model lệch số lượng, Sẵn sàng CTTT.
  - Bảng Sourcing & Kế thừa BOM: Viền lưới ô sắc nét 1px, độ cao dòng chuẩn 32px, padding ô 4px 8px, zebra striping nhẹ nhàng, hiệu ứng hover dòng rõ ràng (#F1F5F9 / #1A2436).
- Member View (src/gui/member_view.py):
  - Tinh chỉnh Stepper 3 bước với icon SVG chỉ báo trạng thái (đang làm, hoàn thành, cảnh báo).
  - Khung Diff View hiển thị kết quả so sánh BOM nổi bật các dòng lệch màu đỏ nhạt (#FEE2E2 / #7F1D1D), dòng khớp màu xanh (#DCFCE7 / #064E3B).
- Hộp thoại Tải BOM PLM (src/gui/plm_download_dialog.py):
  - Layout hiện đại, thanh tiến độ 14px tinh tế kèm % số lượng, khung log Consolas sắc nét.
- Hộp thoại Cài đặt (src/gui/settings_dialog.py):
  - Tích hợp bộ chọn giao diện (Light / Dark / Theo Windows) và lưu cấu hình bền vững.

### R3. Hệ thống Biểu Tượng SVG Chuẩn Hóa (No Emoji)
- Thay thế toàn bộ ký tự emoji thô bằng bộ icon SVG vector sắc nét đặt tại src/gui/assets/icons/ (Lucide / Fluent design).
- Hỗ trợ đổi màu icon theo Theme (tinting/monochrome) để luôn đạt độ tương phản chuẩn.

### R4. Bảo đảm Khả năng Tiếp cận & Độ tương phản (WCAG 2.1 Contrast)
- Văn bản chính đạt tỷ lệ tương phản tối thiểu 7:1 (AAA) so với nền bảng biểu (#0F172A trên #FFFFFF; #F1F5F9 trên #151D2A).
- Văn bản phụ / nhãn ghi chú đạt tối thiểu 4.5:1 (AA) (#475569 trên #FFFFFF; #94A3B8 trên #151D2A).
- Không dùng màu xám trên nền xám (gray-on-gray), không dùng chữ mờ gây mỏi mắt.

### R5. Bộ Kiểm thử Tự Động Hóa & Đo Lường Hợp đồng Giao diện
- Xây dựng bộ test tự động tại tests/unit/test_ui_theme.py:
  - Kiểm tra độ tương phản màu sắc đạt chuẩn WCAG toán học.
  - Kiểm tra tính hợp lệ của cú pháp QSS (không có cảnh báo QSS parse error từ Qt).
  - Kiểm tra sự tồn tại và tải thành công của 100% tệp icon SVG.
  - Kiểm tra khả năng chuyển đổi qua lại giữa 2 theme mà không gây memory leak hay sập giao diện.
- Đảm bảo 100% các bài test hiện có của hệ thống vẫn PASS (Regression Zero).

## 3. Tiêu chí nghiệm thu (Acceptance Criteria)

### A. Độ tương phản & Thẩm mỹ Data-Dense
- [ ] 100% các cặp màu chữ và nền đạt tỷ lệ tương phản >= 4.5:1 (WCAG AA) và >= 7.0:1 cho văn bản bảng biểu chính (WCAG AAA).
- [ ] Tất cả các bảng biểu (Sourcing Table, BOM Diff Table, Log Table) có đường viền ô rõ ràng, dòng chẵn lẻ phân biệt nhẹ, chiều cao dòng tối ưu từ 30–34px.
- [ ] Không xuất hiện emoji làm icon giao diện; 100% nút bấm dùng icon SVG chuẩn Lucide/Fluent.

### B. Chức năng Chuyển đổi Theme (Light & Dark)
- [ ] Người dùng có thể chuyển đổi mượt mà giữa Light Theme và Dark Theme trong Cài đặt hoặc phím tắt/menu.
- [ ] Trạng thái theme được lưu lại và tự động tải đúng trong phiên làm việc tiếp theo.

### C. Độ ổn định & Không làm hỏng chức năng cũ (Zero Regression)
- [ ] Ứng dụng PyQt6 mở lên bình thường, các luồng nghiệp vụ (Download BOM PLM TC2412, Lọc cơ khí, Đối soát CTTT, Xuất Excel) hoạt động trơn tru 100%.
- [ ] Lệnh kiểm thử toàn diện pytest tests/ vượt qua 100% bài kiểm thử.

