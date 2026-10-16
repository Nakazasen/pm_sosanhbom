# Hướng dẫn sử dụng SSBOM Manager (v1.0.7)

Tài liệu này mô tả phần mềm theo một lần chạy thật cho mã máy **110C3M3AK0** (Model Virgo). Ảnh chụp được lấy trực tiếp từ phần mềm trong lần chạy đó. Ảnh chụp trên v1.0.6 có thanh trạng thái ghi v1.0.6; thao tác và giao diện của các bước này không đổi ở v1.0.7.

Tài liệu có 3 ngôn ngữ (Tiếng Việt / 日本語 / 中文). Trong file HTML, dùng nút chuyển ngôn ngữ ở đầu trang và **menu bên trái** để nhảy tới từng mục.

## 1. Cài đặt và cập nhật phần mềm

### 1.1 Cài đặt lần đầu

Chạy file cài đặt tại:

```
\\fstvn01.kdtvn.local\Data\00_KDTVN Common(KDTVN共通)\⑤Production Engineering(製造技術)\Hang muc can luu\Vinh\Pm_sosanhBOM\SSBOM_Manager_Setup_1.0.8.exe
```

Cài xong, mở phần mềm. Số phiên bản hiển thị ở góc dưới bên phải (v1.0.8).

### 1.2 Cập nhật phiên bản mới qua mạng LAN

Khi trên thư mục LAN `release_update` có phiên bản mới hơn, phần mềm hiện hộp thoại "Có Phiên Bản Mới Sẵn Sàng!" với phiên bản hiện tại, phiên bản mới, dung lượng gói và nội dung thay đổi. Chọn "Tải & Cập Nhật Ngay" để cập nhật, hoặc "Để sau".

![Hộp thoại cập nhật phần mềm](images/g27_software_update.png)

Ảnh trên được dựng bằng cách mở hộp thoại với thông tin gói 1.0.7 thật (phiên bản hiện tại ghi v1.0.6 để minh họa). Việc tải và cài tự động chưa được chạy lại khi làm tài liệu.

## 2. Dữ liệu dùng trong ví dụ

Ví dụ lấy từ file mẫu `C3MAKO - Copy.zip`, gồm:

| File | Nội dung |
|---|---|
| `PLM_110C3M3AK0.xlsx` | BOM xuất từ PLM (Teamcenter), 1890 dòng |
| `BOM_110C3M3AK0.xlsm` | BOM mẫu đã làm xong từ trước. Sheet CTTT (2198 dòng) và sheet R3 được dùng làm dữ liệu đầu vào của ví dụ |

Lưu ý khi đọc ví dụ:

- Không có kết nối Teamcenter và SAP khi làm tài liệu, nên file R3 trong ví dụ được xuất từ sheet R3 của BOM mẫu, không phải file xuất trực tiếp từ SAP.
- Dữ liệu CTTT của 9 phụ trách được lấy từ sheet CTTT của BOM mẫu. Tên tài khoản cũ trong BOM mẫu được gán sang tài khoản trong danh sách nhân sự hiện tại (ví dụ `AnQuyDuong` → `Duong_mecha1.2`). Việc gán này là giả định của người làm tài liệu, chỉ để chạy thử.
- Thư mục làm việc trong ví dụ là `sample_data\run` và `sample_data\run_feat` (bản sao dùng thử, không phải dữ liệu thật).
- Ảnh chụp các hộp thoại cấu hình đã được thay tài khoản và mật khẩu thật bằng ký hiệu giữ chỗ.

## 3. Giao diện chung

Cửa sổ chính có thanh công cụ phía trên và 2 màn hình: "Trưởng Nhóm" và "Thành Viên Công Đoạn".

Các nút trên thanh công cụ: "Tải BOM (PLM/SAP)", "Quản lý Bộ lọc BOM PLM", "Áp dụng bộ lọc cho file có sẵn", "Cập nhật BOM mới (PLM & R3)", "Quản lý Trưởng nhóm", "Nhập liệu Thành viên", "Giao diện: Sáng [Ctrl+T]".

### 3.1 Giao diện Sáng và Tối

Bấm nút "Giao diện" trên thanh công cụ hoặc `Ctrl+T` để đổi giữa Sáng và Tối. Cũng có thể chọn trong "Cấu hình hệ thống", tab "Giao diện (UI Theme)". Đổi có hiệu lực ngay, không cần mở lại phần mềm.

![Giao diện Tối, màn hình Lập Dự Án](images/g25_dark_step1.png)

![Giao diện Tối, màn hình So Sánh BOM Tổng](images/g26_dark_step4.png)

### 3.2 Cấu hình hệ thống

Mở bằng menu "Hệ thống" → "Cấu hình hệ thống" (phím tắt `Ctrl+,`). Hộp thoại có 5 tab:

**Teamcenter TC24:** URL cổng, tài khoản, mật khẩu, trình duyệt, chế độ chạy ẩn danh (Headless), thời gian chờ. Nút "Kiểm tra kết nối TC24".

![Cấu hình TC24](images/g20_settings_tc24.png)

**SAP R3:** System Description, SAP Client, tài khoản và mật khẩu SAP, ngôn ngữ, Plant, BOM Usage, Alternative, đường dẫn `saplogon.exe`. Nút "Kiểm tra kết nối SAP GUI".

![Cấu hình SAP R3](images/g21_settings_sap.png)

**Đường dẫn Dữ liệu:** thư mục gốc dự án, file Master Fix Serial Tool, thư mục xuất báo cáo, CSDL nhân sự chung (SQLite) và thư mục cập nhật LAN.

![Cấu hình đường dẫn](images/g22_settings_paths.png)

**Giao diện (UI Theme):** chọn giao diện Sáng hoặc Tối.

![Cấu hình giao diện](images/g23_settings_theme.png)

**Email Mẫu Form:** "Chế độ Thử nghiệm" (Test Mode), email người nhận test, email người gửi mặc định, và đường dẫn các file mẫu trên LAN (từ điển loại máy, `formnguoidung.xlsm`, `form_ssbom.xlsm`).

![Cấu hình email](images/g24_settings_email.png)

Khi "Chế độ Thử nghiệm" đang bật, mọi thư được chuyển tới người kiểm thử chứ không gửi cho người nhận thật (xem mục 4.5, 4.17 và mục 6). Cần tắt chế độ này khi dùng chính thức.

## 4. Quy trình của Trưởng nhóm

Màn hình "Trưởng Nhóm" có 4 tab: "1. Lập Dự Án & Phân Công", "2. Tải & Xử Lý Nguồn BOM", "3. Theo dõi tổng hợp file của phụ trách", "4. So Sánh BOM Tổng & Báo Cáo".

### 4.1 Bước 1. Lập dự án và phân công

1. Chọn "Model máy": `Virgo`. Mã máy `110C3M3AK0` nằm ở bảng "1.2 Danh sách Mã Máy". Thư mục gốc đặt là `sample_data\run`.
2. Ở bảng "1.3 Phân Công Nhân Sự Phụ Trách", chọn 9 phụ trách và gán công đoạn (DRUM, LSU, DLP, IMAGE, FUSER, DP, ISU, HONTAI, FRAME UNIT).
3. Bấm "Khởi Tạo Thư Mục && Phân Công".

![Bước 1: cấu hình dự án](images/e01_step1_setup.png)

Phần mềm báo:

> Đã tạo xong danh sách BOM và các gói file CTTT của phụ trách được phân lập theo từng dòng máy: Model 'Virgo': 1 mã máy
> Cấu trúc: <Thư mục gốc>\<Tên Model>\<Mã máy>\<Tên phụ trách>.xlsm

![Bước 1: sau khi khởi tạo](images/e02_step1_created.png)

Kết quả trong thư mục `Virgo\110C3M3AK0\`: 9 file `<tài khoản>.xlsm` (Chien_mecha1.2, Duong_mecha1.2, Hung_mecha1.1, Ly_mecha1.2, Nhung_mecha1.2, Thai_mecha1.2, Thuong_mecha1.2, Tu_mecha1.2, VietE_mecha1.2). Các file này được sao từ `formnguoidung.xlsm`.

### 4.2 Quản lý Model và Mã máy

Bấm "Quản lý Model & Mã máy..." ở mục 1.1. Hộp thoại gồm 3 phần: chọn Model và giai đoạn sản xuất; danh sách mã máy 4 ký tự đã đăng ký của Model (ví dụ Virgo có 10 mã); nhóm email phụ trách các phòng ban (To, CC1, CC2...). Các nút: "Tạo Model Mới", "Tải lại Excel", "+ Bổ sung mã vào Model", "Lưu Nhóm Mail Vào Từ Điển", "Đưa Danh Sách Mã Này Vào Bảng 1.2", "Chọn Model Này".

![Quản lý Model và Mã máy](images/g16_model_manager.png)

### 4.3 Quản lý nhân sự

Bấm "Quản lý nhân sự..." ở mục 1.3. Hộp thoại "Danh mục Nhân sự Kỹ thuật Sản xuất" liệt kê thành viên (47 người trong ví dụ), có ô tìm theo tên hoặc mã tài khoản, lọc phòng ban, lọc "Chỉ đang hoạt động". Bên phải là form thông tin chi tiết với các nút "+ Thêm thành viên mới", "Cập nhật thông tin", "Xóa thành viên", "Làm mới ô nhập". Phía dưới: "Đồng bộ & Nạp lại từ CSDL", "Nạp lại từ Excel (Refresh)", "Ghi sang Danhsachthanhvien.xlsm", "Mở File Excel". Dữ liệu dùng chung nằm trong CSDL mạng `ssbom_master.db`.

![Quản lý nhân sự](images/g15_roster.png)

### 4.4 Quét kế hoạch PCD

Bấm "Quét Kế Hoạch PCD" ở mục 1.1. Hộp thoại có 3 phần: thông tin kế hoạch sản xuất tháng (chọn năm, tháng, file kế hoạch Excel; nút "Tải Tự Động Từ SharePoint", "Mở SharePoint PCD...", "Quét Dữ Liệu"); danh sách mã máy mới tìm được; thiết lập khởi tạo và gửi email. Nút cuối "Khởi Tạo Thư Mục & Sinh Gói File Làm Việc Cho Mã Đã Chọn" tạo thư mục cho các mã được chọn.

![Quét kế hoạch PCD](images/g17_pcd_scan.png)

Chưa chạy chức năng này với file kế hoạch thật (không có file kế hoạch tháng khi làm tài liệu). Ảnh cho thấy hộp thoại khi chưa có dữ liệu.

### 4.5 Gửi mail phân công cho phụ trách

Bấm "Gửi Mail Phụ Trách" ở mục 1.1. Phần mềm mở cửa sổ xem trước thư (người nhận, CC, tiêu đề, nội dung HTML) với các nút "Mở trong Microsoft Outlook" và "Gửi trực tiếp qua Outlook".

![Thư phân công](images/g18_mail_assign.png)

Trong ảnh, "Chế độ Thử nghiệm" đang bật: thư được chuyển tới người kiểm thử, phần đầu thư ghi người nhận thực tế khi chạy chính thức. Chưa gửi thư thật khi làm tài liệu.

### 4.6 Bước 2. Tải và xử lý nguồn BOM

Bảng "2.2 Bảng Trạng Thái Nguồn Dữ Liệu" ban đầu báo "Thiếu file" ở cả cột PLM TC24 và SAP R3 CS12. Dưới bảng có các nút: "Tải BOM Tự Động (PLM TC24 / SAP R3)", "Nạp Tệp Sẵn Có Từ Ổ Đĩa", "Lọc BOM TC24 Level 1..6 & Sao Lưu", "Cấu Hình Bộ Lọc BOM", "Tạo Lọc Trực Quan", "Lọc Lại File Có Sẵn...", "Làm mới trạng thái".

![Bước 2: chưa có file](images/e03_step2_before.png)

Có hai cách đưa file nguồn vào: tải tự động từ PLM/SAP (mục 4.7), hoặc nạp file đã có trên ổ đĩa (mục 4.8).

### 4.7 Tải BOM tự động từ PLM (Teamcenter) và SAP R3

Bấm "Tải BOM Tự Động (PLM TC24 / SAP R3)" ở Bước 2, hoặc nút "Tải BOM (PLM/SAP)" trên thanh công cụ (`Ctrl+D`). Hộp thoại có phần hướng dẫn 3 bước ngay đầu:

1. Dán danh sách mã máy/BOM vào ô bên dưới.
2. Chọn ngày hiệu lực nếu cần tải từ SAP R3 (Teamcenter TC24 tự động lấy BOM mới nhất, không cần ngày).
3. Bấm nút ở thanh dưới cùng để tải.

![Hộp thoại tải BOM, tab Teamcenter](images/g01_plm_download_tc24.png)

Các thành phần:

- "Dòng máy / Model": mặc định "Tự động nhận diện theo mã BOM".
- Ô chọn "Tự động áp dụng Bộ lọc BOM (BolocBom) theo từng loại máy ngay sau khi tải".
- "Quét Nạp mã...", "Dán từ Clipboard", "Xóa hết": nhập danh sách mã. Khi thư mục máy đã có mã, phần mềm gợi ý "Tìm thấy 1 mã BOM trong thư mục Virgo" kèm nút "Nạp tất cả (1)" và "Chi tiết...".
- Tab "Siemens Teamcenter (TC24)": tài khoản tải TC24 và "Tự động chuẩn hóa 14 cột tiêu chuẩn" cho PLM.
- Tab "SAP R3": "Mã nhà máy (Plant)", "BOM Usage", "Alternative (Alt)", "Ngày hiệu lực".
- Thanh dưới: "Tải đồng thời cả PLM & SAP R3", "Chỉ tải BOM PLM (TC24)", "Chỉ tải BOM SAP R3", "Đóng".

![Hộp thoại tải BOM, tab SAP R3](images/g02_plm_download_sap.png)

Tài khoản, mật khẩu và các thông số kết nối lấy từ "Cấu hình hệ thống" (mục 3.2). Cần điền và bấm "Kiểm tra kết nối" trước.

**Chưa chạy:** việc tải thật từ Teamcenter và SAP cần tài khoản và kết nối mạng, nên khi làm tài liệu chỉ mở hộp thoại để chụp ảnh. Phần mô tả ở trên dựa trên giao diện thực tế, không dựa trên một lần tải thành công.

**Quét mã máy từ thư mục hoặc dự án:** nút "Quét Nạp mã..." mở hộp thoại quét. Hộp thoại quét mã máy từ một thư mục hoặc từ dự án hiện tại và cho biết từng mã đã có file PLM, file R3 hay chưa. Có các bộ lọc nhanh "Tất cả", "Chỉ thiếu R3", "Chỉ thiếu PLM", "Đã đủ cả 2", và các nút "Lọc lại file có sẵn..." và "Nạp vào danh sách tải".

![Quét mã máy/BOM](images/g03_bom_scan.png)

### 4.8 Nạp file BOM có sẵn từ ổ đĩa

1. Bấm "Nạp Tệp Sẵn Có Từ Ổ Đĩa", chọn file PLM và file R3.
2. Phần mềm báo: "Đã nạp và phân tuyến 2 tệp vào các thư mục máy tương ứng." Cả hai cột chuyển sang "Sẵn sàng" (PLM 123 KB, R3 175 KB).

![Bước 2: đã nạp file](images/e04_step2_imported.png)

3. Bấm "Lọc BOM TC24 Level 1..6 & Sao Lưu". Phần mềm báo: "Đã hoàn thành lọc BOM Level 1..6 và sao lưu 1 tệp vào backupTC14full/."

Trong lần chạy này file PLM có 1890 dòng trước khi lọc và 1878 dòng sau khi lọc.

![Bước 2: sau khi lọc](images/e05_step2_filtered.png)

### 4.9 Cấu hình bộ lọc BOM PLM (BolocBom)

Bộ lọc là danh sách các cụm/linh kiện con cần loại bỏ khi phân rã BOM PLM của từng dòng máy. Bấm "Quản lý Bộ lọc BOM PLM" trên thanh công cụ, hoặc "Cấu Hình Bộ Lọc BOM" ở Bước 2.

Hộp thoại gồm: danh sách dòng máy bên trái (kèm số quy tắc, ví dụ Virgo 77 quy tắc); bảng quy tắc bên phải (Tên Linh Kiện / Cụm, Chế Độ Khớp "Chính xác (Full)" hoặc "Chứa từ (Part)", Mã Part, Ghi chú, nút sửa/xóa); form "Thêm Mới Quy Tắc Lọc". Dữ liệu lưu tập trung trên CSDL dùng chung qua LAN (nhãn "LAN Server (UNC)" ở góc trên).

Các nút phía dưới: "Áp Dụng Cho File Có Sẵn...", "Tạo Lọc Trực Quan Từ Cây BOM...", "Nhập từ Excel (Sheet BolocBom)", "Xuất ra Excel", "Khôi Phục Mặc Định", "Đóng".

![Cấu hình bộ lọc BOM](images/g04_filter_config.png)

### 4.10 Tạo bộ lọc trực quan từ cây BOM

Bấm "Tạo Lọc Trực Quan" ở Bước 2. Hộp thoại cho phép nạp file PLM Full (nút "Nạp Tệp BOM PLM Full"), xem cây BOM theo cấp (Level), mở/thu gọn cấp ("Cấp 1-2", "Mở rộng tất cả", "Thu gọn tất cả"), tìm kiếm, rồi chọn trực tiếp các cụm Phantom hoặc cụm lắp ráp cần loại bỏ. Cột "Quy Tắc Lọc Áp Dụng" cho biết dòng nào đang bị quy tắc nào áp dụng. Thanh dưới cùng hiện tổng linh kiện, số linh kiện sau khi lọc và số cụm bị áp quy tắc; nút "Lưu Vào Bộ Lọc Của Model" lưu lựa chọn thành quy tắc.

![Tạo lọc trực quan](images/g05_visual_builder.png)

Trong ví dụ, cây PLM của `110C3M3AK0` có 1,889 linh kiện; với 77 quy tắc của Virgo, 21 cụm bị áp quy tắc và còn lại 1,879 linh kiện theo số hiển thị trong hộp thoại.

### 4.11 Áp dụng bộ lọc cho file có sẵn

Dùng khi đã có file PLM trên ổ đĩa và chỉ cần chạy lại bộ lọc mà không phải tải lại từ Teamcenter. Bấm "Áp dụng bộ lọc cho file có sẵn" trên thanh công cụ, hoặc "Lọc Lại File Có Sẵn..." ở Bước 2.

1. Chọn nguồn: quét cả thư mục, hoặc chọn một hay nhiều file Excel. Có tùy chọn quét cả thư mục con và chỉ nhận file có tiền tố "PLM_".
2. Chọn dòng máy (hoặc để phần mềm tự nhận theo tên file/thư mục). Hộp thoại cho biết số quy tắc đang áp dụng (ví dụ "Có 77 quy tắc đang áp dụng cho [Virgo]").
3. Chọn cách ghi: "Tự động tạo bản sao lưu (.bak) trước khi cắt tỉa" và "Ghi đè trực tiếp lên tệp gốc" hoặc "Lưu thành tệp mới (*_filtered.xlsx)".
4. Bấm "Bắt Đầu Lọc BOM".

![Lọc file có sẵn: trước khi chạy](images/g06_standalone_before.png)

Hộp thoại có cảnh báo: chỉ xóa thêm được các dòng còn sót; dòng đã bị xóa khỏi file sẽ không tự khôi phục, muốn lấy lại cần dùng bản sao lưu hoặc tải lại từ Teamcenter.

Kết quả chạy trên bản sao file PLM của ví dụ:

> Hoàn tất (77 quy tắc): Ban đầu 1,890 dòng -> Còn lại 1,878 dòng (Đã cắt 12 dòng linh kiện con).
> Kết thúc: Đã xử lý 1/1 tệp (1 thành công, 0 lỗi). Tổng số dòng đã cắt tỉa: 12 dòng.

Phần mềm tạo bản sao lưu trước khi ghi đè (`PLM_110C3M3AK0_<ngày giờ>.xlsx`, nằm trong thư mục `backup_before_filter`).

![Lọc file có sẵn: sau khi chạy](images/g07_standalone_after.png)

### 4.12 Bước 3. Phụ trách nhập liệu (màn hình "Thành Viên Công Đoạn")

Phần này do từng kỹ sư làm trên máy của mình. Trong ví dụ, mỗi tài khoản được nhập lần lượt trên cùng một máy.

1. Mở file phân công của mình (nút "Mở file phân công..."). Màn hình hiện tên phụ trách, mã máy, phòng ban và trạng thái "CHƯA HOÀN THÀNH (Q2 TRỐNG)".
2. Nhập danh sách linh kiện ở tab "1. Linh kiện CTTT (A:N)" (nút "Thêm dòng" hoặc "Nhập từ Excel/CSV...").

![Bước 3: nhập linh kiện](images/e07_member_entered.png)

3. Bấm "Kiểm tra Sơ bộ (Self-Check PLM & R3)" để đối chiếu số lượng với PLM và R3.

![Bước 3: sau Self-Check](images/e08_member_selfcheck.png)

4. Bấm "Xác nhận Hoàn Thành (đóng dấu Q2 = OK)". Nếu còn linh kiện sai khác chưa điền "Giải thích", phần mềm hỏi có tiếp tục không. Sau khi đồng ý, phần mềm báo:

> Đã đóng dấu Q2 = OK và lưu file thành công!

![Bước 3: đã xác nhận](images/e09_member_submitted.png)

Trong ví dụ, số linh kiện của 9 phụ trách lần lượt là: Duong 312, Ly 481, Nhung 356, Hung 87, VietE 109, Tu 210, Thai 281, Chien 307, Thuong 55 (tổng 2198).

### 4.13 Theo dõi và tổng hợp (tab 3)

Tab 3 liệt kê từng phụ trách và trạng thái file. Bấm "Quét trạng thái ngay" để cập nhật.

Khi cả 9 file có Q2 = OK, thanh trạng thái hiện "Hoàn tất 100% (9/9 file đạt OK). Sẵn sàng tổng hợp dữ liệu." và nút "TỔNG HỢP DỮ LIỆU THÀNH VIÊN" sáng lên.

![Bước 3 (Trưởng nhóm): 9/9 OK](images/e11_step3_all_ok.png)

Bấm "TỔNG HỢP DỮ LIỆU THÀNH VIÊN". Kết quả trong ví dụ: 2198 dòng được tổng hợp. Từ v1.0.7, phần mềm chỉ coi một file là đã nộp khi ô Q2 = OK; file mới tạo từ mẫu, chưa nhập linh kiện, vẫn hiện "Chưa hoàn thành". Cột "Số LK" là số linh kiện thực tế (đếm theo Mã linh kiện).

### 4.14 Bước 4. Tạo file BOM tổng (tab 4)

Tab 4, mục "4.1 Khởi Tạo File So Sánh BOM Tổng", mã máy `110C3M3AK0`. Trước khi tạo, dòng trạng thái ghi "Chưa tạo file BOM tổng hợp."

![Bước 4: trước khi tạo](images/e13_step4_before.png)

Bấm "Tạo File BOM Tổng && Refresh Pivot Tables". Dòng trạng thái đổi thành "Đã tạo xong file: BOM_110C3M3AK0.xlsm (168 KB)".

![Bước 4: đã tạo file](images/e14_step4_generated.png)

File được lưu tại `Virgo\110C3M3AK0\BOM_110C3M3AK0.xlsm`. Các nút "Mở File BOM Tổng" và "Mở Thư Mục Máy" mở file và thư mục này.

File 168 KB này được tạo khi tắt kết nối Excel (COM), chưa mở bằng Excel nên chưa kiểm tra nội dung. Phần cập nhật BOM ở mục 4.15 dùng file BOM mẫu đầy đủ (có macro, 3,2 MB) và chạy với Excel thật.

### 4.15 Cập nhật BOM mới (PLM & R3) vào file BOM

Dùng khi có file PLM hoặc R3 mới và cần đưa vào file BOM đã có mà vẫn giữ các cột giải trình. Có hai chỗ bấm:

- Nút "Cập nhật BOM mới (PLM & R3)" trên thanh công cụ: mở hộp thoại riêng (mô tả bên dưới).
- Ba nút trong tab 4 mục "Cập nhật BOM mới": "Cập nhật PLM vào BOM", "Cập nhật R3 vào BOM", "Cập nhật cả PLM & R3". Ba nút này dùng cùng bộ xử lý với hộp thoại; chưa bấm riêng từng nút khi làm tài liệu.

**Các bước trong hộp thoại "Cập Nhật Dữ Liệu BOM Mới (PLM & SAP R3)":**

1. Chọn "Thư mục máy" (nút "Duyệt thư mục..."). Phần mềm tìm trong thư mục: file BOM so sánh, file `PLM*` mới và file `R3*` mới. Bấm "Quét lại" nếu vừa thêm file. Có thể chọn file riêng bằng nút "Chọn file...".
2. Giữ ô "Tự động chuyển file PLM/R3 cũ vào thư mục 'capnhat\old\'" nếu muốn lưu trữ file thô.
3. Bấm "7. Cập nhật PLM mới", "8. Cập nhật R3 mới" hoặc "Cập nhật cả PLM & R3".

![Hộp thoại cập nhật BOM, đã quét thư mục](images/g08_update_dialog_scanned.png)

Hộp thoại ghi rõ phần này tương ứng nút 7 và nút 8 của bản VBA trước đây, và tự kế thừa 3 cột: Cột O (Giải thích), Cột P (Người phụ trách), Cột Q (Quản lý check), đồng bộ lại Pivot Table.

Trong lần chạy thật với file mẫu (Excel thật, chạy cả PLM và R3), quá trình mất khoảng 52 đến 65 giây:

![Hộp thoại cập nhật BOM, sau khi chạy xong](images/g09_update_dialog_done.png)

Phần mềm báo kết quả:

![Thông báo cập nhật hoàn tất](images/g10_msg_update_done.png)

Những gì đã kiểm tra sau khi chạy:

| Mục | Trước | Sau |
|---|---|---|
| Kích thước file BOM | 3,245,301 byte | 3,974,239 byte |
| Số sheet | 9 | 11 (thêm `PLM_old_2` và `R3_old`, là bản lưu của dữ liệu trước khi cập nhật) |
| VBA macro (`vbaProject.bin`) | Có | Có |

- Trong thư mục `capnhat\old\` có 2 bản sao lưu BOM (`BOM_110C3M3AK0_backup_<ngày giờ>.xlsm`) và file PLM, R3 thô đã được chuyển vào.
- File BOM sau cập nhật (và sau bước lọc ở mục 4.16) được mở lại bằng Excel thật để kiểm tra: mở được, 11 sheet, có VBA project.
- Số "48 linh kiện kế thừa giải trình" là số do phần mềm báo trong lần chạy này.

Lưu ý về ví dụ: trên sheet CTTT, các cột so sánh R3 hiện `#N/A` cả ở file BOM mẫu ban đầu lẫn sau khi cập nhật. File R3 dùng ở đây không phải file xuất trực tiếp từ SAP, nên chưa kết luận được về cột này. Chưa chạy thử với file R3 xuất thật từ SAP.

### 4.16 Lọc dữ liệu bất đồng bộ (NG / #N/A)

Sau khi cập nhật, bấm "Lọc Bất Đồng Bộ (NG/#N/A)" trong tab 4 để chỉ giữ lại các dòng có NG hoặc #N/A, ẩn các dòng OK. Nút này làm việc trên file BOM của mã máy đang chọn.

Kết quả trong ví dụ (sau khi cập nhật ở mục 4.15):

![Thông báo lọc bất đồng bộ](images/g11_msg_filter_done.png)

Kích thước file BOM sau khi lọc: 3,398,897 byte.

Các sheet của file BOM sau khi cập nhật và lọc, chụp từ Excel thật (mở chế độ chỉ đọc):

Sheet CTTT (cột "Compare" cho biết OK hoặc NG so với PLM; các cột R3 và REV hiện `#N/A`/`#REF!` như đã nêu ở mục 4.15):

![Sheet CTTT](images/g12_excel_cttt.png)

Sheet PLM (BOM PLM đã cập nhật, cột "Giải thích" và "Phụ trách" được giữ lại):

![Sheet PLM](images/g13_excel_plm.png)

Sheet CTTT_Total (tổng số lượng theo mã linh kiện):

![Sheet CTTT_Total](images/g14_excel_cttt_total.png)

### 4.17 Gửi mail nhờ quản lý kiểm tra BOM

Bấm "Gửi Mail Nhờ Quản Lý Check BOM..." trong tab 4. Phần mềm mở cửa sổ xem trước thư gửi quản lý, kèm file BOM tổng làm tệp đính kèm, với các nút "Mở trong Microsoft Outlook" và "Gửi trực tiếp qua Outlook". Nút "Xem trước Email" xem nội dung 2 luồng thư (luồng 1 gửi thành viên nhắc 18 điểm kiểm tra; luồng 2 gửi quản lý báo cáo xác nhận) ở mục "4.3 Hệ Thống Gửi Email Outlook 2 Tầng".

![Thư gửi quản lý](images/g19_mail_manager.png)

Trong ảnh "Chế độ Thử nghiệm" đang bật. Chưa gửi thư thật khi làm tài liệu.

### 4.18 Mục 4.2 JIG và đánh giá 4M

Tab 4 mục "4.2 Quản lý Danh Mục Master List JIG & Đánh Giá 4M (KTSX)" có: chọn dòng máy JIG, nút "Nạp JIG từ Master", ô "Có phát sinh thay đổi 4M (Man, Machine, Material, Method)", đánh giá 4M (Đạt OK / Không đạt NG / Chờ xác nhận), ô kỹ sư KTSX xác nhận và ngày đánh giá (xem ảnh ở mục 4.14). Chưa chạy chức năng này khi làm tài liệu.

## 5. Tóm tắt những gì đã và chưa chạy khi làm tài liệu

| Tính năng | Tình trạng |
|---|---|
| Lập dự án, tạo thư mục, phân công (Bước 1) | Đã chạy thật |
| Nạp file từ ổ đĩa, lọc BOM Level 1..6 (Bước 2) | Đã chạy thật |
| Nhập liệu thành viên, Self-Check, xác nhận Q2 = OK (Bước 3) | Đã chạy thật cho 9 tài khoản |
| Tổng hợp dữ liệu thành viên (Bước 3) | Đã chạy thật (2198 dòng) |
| Tạo file BOM tổng (Bước 4) | Đã chạy, tắt Excel COM; chưa mở file kiểm tra |
| Cập nhật PLM và R3 vào BOM (Excel thật) | Đã chạy thật, mở lại file bằng Excel để kiểm tra |
| Lọc bất đồng bộ (NG/#N/A) | Đã chạy thật |
| Áp dụng bộ lọc cho file có sẵn | Đã chạy thật trên bản sao |
| Cấu hình bộ lọc, lọc trực quan, quét mã máy, quản lý model/nhân sự, cấu hình hệ thống, đổi giao diện | Đã mở và chụp; không thay đổi dữ liệu |
| Tải BOM tự động từ Teamcenter và SAP | Chưa chạy (cần tài khoản, mạng) |
| Quét kế hoạch PCD với file kế hoạch thật | Chưa chạy |
| Gửi email thật qua Outlook | Chưa chạy |
| Mục 4.2 JIG / 4M | Chưa chạy |
| Cập nhật phần mềm tự động (tải và cài) | Chưa chạy lại; hộp thoại được dựng để chụp |

## 6. Ghi chú

- Màn hình Thành Viên: sau Self-Check, cột "Kết Quả So Sánh" có thể hiện NG ở những dòng số lượng PLM và R3 bằng nhau. Chưa xác định nguyên nhân; nếu gặp, báo lại cho người phụ trách phần mềm.
- Giao diện Tối: các tab bước chưa chọn trong màn hình Trưởng Nhóm hiện nền sáng, chữ xám. Vẫn đọc được nhưng khác với phần còn lại của giao diện Tối.
- Với "Chế độ Thử nghiệm" bật, thư gửi đi bị chuyển tới người kiểm thử. Kiểm tra trong "Cấu hình hệ thống" → tab "Email Mẫu Form" trước khi dùng chính thức.
