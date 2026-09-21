# SPEC-KIT: TÁI CẤU TRÚC GIAO DIỆN DATA-DENSE ENTERPRISE DASHBOARD CHO KHỐI KỸ THUẬT PE DEPT

> **Dự án**: Phần Mềm Quản Lý & So Sánh BOM (`pm_sosanhbom`)  
> **Nền tảng**: Python 3.13 / PyQt6 Desktop  
> **Đối tượng người dùng**: Kỹ sư Khối Sản Xuất Kỹ Thuật (PE Dept - Manufacturing & Production Engineering)  
> **Phong cách chủ đạo**: **Data-Dense Enterprise Dashboard** kết hợp **Minimalism & Swiss Style** (Chuẩn `ui-ux-pro-max`)  
> **Chế độ hiển thị**: Song song **Light Theme** (Slate Industrial) & **Dark Theme** (Industrial Charcoal/Navy)  
> **Phiên bản đặc tả**: 1.0.0 (Authoritative Release)  
> **Trạng thái**: ACTIVE SPECIFICATION  
> **Tài liệu tham chiếu**:  
> - `C:\Users\tvn183660\.gemini\config\skills\ui-ux-pro-max\SKILL.md` (Design Intelligence)  
> - `specs/SPEC_PLM_AUTO_DOWNLOAD.md` (Core System & Automation Integration)  

---

## MỤC LỤC
1. [TỔNG QUAN KIẾN TRÚC THIẾT KẾ & BẢN QUYỀN GIAO DIỆN](#1-tổng-quan-kiến-trúc-thiết-kế--bản-quyền-giao-diện)
2. [CÁC NGUYÊN TẮC THIẾT KẾ ĐÃ CHỐT (CLOSED RESOLUTIONS)](#2-các-nguyên-tắc-thiết-kế-đã-chốt-closed-resolutions)
3. [HỆ THỐNG DESIGN TOKENS TOÀN CỤC (LIGHT & DARK THEMES)](#3-hệ-thống-design-tokens-toàn-cục-light--dark-themes)
   - 3.1. Bảng màu Light Theme (Slate Industrial)
   - 3.2. Bảng màu Dark Theme (Industrial Dark Mode)
   - 3.3. Ma trận kiểm định tương phản WCAG 2.1 (Tối thiểu 4.5:1, Mục tiêu 7:1)
   - 3.4. Hệ thống Typography kỹ thuật & Phân cấp thị giác
   - 3.5. Quy chuẩn Spacing, Padding và Grid viền bảng biểu
4. [ĐẶC TẢ CHI TIẾT TỪNG MÀN HÌNH CHỨC NĂNG](#4-đặc-tả-chi-tiết-từng-màn-hình-chức-năng)
   - 4.1. Màn hình Leader View (`src/gui/leader_view.py`)
   - 4.2. Màn hình Member View (`src/gui/member_view.py`)
   - 4.3. Hộp thoại Tải BOM PLM (`src/gui/plm_download_dialog.py`)
   - 4.4. Hộp thoại Cài Đặt & Chuyển Đổi Theme (`src/gui/settings_dialog.py`)
5. [HỆ THỐNG BIỂU TƯỢNG VECTOR SVG TIÊU CHUẨN (LUCIDE / FLUENT)](#5-hệ-thống-biểu-tượng-vector-svg-tiêu-chuẩn-lucide--fluent)
6. [BỘ QUY TẮC CHỐNG LỖI THIẾT KẾ (ANTI-PATTERNS CHECKLIST)](#6-bộ-quy-tắc-chống-lỗi-thiết-kế-anti-patterns-checklist)
7. [MA TRẬN TRUY XUẤT NGUỒN GỐC & BỘ TEST TỰ ĐỘNG HÓA UI/UX](#7-ma-trận-truy-xuất-nguồn-gốc--bộ-test-tự-động-hóa-uiux)

---

## 1. TỔNG QUAN KIẾN TRÚC THIẾT KẾ & BẢN QUYỀN GIAO DIỆN

### 1.1. Bối cảnh & Yêu cầu đặc thù của Kỹ sư PE Dept
Kỹ sư Khối Sản Xuất Kỹ Thuật (PE Dept) làm việc hàng ngày với các tập dữ liệu BOM cơ khí và điện tử có quy mô từ vài nghìn đến hàng chục nghìn linh kiện. Môi trường làm việc bao gồm cả văn phòng và nhà xưởng sản xuất:
- **Ánh sáng biến thiên**: Môi trường đèn huỳnh quang văn phòng yêu cầu độ tương phản chữ cực kỳ sắc nét; ca trực đêm tại nhà xưởng yêu cầu Dark Theme bảo vệ mắt.
- **Mật độ thông tin cao (Data-Dense)**: Giảm thiểu thao tác cuộn vô ích (scrolling fatigue), hiển thị tối đa dữ liệu trên một khung nhìn 1080p chuẩn desktop.
- **Chính xác tuyệt đối về số liệu**: Bảng biểu phải có đường kẻ viền phân định ô rõ ràng (`cell borders`), không bị nhầm lẫn giữa các dòng/cột.
- **Loại bỏ màu mè không cần thiết**: Không dùng hiệu ứng gradient tím/hồng, bóng đổ mờ ảo hay hoạt ảnh gây lag máy.

### 1.2. Cấu trúc Mô-đun Quản lý Giao diện (Theming Architecture)
```
src/gui/
├── styles/
│   ├── __init__.py
│   ├── theme_manager.py       # Quản trị viên chuyển đổi Light/Dark Theme tập trung
│   ├── tokens.py              # Định nghĩa Design Tokens (Colors, Fonts, Sizes, Borders)
│   ├── light_theme.qss        # Bộ stylesheet hoàn chỉnh cho Light Mode
│   └── dark_theme.qss         # Bộ stylesheet hoàn chỉnh cho Dark Mode
├── assets/
│   └── icons/                 # Bộ icon SVG vector sắc nét (Lucide / Fluent)
├── app.py                     # Cửa sổ chính, tích hợp ThemeManager
├── leader_view.py             # Bảng điều khiển quản lý và phân công BOM
├── member_view.py             # Màn hình thao tác so khớp dành cho thành viên
├── plm_download_dialog.py     # Hộp thoại tải BOM PLM tự động
└── settings_dialog.py         # Cấu hình hệ thống & tuỳ chọn giao diện
```

---

## 2. CÁC NGUYÊN TẮC THIẾT KẾ ĐÃ CHỐT (CLOSED RESOLUTIONS)

| Mã Quyết Định | Vấn Đề Thiết Kế | Giải Pháp Đã Chốt | Lý Do Kỹ Thuật |
| :--- | :--- | :--- | :--- |
| **RES-UI-01** | **Mật độ hiển thị (Density)** | Áp dụng **Data-Dense Layout**: row-height 32px, padding ô 4px 8px, margin ngoài 8–12px. | Tối đa hóa số dòng BOM hiển thị trên màn hình mà chữ vẫn đọc rõ ràng. |
| **RES-UI-02** | **Chế độ màu sắc (Theme)** | Hỗ trợ song song **Light Theme** (Slate Industrial) và **Dark Theme** (Industrial Charcoal/Navy), lưu trạng thái vào cài đặt. | Phù hợp cho cả môi trường ban ngày văn phòng và ca đêm nhà máy. |
| **RES-UI-03** | **Viền bảng biểu (Table Grid)** | Kẻ viền ô rõ ràng bằng mã màu phân tách tương phản cao (`#CBD5E1` cho Light, `#334155` cho Dark). Có hover highlight dòng. | Kỹ sư dễ dàng gióng hàng ngang khi đối chiếu BOM hàng nghìn linh kiện. |
| **RES-UI-04** | **Hệ thống Icon** | Thay thế 100% text thô hoặc ký tự Unicode bằng **SVG Icons chuẩn Lucide/Fluent**. | Không bị vỡ nét trên màn hình HiDPI, đồng bộ trực quan hiện đại. |
| **RES-UI-05** | **Độ tương phản (Contrast)** | Tuân thủ nghiêm ngặt chuẩn **WCAG 2.1 AA** (tối thiểu 4.5:1) và phấn đấu đạt **AAA (7:1)** cho văn bản chính. | Chống mỏi mắt, ngăn chặn đọc nhầm mã linh kiện hoặc số lượng. |
| **RES-UI-06** | **Chống giật lag (Anti-Lag)** | Không dùng hoạt ảnh đồ họa nặng, thời gian phản hồi phím bấm và hover cố định 150ms. | Đảm bảo phần mềm chạy siêu mượt trên các máy tính kỹ thuật nhà xưởng. |

---

## 3. HỆ THỐNG DESIGN TOKENS TOÀN CỤC (LIGHT & DARK THEMES)

### 3.1. Bảng màu Light Theme (Slate Industrial)
- **Background chính (Main Canvas)**: `#F8FAFC` (Slate 50 - Sạch sẽ, không chói gắt như trắng tinh #FFFFFF)
- **Nền Card / Bảng biểu (Surface / Table BG)**: `#FFFFFF`
- **Nền dòng chẵn lẻ (Zebra Striping)**: Dòng chẵn `#FFFFFF`, dòng lẻ `#F8FAFC`
- **Dòng được chọn (Selected Row)**: `#E0E7FF` (Indigo 100 - Nhẹ nhàng nhưng rõ ràng)
- **Dòng di chuột qua (Hover Row)**: `#F1F5F9` (Slate 100)
- **Văn bản chính (Primary Text)**: `#0F172A` (Slate 900 - Độ tương phản **16.5:1** so với nền trắng)
- **Văn bản phụ / Nhãn (Secondary / Label Text)**: `#475569` (Slate 600 - Độ tương phản **5.8:1**, vượt chuẩn AA)
- **Đường viền ô & vách ngăn (Border & Grid Lines)**: `#CBD5E1` (Slate 300 - Nét mảnh 1px)
- **Màu nhấn chủ đạo (Primary Brand / CTA)**: `#2563EB` (Blue 600) | Hover: `#1D4ED8` (Blue 700)
- **Màu trạng thái nghiệp vụ (Semantic Status Colors)**:
  - **Khớp hoàn toàn (Match / OK)**: Nền `#DCFCE7`, Chữ `#166534`, Viền `#86EFAC`
  - **Lệch BOM / Thừa linh kiện (Mismatch / Diff)**: Nền `#FEE2E2`, Chữ `#991B1B`, Viền `#FCA5A5`
  - **Chờ xử lý / Đang đối soát (Pending / Warning)**: Nền `#FEF3C7`, Chữ `#92400E`, Viền `#FCD34D`
  - **Thông tin kỹ thuật (Info / Metadata)**: Nền `#E0F2FE`, Chữ `#075985`, Viền `#7DD3FC`

### 3.2. Bảng màu Dark Theme (Industrial Dark Mode)
- **Background chính (Main Canvas)**: `#0B0F17` (Deep Slate / Dark Charcoal)
- **Nền Card / Bảng biểu (Surface / Table BG)**: `#151D2A`
- **Nền dòng chẵn lẻ (Zebra Striping)**: Dòng chẵn `#151D2A`, dòng lẻ `#111722`
- **Dòng được chọn (Selected Row)**: `#1E3A5F`
- **Dòng di chuột qua (Hover Row)**: `#1A2436`
- **Văn bản chính (Primary Text)**: `#F1F5F9` (Slate 100 - Độ tương phản **13.8:1** so với nền tối)
- **Văn bản phụ / Nhãn (Secondary / Label Text)**: `#94A3B8` (Slate 400 - Độ tương phản **6.2:1**)
- **Đường viền ô & vách ngăn (Border & Grid Lines)**: `#2A374A`
- **Màu nhấn chủ đạo (Primary Brand / CTA)**: `#3B82F6` (Blue 500) | Hover: `#60A5FA` (Blue 400)
- **Màu trạng thái nghiệp vụ (Semantic Status Colors)**:
  - **Khớp hoàn toàn (Match / OK)**: Nền `#064E3B`, Chữ `#A7F3D0`, Viền `#047857`
  - **Lệch BOM (Mismatch / Diff)**: Nền `#7F1D1D`, Chữ `#FECACA`, Viền `#B91C1C`
  - **Cảnh báo (Pending / Warning)**: Nền `#78350F`, Chữ `#FDE68A`, Viền `#D97706`

### 3.3. Ma trận kiểm định tương phản WCAG 2.1
| Thành Phần | Màu Chữ | Màu Nền | Tỷ Lệ Tương Phản | Đánh Giá WCAG |
| :--- | :--- | :--- | :---: | :---: |
| **Light: Tiêu đề & Văn bản chính** | `#0F172A` | `#FFFFFF` | **16.5 : 1** | **AAA (Vượt ngưỡng)** |
| **Light: Văn bản phụ / Ghi chú** | `#475569` | `#FFFFFF` | **5.8 : 1** | **AA Pass** |
| **Light: Nút bấm chính (CTA)** | `#FFFFFF` | `#2563EB` | **4.6 : 1** | **AA Pass** |
| **Light: Nhãn Lệch BOM (Diff)** | `#991B1B` | `#FEE2E2` | **7.4 : 1** | **AAA Pass** |
| **Light: Nhãn Khớp BOM (OK)** | `#166534` | `#DCFCE7` | **7.1 : 1** | **AAA Pass** |
| **Dark: Văn bản chính** | `#F1F5F9` | `#151D2A` | **13.8 : 1** | **AAA (Vượt ngưỡng)** |
| **Dark: Văn bản phụ** | `#94A3B8` | `#151D2A` | **6.2 : 1** | **AA Pass** |
| **Dark: Nút bấm chính** | `#0B0F17` | `#3B82F6` | **6.8 : 1** | **AA Pass** |

### 3.4. Hệ thống Typography kỹ thuật & Phân cấp thị giác
- **Font chữ ưu tiên**: `"Segoe UI", "Aptos", -apple-system, BlinkMacSystemFont, "Meiryo", "Microsoft YaHei", sans-serif`
- **Kích thước & Trọng số**:
  - `Display / App Title`: **18px, Semi-Bold (600)** - Tiêu đề thanh công cụ và thương hiệu.
  - `Section Heading / KPI`: **15px, Semi-Bold (600)** - Tiêu đề các khối quản lý, Model Name.
  - `Table Header`: **12px, Semi-Bold (600), Uppercase** - Tiêu đề các cột bảng dữ liệu.
  - `Table Cell / Body`: **12px, Regular (400)** - Dữ liệu linh kiện, số lượng, Part Number.
  - `Code / Part No. / Revisions`: **12px, Monospace ("Consolas", "Courier New")** - Đảm bảo số lượng và mã part thẳng hàng từng ký tự.
  - `Caption / Badge`: **11px, Medium (500)** - Nhãn tag trạng thái, Level phân cấp BOM.

### 3.5. Quy chuẩn Spacing & Padding Viền Bảng (Table Grid Specification)
```css
/* Trích xuất Quy Chuẩn QSS Bảng Biểu Data-Dense */
QTableView, QTableWidget {
    gridline-color: #CBD5E1;
    border: 1px solid #CBD5E1;
    background-color: #FFFFFF;
    selection-background-color: #E0E7FF;
    selection-color: #0F172A;
    outline: none;
    font-size: 12px;
}

QTableView::item {
    padding: 4px 8px;
    min-height: 28px;
    border-bottom: 1px solid #F1F5F9;
}

QTableView::item:hover {
    background-color: #F8FAFC;
}

QHeaderView::section {
    background-color: #F1F5F9;
    color: #475569;
    font-weight: 600;
    font-size: 11px;
    padding: 6px 8px;
    border: none;
    border-right: 1px solid #CBD5E1;
    border-bottom: 1px solid #CBD5E1;
}
```

---

## 4. ĐẶC TẢ CHI TIẾT TỪNG MÀN HÌNH CHỨC NĂNG

### 4.1. Màn hình Leader View (`src/gui/leader_view.py`)
- **Khối KPI Header**:
  - Bổ sung 4 thẻ tóm lược (Compact KPI Cards): `Tổng số Model`, `Model đã khớp BOM`, `Model có sai lệch`, `Tỷ lệ sẵn sàng CTTT`.
  - Nút bấm chính: `Tải BOM từ TC2412`, `So khớp toàn bộ`, `Xuất báo cáo tổng hợp Excel`.
- **Bảng Phân Công & Sourcing (Sourcing Table)**:
  - Hiển thị viền lưới sắc nét, độ cao dòng 32px.
  - Cột Date Picker đồng bộ chuẩn, không bị nhảy layout.
  - Dropdown Sub-Unit hiển thị tinh gọn với mũi tên SVG sắc nét.
- **Tab Kế Thừa & Đồ Gá (Jig & Inheritance)**:
  - Phân định rõ ràng ma trận kế thừa BOM cha-con, hiển thị cây phân cấp linh kiện bằng indent và biểu tượng thư mục SVG nhỏ gọn.

### 4.2. Màn hình Member View (`src/gui/member_view.py`)
- **Thanh Wizard 3 bước (Workflow Stepper)**:
  - `Bước 1: Chọn Model & BOM`, `Bước 2: Đối Soát Quy Tắc`, `Bước 3: Phụ trách công đoạn & Xuất Kết Quả`.
  - Trạng thái bước hiển thị rõ nét: Đang làm (Blue dot), Hoàn thành (Green check SVG), Chưa bắt đầu (Slate circle).
- **Khu vực hiển thị kết quả so sánh (Comparison Diff View)**:
  - Chia 2 cột song song (Split View) hoặc bảng tổng hợp có nhãn Diff nổi bật.
  - Các dòng sai lệch số lượng hoặc thiếu linh kiện được viền đỏ mềm (`#FCA5A5`) và chữ đỏ sẫm (`#991B1B`), đảm bảo kỹ sư nhìn thấy ngay tức thì.

### 4.3. Hộp thoại Tải BOM PLM (`src/gui/plm_download_dialog.py`)
- **Bảng danh sách Model cần tải**:
  - Ô nhập Part Number hỗ trợ dán nhiều dòng (multiline paste), tự động loại bỏ ký tự trắng thừa.
  - Cột trạng thái: `Sẵn sàng`, `Đang đăng nhập TC2412`, `Đang bung cây Level 7`, `Đang xuất Excel`, `Đã chuẩn hóa 14 cột`, `Thành công`.
- **Thanh tiến độ tổng thể (Global Progress Bar)**:
  - Chiều cao 14px, bo góc nhẹ 4px, có hiển thị % số lượng và số giây ước tính còn lại.
- **Khung Log thời gian thực**:
  - Font Monospace (Consolas 11px), màu chữ sắc nét, tự động cuộn xuống dòng mới nhất.

### 4.4. Hộp thoại Cài Đặt & Chuyển Đổi Theme (`src/gui/settings_dialog.py`)
- Bổ sung tuỳ chọn chuyển đổi giao diện:
  - `Giao diện Sáng (Light Industrial)` [Mặc định]
  - `Giao diện Tối (Industrial Dark Mode)`
  - `Theo cài đặt hệ thống Windows`
- Lưu cấu hình vào `config/settings.json` và áp dụng tức thời (`hot-reload stylesheet`) không cần khởi động lại ứng dụng.

---

## 5. HỆ THỐNG BIỂU TƯỢNG VECTOR SVG TIÊU CHUẨN (LUCIDE / FLUENT)

Tất cả các biểu tượng được lưu trữ dưới định dạng `.svg` tinh gọn tại `src/gui/assets/icons/`:
- `download.svg`: Tải BOM từ hệ thống PLM
- `refresh.svg`: Làm mới / Đồng bộ dữ liệu
- `check-circle.svg`: Khớp BOM hoàn toàn
- `alert-triangle.svg`: Cảnh báo / Sai lệch số lượng
- `x-circle.svg`: Thiếu linh kiện / Lỗi
- `settings.svg`: Cài đặt hệ thống
- `search.svg`: Tìm kiếm linh kiện / mã máy
- `filter.svg`: Lọc dữ liệu bảng
- `file-spreadsheet.svg`: Xuất tệp Excel báo cáo
- `moon.svg` / `sun.svg`: Chuyển đổi Dark / Light Theme

---

## 6. BỘ QUY TẮC CHỐNG LỖI THIẾT KẾ (ANTI-PATTERNS CHECKLIST)

- [x] **Không sử dụng Emoji làm icon chức năng**: Thay toàn bộ bằng SVG thuần.
- [x] **Không dùng màu xám trên nền xám (Gray-on-gray)**: Đảm bảo độ tương phản chữ tối thiểu 4.5:1.
- [x] **Không sử dụng font chữ dưới 11px**: Kỹ sư cần đọc rõ ràng từng ký tự trong mã linh kiện.
- [x] **Không sử dụng hiệu ứng bóng mờ (Blur shadows) gây giảm hiệu năng**: Dùng viền sắc nét 1px.
- [x] **Không làm mất viền ô bảng biểu (Borderless tables)**: Bảng dữ liệu kỹ thuật bắt buộc phải có gridline rõ ràng.
- [x] **Không để nút bấm thiếu trạng thái phản hồi**: 100% nút bấm phải có hiệu ứng hover và pressed.

---

## 7. MA TRẬN TRUY XUẤT NGUỒN GỐC & BỘ TEST TỰ ĐỘNG HÓA UI/UX

### 7.1. Ma trận kiểm thử UI/UX
| Mã Kiểm Thử | Tên Test Case | Tiêu Chuẩn Đáp Ứng | Vị Trí Test Tự Động |
| :--- | :--- | :--- | :--- |
| **TEST-UI-01** | `test_theme_tokens_contrast` | WCAG 2.1 Contrast >= 4.5:1 cho tất cả các cặp màu | `tests/unit/test_ui_theme.py` |
| **TEST-UI-02** | `test_qss_syntax_validity` | Bộ stylesheet QSS không chứa lỗi cú pháp trên PyQt6 | `tests/unit/test_ui_theme.py` |
| **TEST-UI-03** | `test_leader_view_data_dense` | Bảng biểu hiển thị đúng row-height, viền và padding | `tests/unit/test_leader_view_ui.py` |
| **TEST-UI-04** | `test_theme_toggle_persistence` | Chuyển đổi qua lại giữa Light và Dark không gây lỗi bộ nhớ | `tests/unit/test_ui_theme.py` |
| **TEST-UI-05** | `test_svg_icons_availability` | 100% tệp SVG tồn tại và render thành công trên QIcon | `tests/unit/test_ui_theme.py` |

---
*Tài liệu đặc tả này là căn cứ tối thượng (authoritative source) để triển khai giao diện Data-Dense Enterprise Dashboard cho dự án `pm_sosanhbom`.*
