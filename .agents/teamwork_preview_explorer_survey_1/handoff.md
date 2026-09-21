# BÁO CÁO KHẢO SÁT CHI TIẾT HIỆN TRẠNG MÃ NGUỒN GUI (EXPLORER 1)

> **Mã tác vụ**: Teamwork Explorer 1 - Survey Codebase GUI  
> **Thư mục làm việc**: `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_1`  
> **Tài liệu đặc tả đối chiếu**: `specs/SPEC_UI_UX_ENTERPRISE_DASHBOARD.md` & `specs/SPEC_PLM_AUTO_DOWNLOAD.md`  
> **Thời điểm hoàn thành**: 2026-09-21T08:50:00+07:00  
> **Phạm vi khảo sát**: Toàn bộ mô-đun giao diện tại `src/gui/` (`app.py`, `leader_view.py`, `member_view.py`, `plm_download_dialog.py`, `settings_dialog.py`, `update_dialog.py`) cùng các bộ test liên quan (`tests/unit/test_gui_and_reporting.py`, `tests/tier1_features/test_f22_leader_workspace.py`, `tests/tier1_features/test_f23_member_workspace.py`).

---

## 1. OBSERVATION (CÁC QUAN SÁT TRỰC TIẾP TỪ MÃ NGUỒN)

### 1.1. Bảng Tổng Hợp Hiện Trạng Các Tệp GUI Tại `src/gui/`

| Tệp Mã Nguồn | Kích Thước | Số Dòng | Lớp Chính (Classes) | Số Lượng Emoji | Lệnh setStyleSheet | Các Signal PyQt6 Chính |
| :--- | :---: | :---: | :--- | :---: | :---: | :--- |
| `src/gui/app.py` | 14.2 KB | 345 | `LogEmitter`, `QLogHandler`, `SSBOMMainWindow` | 5 | 5 | `new_record(str)` |
| `src/gui/leader_view.py` | 113.9 KB | 2562 | `ProjectStage`, `DepartmentGroup`, `StaffAssignment`, `MachineTarget`, `MemberSubmissionStatus`, `LeaderSessionState`, `BatchReconciliationWorker`, `EmailPreviewDialog`, `WizardStepHeader`, `Step1ProjectSetupWidget`, `Step2DataSourcingWidget`, `Step3TrackingConsolidationWidget`, `Step4ComparisonReportingWidget`, `LeaderWorkspaceView` | 13 (32 vị trí) | 29 | `batch_finished(object)`, `report_generated(str)`, `step_clicked(int)`, `step_completed(bool)`, `progress(int, str)` |
| `src/gui/member_view.py` | 76.3 KB | 1665 | `MemberWorkspaceView` | 13 (22 vị trí) | 25 | `submission_completed(dict)` |
| `src/gui/plm_download_dialog.py` | 34.6 KB | 782 | `BOMDownloadItem`, `UnifiedBOMDownloadWorker`, `PLMDownloadDialog` | 6 (9 vị trí) | 18 | `progress(int, str)`, `log_message(str)`, `finished(bool, str)` |
| `src/gui/settings_dialog.py` | 16.3 KB | 371 | `SettingsDialog` | 1 (1 vị trí) | 1 | `settings_saved(dict)` |
| `src/gui/update_dialog.py` | 9.5 KB | 249 | `UpdateWorker`, `UpdateDialog` | 4 (5 vị trí) | 10 | `progress_signal(int, int)`, `status_signal(str)`, `finished_signal(bool, str)` |

---

### 1.2. Khảo Sát Chi Tiết Từng Màn Hình Chức Năng

#### A. Main Application Window (`src/gui/app.py`)
- **Lớp điều phối**: `SSBOMMainWindow(QMainWindow)` (dòng 71–344).
- **Cấu trúc khung nhìn**:
  - `central_widget` chứa `QTabWidget` (`self.tabs`, dòng 110) phân quyền 2 vai trò:
    - Tab 0 (dòng 115): `"👔 Trưởng Nhóm"` -> `LeaderWorkspaceView(parent=self, base_dir=self.base_dir)`
    - Tab 1 (dòng 119): `"👷 Thành Viên Công Đoạn"` -> `MemberWorkspaceView(parent=self, base_dir=self.base_dir)`
- **Thanh Menu Bar (`self.menuBar()`)**:
  - Menu `"Hệ thống"` (dòng 135): `action_download_plm` (`"📥 Tải BOM Tự Động (PLM / SAP R3)..."`, Ctrl+D), `action_settings` (`"⚙ Cấu hình kết nối hệ thống..."`, Ctrl+,), `action_exit` (`"Đóng ứng dụng"`, Alt+F4).
  - Menu `"Không gian làm việc"` (dòng 155): `action_goto_leader`, `action_goto_member`.
  - Menu `"Trợ giúp"` (dòng 165): `action_check_updates` (`"🔄 Kiểm tra cập nhật phần mềm..."`), `action_about` (`"Về chương trình..."`).
- **Thanh công cụ (`QToolBar`, dòng 175–185)**: Quick-actions kết nối tới tải BOM, đổi workspace, cấu hình.
- **Thanh trạng thái (`QStatusBar`, dòng 186–216)**:
  - `self.lbl_status_msg` (thông báo chính).
  - `self.lbl_author`: `"Người viết: Bùi Đức Vinh - Phòng PTHT Chế Tạo"` (Hardcoded `#0078D4`).
  - `self.lbl_tc_indicator`: `"● TC24: Kết nối Web"` (Hardcoded `#10B981`).
  - `self.lbl_sap_indicator`: `"● SAP R3: Sẵn sàng"` (Hardcoded `#10B981`).
  - `self.lbl_version`: `"v1.0.0"` (Hardcoded `#6c757d`).
- **Khung nhật ký hoạt động (Log Dock, dòng 217–240)**:
  - `self.log_dock = QDockWidget("Nhật ký hoạt động", self)` ở đáy màn hình.
  - `self.log_text_edit = QPlainTextEdit()`, hardcoded style: `background-color: #1e1e1e; color: #d4d4d4; font-family: Consolas, monospace; font-size: 11px;`.
  - Kết nối qua `QLogHandler(logging.Handler)` phát tín hiệu `new_record` an toàn thread.

#### B. Leader Management Workspace (`src/gui/leader_view.py`)
- **Lớp điều phối**: `LeaderWorkspaceView(QWidget)` (dòng 2257–2562).
- **Quy trình tuần tự 4 bước (Sequential 4-Step Wizard)**:
  1. `WizardStepHeader` (dòng 393–466): Thanh breadcrumb 4 nút bấm (`QPushButton`). Nút hiện tại gắn nhãn `▶ {title}` (nền `#0078D4`), nút hoàn thành gắn `✓ {title}` (nền `#E8F5E9`), nút chưa mở gắn `🔒 {title}` (nền `#F5F5F5`).
  2. `QStackedWidget` (`self.step_stack`, dòng 2293) quản lý 4 bước:
     - **Bước 1: `Step1ProjectSetupWidget` (dòng 491–1109)**:
       - Cấu hình dự án: `model_combo` (Virgo, Libra2, Iris2024, Sirius2, Mebius, Polaris), `stage_combo` (MP/PP: ma1, DMT/PMT: maT), đường dẫn gốc `edit_base_dir`, nút `"📂 Mở thư mục"`.
       - Bảng mã máy: `machine_table` (`QTableWidget`, 4 cột: `STT`, `Mã Máy`, `Bỏ qua (X)`, `Ghi chú`). Nút: `"+ Thêm mã"`, `"- Xóa mã"`, `"📋 Nhập danh sách"`.
       - Bảng phân công nhân sự (Sheet Lichsu): `staff_table` (`QTableWidget`, 5 cột: `Áp dụng`, `Phụ trách công đoạn`, `Phòng Ban`, `Mã máy`, `Công Đoạn`). Nút lọc `"Cơ 1"`, `"Cơ 2"`, `"Cơ 3"`, nút `"Chọn tất cả"`, `"Tự động gán công đoạn"`.
       - Nút thực thi: `self.btn_create_folders` (`"📁 Khởi tạo Cây Thư Mục & Sinh Gói Nộp Thành Viên"`, nền `#0078D4`).
     - **Bước 2: `Step2DataSourcingWidget` (dòng 1110–1476)**:
       - Cấu hình ngày: `radio_common_date` (1 ngày chung cho DMT/PP) vs `radio_individual_date` (ngày riêng cho từng mã MP), `date_edit_common` (`QDateEdit`).
       - Bảng nguồn dữ liệu: `sourcing_table` (`QTableWidget`, 6 cột: `STT`, `Mã Máy`, `Ngày Hiệu Lực`, `File PLM TC24`, `File SAP R3 CS12`, `Bộ Lọc BOM Lv1..6`).
       - Nút hành động: `btn_download_plm` (`"📥 Tải BOM Tự Động..."`, `#0078D4`), `btn_import_files` (`"📂 Nạp Tệp Sẵn Có..."`, `#0D9488`), `btn_filter_bom` (`"⚡ Lọc BOM TC24 Level 1..6 & Sao Lưu"`, `#D97706`), `btn_refresh_sourcing` (`"🔄 Làm mới trạng thái"`).
     - **Bước 3: `Step3TrackingConsolidationWidget` (dòng 1477–1844)**:
       - Giám sát nộp bài: `combo_mach_filter`, `btn_scan_now` (`"🔄 Quét trạng thái ngay"`), `chk_auto_scan` (`QTimer` quét chu kỳ 15 giây).
       - Bảng trạng thái: `submission_table` (`QTableWidget`, 8 cột: `STT`, `Mã Máy`, `Công Đoạn`, `Phụ trách công đoạn`, `Trạng Thái Nộp`, `Số LK`, `MSI`, `Thời Gian Nộp`).
       - Cổng chặn Fail-Closed: `self.btn_consolidate` (`"⚡ TỔNG HỢP DỮ LIỆU THÀNH VIÊN"`). Nếu còn dù chỉ 1 thành viên chưa nộp (Q2 != 'OK') thì nút bị `setEnabled(False)`, nền xám `#94A3B8`. Khi 100% OK mới chuyển xanh `#10B981`.
     - **Bước 4: `Step4ComparisonReportingWidget` (dòng 1845–2256)**:
       - So sánh & tạo BOM: `combo_active_machine`, `btn_gen_report` (`"⚡ Tạo File BOM Tổng & Refresh Pivot Tables"`, `#10B981`), `btn_open_bom` (`"📊 Mở File BOM Tổng"`), `btn_open_folder` (`"📁 Mở Thư Mục Máy"`).
       - Master JIG & 4M: `combo_jig_model` (15 dòng máy), `btn_load_jig` (`"📥 Nạp JIG từ Master"`), `chk_4m_change`, nhóm Radio 4M (`radio_4m_ok`, `radio_4m_ng`, `radio_4m_pending`), `edit_ktsx_reviewer`, `date_4m_eval`.
       - Email 2 tầng qua Outlook: `mail_tabs` (`QTabWidget`):
         - Luồng 1 (dòng 1992): `"📨 Luồng 1 - Gửi Thành Viên (18 Điểm Check)"` + nút `"✉ Gửi trực tiếp qua Outlook"`.
         - Luồng 2 (dòng 2028): `"👔 Luồng 2 - Gửi Quản Lý Báo Cáo Xác Nhận"` + nút `"✉ Gửi trực tiếp qua Outlook"`.

#### C. Member Input Workspace (`src/gui/member_view.py`)
- **Lớp điều phối**: `MemberWorkspaceView(QWidget)` (dòng 106–1664).
- **Banner thông tin phân công (Zone 1, dòng 150–233)**:
  - Row A: `btn_open_assignment` (`"📂 Mở gói phân công..."`), `btn_select_machine_dir` (`"📁 Chọn thư mục máy..."`), `lbl_assignment_path`, `btn_load_refs` (`"📂 Nạp PLM & R3 đối chiếu..."`), `lbl_ref_status`.
  - Row B: Tự động hiển thị huy hiệu: `lbl_engineer_name`, `lbl_machine_code`, `lbl_department`, `lbl_submission_seal` (`"⏳ CHƯA NỘP (Q2 TRỐNG)"` / `"✅ ĐÃ NỘP BÀI (Q2 = OK)"`).
  - Row C: `sub_unit_combo` (10 sub-units), `author_edit` (tích hợp `QCompleter` 38 kỹ sư Cơ 1, 2, 3), `model_combo`.
- **Hệ thống 3 bảng nhập liệu (Zone 2, dòng 238–471)**:
  - Tab 1: `"📋 1. Linh kiện CTTT (A:N)"` -> `cttt_table` (`QTableWidget`, **15 cột**: `Trang CTTT`, `Mã Linh Kiện`, `Tên Linh Kiện`, `Số Lượng`, `SL PLM (Tham chiếu)`, `SL R3 (Tham chiếu)`, `Giải Thích Sai Khác`, `Kết Quả So Sánh`, `Unit / Công đoạn`, `Người Phụ Trách`, `So Sánh CTTT vs PLM`, `Rev PLM`, `So Sánh CTTT vs R3`, `Rev R3`, `So Sánh Rev PLM vs R3`).
    - Nút: `btn_add_row` (`"➕ Thêm dòng"`), `btn_remove_row` (`"➖ Xóa dòng"`), `btn_import_excel` (`"📥 Nhập từ Excel/CSV..."`), `btn_paste_clipboard` (`"📋 Dán Clipboard"`), `btn_clear_table` (`"🗑️ Xóa hết"`).
  - Tab 2: `"🏷️ 2. Quản lý MSI (35 Unit)"` -> `msi_table` (`QTableWidget`, **12 cột**, 35 dòng định danh chuẩn hóa Unit & Thân máy HONTAI).
    - Quick-Edit panel: `msi_barcode_edit`, `msi_unit_code_edit`, `msi_unit_name_edit`, `msi_code_edit` (3 ký tự MSI), `msi_service_edit`, `msi_abs_combo`, `msi_volt_combo`, `msi_label_type_combo`, `lbl_msi_badge`, `btn_apply_msi_edit` (`"💾 Cập nhật vào bảng MSI"`).
  - Tab 3: `"🏷️ 3. Nhãn LCP 7980 / 7990"` -> `label_table` (`QTableWidget`, **12 cột**).
    - Quick spec fields: `label_spec_combo`, `label_code_edit`, `label_pos_edit`, `label_check_combo`.
- **Thanh kiểm tra sơ bộ & nộp bài (Zone 3 & 4, dòng 476–512)**:
  - `btn_self_check` (`"🔍 Kiểm tra Sơ bộ (Self-Check PLM & R3)"`, `#0078D4`) -> gọi `run_preliminary_self_check()` so sánh tại chỗ CTTT với BOM PLM & R3 trong thư mục máy.
  - `lbl_check_summary`: Hiển thị số lượng `"Kết quả: X OK, Y NG / Tổng Z linh kiện"`.
  - `btn_unlock` (`"🔓 Hủy nộp / Mở khóa"`, `#6c757d`) -> xóa dấu Q2="OK".
  - `btn_submit` (`"🚀 Xác nhận Nộp (Đóng dấu Q2 = OK)"`, `#10B981`) -> ghi Q2="OK", khóa giao diện, phát signal `submission_completed`.

#### D. PLM / SAP R3 Automated Downloader (`src/gui/plm_download_dialog.py`)
- **Lớp điều phối**: `PLMDownloadDialog(QDialog)` (dòng 330–782).
- **Kiến trúc luồng nền**: `UnifiedBOMDownloadWorker(QObject)` chạy trong `QThread` (dòng 130–328). Tự động điều phối Selenium Headless (TC2412) và SAP GUI Scripting (`saplogon.exe` CS12).
- **Layout 3 bước dành cho non-tech**:
  - Bước 1: `txt_parts` (`QPlainTextEdit`, font Consolas 10pt) hỗ trợ dán nhiều dòng. Nút `"📋 Dán từ Clipboard"`, `"🧹 Xóa hết"`.
  - Bước 2: `tab_config` (`QTabWidget`):
    - Tab Teamcenter (TC24): Tài khoản `vn_pe01..04`, url, headless, timeout.
    - Tab SAP R3: Plant 2200, BOM Usage pp01, Alt 01, `date_common`, `chk_custom_dates` (`"⚙️ Nhập ngày hiệu lực riêng..."`), `date_table` (3 cột).
    - Thư mục lưu đích: `edit_dest`, `btn_browse` (`"📂 Duyệt..."`).
  - Bước 3: Nút tải:
    - `btn_both` (`"⚡ Tải đồng thời cả PLM & SAP R3"`, `#0d6efd`)
    - `btn_plm_only` (`"📥 Chỉ tải BOM PLM (TC24)"`, `#0284c7`)
    - `btn_sap_only` (`"📥 Chỉ tải BOM SAP R3"`, `#059669`)
  - Tiến trình & Log: `progress_bar` (độ cao hiện tại 20px, `#0d6efd`), `log_box` (`QTextEdit`, Consolas 9pt, nền `#f8f9fa`).

#### E. System Configuration Dialog (`src/gui/settings_dialog.py`)
- **Lớp điều phối**: `SettingsDialog(QDialog)` (dòng 65–371).
- **Tab cấu hình (`tab_widget`)**:
  - Tab 1: `"Teamcenter TC24"`: URL, user, pass (EchoMode.Password), browser (edge/chrome), headless checkbox, timeout spinbox, nút kiểm tra kết nối `btn_test_tc`.
  - Tab 2: `"SAP R3"`: System ID (P1J), Client (800), Plant (2200), Usage (pp01), Alt (01), đường dẫn `saplogon.exe`, nút kiểm tra kết nối `btn_test_sap`.
  - Tab 3: `"Đường dẫn & Dữ liệu"`: `base_dir_edit`, `fs_path_edit` (Master FIX_SERIAL), `reports_dir_edit`, `default_model_combo`.
- **Thao tác lưu/hủy**: `btn_reset` (`"Khôi phục mặc định"`), `btn_save` (`"💾 Lưu cấu hình"`, `#10B981`), `btn_cancel`.
- **Lưu cấu hình**: Lưu cấu trúc JSON bền vững tại `config/settings.json`.

---

### 1.3. Khảo Sát Chi Tiết Các Bảng Biểu (Tables) Trong Toàn Bộ Hệ Thống

| Tệp | Tên Biến Bảng | Lớp | Số Cột | Tên Các Cột | Chiều Cao Dòng Hiện Tại | Đường Viền / Gridline Hiện Tại |
| :--- | :--- | :--- | :---: | :--- | :---: | :---: |
| `leader_view.py` | `self.machine_table` | `QTableWidget` | 4 | STT, Mã Máy, Bỏ qua (X), Ghi chú | Mặc định (~24px) | Mặc định Qt (mờ, không rõ 1px) |
| `leader_view.py` | `self.staff_table` | `QTableWidget` | 5 | Áp dụng, Phụ trách công đoạn, Phòng Ban, Mã máy, Công Đoạn | Mặc định (~26px do có ComboBox) | Mặc định Qt |
| `leader_view.py` | `self.sourcing_table` | `QTableWidget` | 6 | STT, Mã Máy, Ngày Hiệu Lực, File PLM TC24, File SAP R3 CS12, Bộ Lọc BOM Lv1..6 | Mặc định (~24px) | Mặc định Qt |
| `leader_view.py` | `self.submission_table` | `QTableWidget` | 8 | STT, Mã Máy, Công Đoạn, Phụ trách công đoạn, Trạng Thái Nộp, Số LK, MSI, Thời Gian Nộp | Mặc định (~24px) | Mặc định Qt (`alternatingRowColors=True`) |
| `member_view.py` | `self.cttt_table` | `QTableWidget` | 15 | Trang CTTT, Mã Linh Kiện, Tên Linh Kiện, Số Lượng, SL PLM, SL R3, Giải Thích, Kết Quả, Unit, Người Phụ Trách, CTTT vs PLM, Rev PLM, CTTT vs R3, Rev R3, Rev PLM vs R3 | Mặc định (~24px) | Mặc định Qt (`alternatingRowColors=True`) |
| `member_view.py` | `self.msi_table` | `QTableWidget` | 12 | Mã LK Barcode, Mã UNIT, Tên UNIT, 3 Ký tự MSI, SEVICE, ABS, Trang CTTT, Điện áp, Loại Label, Tên phụ trách, Kết quả so sánh, Ghi chú | Mặc định (~24px) | Mặc định Qt (`alternatingRowColors=True`) |
| `member_view.py` | `self.label_table` | `QTableWidget` | 12 | Công đoạn (7980), Trang CTTT (7980), Mã label (7980), Tên label (7980), Số lượng (7980), Phụ trách (7980), Công đoạn (7990), Trang CTTT (7990), Mã label (7990), Tên label (7990), Số lượng (7990), Phụ trách (7990) | Mặc định (~24px) | Mặc định Qt (`alternatingRowColors=True`) |
| `plm_download_dialog.py` | `self.date_table` | `QTableWidget` | 3 | STT, Mã Máy, Ngày Hiệu Lực (SAP R3 CS12) | Mặc định (~24px) | Mặc định Qt |

---

### 1.4. Phát Hiện Lỗ Hổng & Điểm Buộc Nghiệp Vụ Tử Huyệt (Forensic Findings)

#### Phát hiện 1: Lỗi Phụ Thuộc Chuỗi Trạng Thái Có Ký Tự Emoji Checkmark `"✓"`
- **Vị trí**: `src/gui/leader_view.py`, dòng 1345, 1349, 1367, 1371.
- **Đoạn mã hiện tại**:
  ```python
  # Dòng 1345:
  plm_status = f"✓ Sẵn sàng ({sz_kb} KB)"
  # Dòng 1349:
  if "✓" in plm_status:
      item_plm.setBackground(QColor(f"#{COLOR_GREEN_FILL_HEX}"))
      ...
  else:
      ...
      all_ready = False  # <--- NGUY HIỂM!

  # Dòng 1367:
  r3_status = f"✓ Sẵn sàng ({sz_kb} KB)"
  # Dòng 1371:
  if "✓" in r3_status:
      ...
  else:
      all_ready = False  # <--- NGUY HIỂM!
  ```
- **Rủi ro**: Nếu agent xóa ký tự emoji `✓` khỏi chuỗi hiển thị mà không sửa điều kiện `if "✓" in ...`, biến `all_ready` sẽ vĩnh viễn bị gán `False`, làm vô hiệu hóa luồng nạp và lọc BOM tiếp theo!
- **Giải pháp**: Điều kiện phải kiểm tra trạng thái logic `if m.plm_file is not None:` và `if m.r3_file is not None:` thay vì kiểm tra chuỗi text hiển thị.

#### Phát hiện 2: Ràng Buộc Khớp Chuỗi Trạng Thái Bài Nộp trong Unit Test `test_f22`
- **Vị trí**: `src/gui/leader_view.py`, dòng 1762–1768:
  ```python
  def get_sub_unit_statuses(self) -> dict[str, str]:
      mapping: dict[str, str] = {}
      for sub in self.state.submissions:
          key = sub.sub_unit if sub.sub_unit != "CTTT" else sub.engineer_name
          mapping[key] = "Đã nộp" if sub.is_submitted_ok else "Chưa nộp"
      return mapping
  ```
- **Ràng buộc**: `tests/tier1_features/test_f22_leader_workspace.py` (dòng 78) kiểm tra trực tiếp:
  `assert mapping.get("LSU") == "Đã nộp"`
- **Yêu cầu**: Tuyệt đối không thay đổi giá trị trả về của `get_sub_unit_statuses()` (vẫn giữ nguyên `"Đã nộp"` và `"Chưa nộp"`).

#### Phát hiện 3: Bộ Thuộc Tính & Phương Thức Ủy Quyền (Delegations) Bắt Buộc Của `LeaderWorkspaceView`
- Trong `leader_view.py` (dòng 2390–2420), lớp `LeaderWorkspaceView` cung cấp các property ủy quyền xuống các Step con:
  - `view.model_combo -> self.step1_widget.model_combo`
  - `view.stage_combo -> self.step1_widget.stage_combo`
  - `view.submission_table -> self.step3_widget.submission_table`
  - `view.create_project_folder_structure()`
  - `view.scan_member_submissions()`
  - `view.get_sub_unit_statuses()`
  - `view._open_plm_download_dialog()`
  - `view.trigger_batch_reconciliation()`
- Toàn bộ test trong `test_gui_and_reporting.py` và `test_f22_leader_workspace.py` truy cập qua các interface này. Việc tái cấu trúc giao diện **bắt buộc phải giữ nguyên vẹn 100%** các delegation này.

#### Phát hiện 4: Bộ Phương Thức Nhập Liệu Của `MemberWorkspaceView`
- `MemberWorkspaceView` được gọi trực tiếp bởi `tests/tier1_features/test_f23_member_workspace.py` và `test_gui_and_reporting.py`:
  - `view.clear_cttt_table()`
  - `view.add_cttt_row(page, part_code, part_name, quantity, explanation)`
  - `view.cttt_table`
  - `view.set_reference_data(plm_df, r3_df)`
  - `view.run_preliminary_self_check()` (trả về dict có key `"status"`, `"ok_count"`, `"ng_count"`, `"total"`)
  - `view.label_spec_combo`
  - `view.label_code_edit`
  - `view.get_cttt_table_data()`
  - `view.submit_data()`
  - `view.unlock_submission()`
  - `view.author_edit`, `view.sub_unit_combo`, `view.model_combo`
  - Signal `submission_completed = pyqtSignal(dict)`

---

## 2. LOGIC CHAIN (CHỖI SUY LUẬN TỪ QUAN SÁT ĐẾN GIẢI PHÁP NÂNG CẤP)

```
[QUAN SÁT 1: 50+ Emoji nằm rải rác]
  ├── Chuỗi text nút bấm chứa: ⚡, 📥, 📁, 📂, 📋, 🔄, 🚀, 💾, 🔍, 🏷️...
  └── Leader View kiểm tra logic: if "✓" in plm_status (L1349)
        │
        ▼ (Suy luận bước 1)
        Phải phân tách triệt để: Icon hiển thị (dùng SVG thuần via QIcon) KHÔNG ĐƯỢC lẫn vào chuỗi văn bản logic.
        Sửa logic L1349/L1371 sang kiểm tra biến boolean/đối tượng thay vì chuỗi.

[QUAN SÁT 2: Stylesheets phân tán qua setStyleSheet inline]
  ├── Hơn 80 lệnh setStyleSheet cục bộ trên từng widget với mã màu cũ (#0078D4, #10B981, #FFC7CE, #C6EFCE).
  └── Chưa có thư mục src/gui/styles/ và src/gui/assets/icons/ theo SPEC_UI_UX_ENTERPRISE_DASHBOARD.
        │
        ▼ (Suy luận bước 2)
        Cần thiết lập hệ thống Stylesheet tập trung:
        - tokens.py: Định nghĩa bảng mã màu Slate Industrial (Light) & Dark Charcoal (Dark).
        - light_theme.qss & dark_theme.qss: Quy tắc styling tập trung cho toàn bộ QWidget, QTableView, QPushButton...
        - theme_manager.py: Quản lý switch theme thời gian thực (Hot-reload), lưu vào settings.json.
        - Thay thế toàn bộ inline stylesheet hardcoded bằng class selector hoặc QSS động.

[QUAN SÁT 3: Thiếu các thành phần cốt lõi của phong cách Data-Dense]
  ├── Leader View thiếu: 4 thẻ KPI Header Cards (Tổng Model, Model khớp BOM, Model lệch, Sẵn sàng CTTT).
  ├── Member View thiếu: Stepper 3 bước (Workflow Stepper) và Khung hiển thị Diff View chuyên biệt.
  ├── Settings Dialog thiếu: Bộ chọn Theme (Light / Dark / System) và cơ chế lưu theme.
  └── Toàn bộ 8 bảng biểu đang có chiều cao dòng mặc định (~24px), viền mờ, không đạt chuẩn 32px viền sắc nét.
        │
        ▼ (Suy luận bước 3)
        Triển khai bổ sung các thành phần UI mới mà KHÔNG làm thay đổi cấu trúc widget phân tầng hiện tại:
        - Thêm QFrame KPI Cards phía trên WizardStepHeader trong Leader Workspace.
        - Thêm QFrame Stepper 3 bước phía trên Zone 1 trong Member Workspace.
        - Bổ sung Diff View tab/panel hiển thị rõ ràng các dòng sai lệch với viền #FCA5A5 và chữ #991B1B.
        - Thêm Tab Giao diện vào SettingsDialog.
        - Đặt QHeaderView::section và QTableView::item min-height: 28px, padding: 4px 8px, row-height 32px trong QSS.

[QUAN SÁT 4: Ma trận kiểm thử hồi quy nghiêm ngặt]
  ├── tests/unit/test_gui_and_reporting.py (30 bài test đang PASS 100%).
  └── tests/tier1_features/test_f22_leader_workspace.py & test_f23 (10 bài test đang PASS 100%).
        │
        ▼ (Suy luận bước 4)
        Mọi cải tiến giao diện phải tuân thủ nghiêm ngặt nguyên tắc Zero Regression:
        - Không đổi tên thuộc tính hoặc method công khai.
        - Không thay đổi kiểu dữ liệu hoặc tham số của các signal/slot.
        - Đảm bảo các test case kiểm tra thuộc tính widget vẫn chạy thành công trong môi trường headless QT_QPA_PLATFORM=offscreen.
```

---

## 3. CAVEATS (CÁC GIỚI HẠN & GIẢ ĐỊNH)

1. **Giả định về hiển thị Icon SVG**: Trong môi trường Windows và PyQt6, việc nạp `.svg` vào `QIcon` yêu cầu plugin `qsvg` của PyQt6 (mặc định đi kèm `PyQt6-Qt6`). Đã kiểm tra môi trường Python 3.13 của dự án có đầy đủ thư viện PyQt6.
2. **Quy tắc bảo vệ mã nguồn cũ**: Explorer 1 hoạt động ở chế độ **Read-only Investigation**, tuyệt đối không tự ý sửa đổi mã nguồn sản xuất tại `src/gui/`. Toàn bộ đề xuất được ghi nhận thành đặc tả kỹ thuật và file bản mẫu để Implementer thực thi.
3. **Phạm vi hiển thị đa ngôn ngữ (i18n)**: Mặc dù `src/ui/i18n.py` hỗ trợ 3 ngôn ngữ (vi, ja, zh), giao diện `src/gui/` hiện tại đang sử dụng tiếng Việt làm ngôn ngữ hiển thị mặc định. Việc refactor sang Data-Dense phải giữ các chuỗi tiếng Việt rõ ràng, chuyên nghiệp cho khối kỹ thuật PE Dept.

---

## 4. CONCLUSION & KẾ HOẠCH HÀNH ĐỘNG CHI TIẾT (ACTIONABLE ROADMAP)

### 4.1. Kiến Trúc Tệp Cần Khởi Tạo & Nâng Cấp

```
src/gui/
├── styles/                          <-- [TẠO MỚI]
│   ├── __init__.py                  # Khai báo gói styles
│   ├── tokens.py                    # Design Tokens chuẩn: Light & Dark palette, spacing, typography
│   ├── theme_manager.py             # Quản lý áp dụng Theme, Hot-reload, tích hợp settings.json
│   ├── light_theme.qss              # Stylesheet Light Industrial (WCAG AAA contrast, viền #CBD5E1)
│   └── dark_theme.qss               # Stylesheet Industrial Dark Mode (viền #2A374A)
├── assets/icons/                    <-- [TẠO MỚI]
│   ├── download.svg                 # Thay thế 📥
│   ├── refresh.svg                  # Thay thế 🔄
│   ├── check-circle.svg             # Thay thế ✓, ✅
│   ├── alert-triangle.svg           # Thay thế ⚠️, ⏳
│   ├── x-circle.svg                 # Thay thế ❌, ✗
│   ├── settings.svg                 # Thay thế ⚙
│   ├── folder.svg                   # Thay thế 📁, 📂
│   ├── clipboard.svg                # Thay thế 📋
│   ├── plus.svg                     # Thay thế ➕
│   ├── minus.svg                    # Thay thế ➖
│   ├── trash.svg                    # Thay thế 🗑️, 🧹
│   ├── search.svg                   # Thay thế 🔍
│   ├── lock.svg                     # Thay thế 🔒
│   ├── unlock.svg                   # Thay thế 🔓
│   ├── mail.svg                     # Thay thế ✉, 📨
│   ├── send.svg                     # Gửi thư
│   ├── file-spreadsheet.svg         # Thay thế 📊
│   ├── sun.svg / moon.svg           # Icon chuyển Theme
│   └── zap.svg                      # Thay thế ⚡
├── app.py                           <-- [NÂNG CẤP] Tích hợp ThemeManager, xóa emoji trên tab/menu/toolbar, thêm nút toggle theme
├── leader_view.py                   <-- [NÂNG CẤP] Thêm 4 thẻ KPI Cards, bảng 32px, xóa emoji, sửa logic L1349/L1371
├── member_view.py                   <-- [NÂNG CẤP] Thêm Stepper 3 bước, bảng 32px, Diff View màu chuẩn WCAG, xóa emoji
├── plm_download_dialog.py           <-- [NÂNG CẤP] Thanh tiến độ 14px, log Consolas 11px, xóa emoji
├── settings_dialog.py               <-- [NÂNG CẤP] Thêm Tab Giao diện (Theme), hot-reload không cần restart
└── update_dialog.py                 <-- [NÂNG CẤP] Xóa emoji, áp dụng tokens
```

---

### 4.2. Chi Tiết Các Thay Đổi Mã Nguồn Cụ Thể (Before -> After Mapping)

#### A. Tại `src/gui/leader_view.py`:
1. **Sửa Logic Checkmark Tử Huyệt (L1349 & L1371)**:
   - *Trước*: `if "✓" in plm_status:` và `if "✓" in r3_status:`
   - *Sau*: `if m.plm_file is not None and m.plm_file.exists():` và `if m.r3_file is not None and m.r3_file.exists():`
2. **Bổ sung Khối KPI Header Cards**:
   - Vị trí: Đặt ngay trên `self.nav_header = WizardStepHeader(self)` trong `LeaderWorkspaceView._init_ui`.
   - Cấu trúc: `QFrame` bo góc 6px, chia 4 cột hiển thị:
     - Card 1: `Tổng số Model / Mã máy` (số lượng từ `self.state.machines`)
     - Card 2: `Model đã nạp đủ BOM` (số model có cả PLM & R3)
     - Card 3: `Tiến độ nộp bài CTTT` (tỷ lệ bài nộp OK / tổng bài)
     - Card 4: `Trạng thái đối soát` (Hoàn tất / Đang chờ)
   - 3 Quick Actions: Nút `Tải BOM TC24`, `Quét nộp bài`, `Xuất Excel`.
3. **Quy chuẩn chiều cao bảng 32px & viền sắc nét**:
   - Bổ sung hàm thiết lập thuộc tính chuẩn cho `machine_table`, `staff_table`, `sourcing_table`, `submission_table`:
     ```python
     table.verticalHeader().setDefaultSectionSize(32)
     table.verticalHeader().setMinimumSectionSize(28)
     table.setShowGrid(True)
     ```

#### B. Tại `src/gui/member_view.py`:
1. **Bổ sung Workflow Stepper 3 Bước**:
   - Vị trí: Đặt trên `Zone 1: Banner Thông tin Phân công`.
   - 3 bước:
     - `Bước 1: Chọn Model & BOM` (Sáng khi mở gói phân công)
     - `Bước 2: Đối Soát Quy Tắc` (Sáng khi chỉnh sửa CTTT/MSI và Self-Check)
     - `Bước 3: Phụ trách công đoạn & Xuất Kết Quả` (Sáng khi nộp bài Q2 = OK)
   - Trạng thái trực quan: Icon SVG tròn (Đang thực hiện: Xanh dương; Hoàn tất: Xanh lá check; Chưa làm: Xám nhạt).
2. **Khung Hiển Thị So Sánh Sai Lệch (Diff View)**:
   - Nâng cấp `_style_table_row` sang mã màu WCAG AAA chuẩn:
     - Khớp hoàn toàn: Nền `#DCFCE7`, Chữ `#166534` (Light) / Nền `#064E3B`, Chữ `#A7F3D0` (Dark).
     - Lệch BOM: Nền `#FEE2E2`, Chữ `#991B1B`, Viền ô `#FCA5A5` (Light) / Nền `#7F1D1D`, Chữ `#FECACA` (Dark).
3. **Chuẩn hóa bảng biểu**:
   - Đặt `self.cttt_table.verticalHeader().setDefaultSectionSize(32)`.
   - Đặt `self.msi_table.verticalHeader().setDefaultSectionSize(32)`.
   - Đặt `self.label_table.verticalHeader().setDefaultSectionSize(32)`.

#### C. Tại `src/gui/plm_download_dialog.py`:
1. **Thanh tiến độ tinh tế 14px**:
   - Thay đổi stylesheet và kích thước: `self.progress_bar.setFixedHeight(14)`, định dạng hiển thị % số lượng và số giây ước tính.
2. **Khung Log Consolas 11px**:
   - Cập nhật font: `self.log_box.setFont(QFont("Consolas", 10))` (tương đương 11px sắc nét).
   - Tự động cuộn: `self.log_box.moveCursor(QTextCursor.MoveOperation.End)`.

#### D. Tại `src/gui/settings_dialog.py`:
1. **Thêm Tab Giao diện (Theme Selection)**:
   - Thêm Tab 4 `"Giao diện & Hiển thị"` vào `tab_widget`.
   - Cung cấp `QComboBox` hoặc 3 `QRadioButton`:
     - `Giao diện Sáng (Light Industrial)` [Mặc định]
     - `Giao diện Tối (Industrial Dark Mode)`
     - `Theo cài đặt hệ thống Windows`
   - Gọi `ThemeManager.set_theme(...)` ngay khi người dùng thay đổi lựa chọn (live hot-reload).
   - Lưu cấu hình vào `settings["theme"] = "light" | "dark" | "system"`.

---

## 5. VERIFICATION METHOD (PHƯƠNG PHÁP XÁC THỰC ĐỘC LẬP)

Để xác thực độc lập rằng các phát hiện và phân tích khảo sát là chính xác, đồng thời đảm bảo mã nguồn mới không gây hồi quy (Zero Regression), chạy các lệnh kiểm thử tự động sau:

### 5.1. Kiểm thử toàn bộ các bài test GUI hiện có
```powershell
# Chạy bộ test GUI và báo cáo tổng hợp (30 passed hiện tại)
pytest tests/unit/test_gui_and_reporting.py -v

# Chạy bộ test tính năng F22 (Leader Workspace) và F23 (Member Workspace) (10 passed hiện tại)
pytest tests/tier1_features/test_f22_leader_workspace.py tests/tier1_features/test_f23_member_workspace.py -v
```
*Điều kiện đạt*: 100% (40/40) bài kiểm thử đạt kết quả **PASSED**.

### 5.2. Kiểm thử hợp đồng giao diện mới sau khi triển khai
```powershell
# Kiểm thử Design Tokens, tỷ lệ tương phản WCAG 2.1, cú pháp QSS và độ sẵn sàng của SVG
pytest tests/unit/test_ui_theme.py -v
```
*Điều kiện đạt*: 
- Tỷ lệ tương phản màu chữ chính trên nền bảng đạt $\ge 7:1$ (WCAG AAA).
- Tỷ lệ tương phản nhãn phụ đạt $\ge 4.5:1$ (WCAG AA).
- Không có cảnh báo lỗi phân giải cú pháp QSS từ Qt.
- 100% tệp SVG tồn tại và render hợp lệ trên `QIcon`.

### 5.3. Kiểm thử khởi chạy không đầu (Headless Health-Check)
```powershell
python SSBOM_Launcher.py --health-check
```
*Điều kiện đạt*: Đầu ra in chuỗi `[OK] SSBOM_App health check passed.` với mã thoát `0`.

---
*Báo cáo khảo sát này cung cấp đầy đủ cơ sở thực nghiệm và chỉ dẫn kỹ thuật để Đội ngũ Triển khai (Coder / Implementer) tiến hành xây dựng giao diện Data-Dense Enterprise Dashboard theo đúng chuẩn `SPEC_UI_UX_ENTERPRISE_DASHBOARD.md` mà không làm đứt gãy bất kỳ logic nghiệp vụ nào.*
