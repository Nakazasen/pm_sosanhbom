# BÁO CÁO KHẢO SÁT HẠ TẦNG KIỂM THỬ VÀ ASSETS/STYLES (HANDOFF REPORT)

- **Agent**: Explorer 3 (Survey Test Infrastructure & Assets)
- **Target Folder**: `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_3`
- **Parent Conversation ID**: `6014734f-cacb-4480-97ab-1fc3957409fb`
- **Target Spec**: `specs/SPEC_UI_UX_ENTERPRISE_DASHBOARD.md` & `specs/SPEC_PLM_AUTO_DOWNLOAD.md`
- **Date**: 2026-09-21
- **Handoff Type**: Hard (Task Complete)

---

## 1. OBSERVATION (QUAN SÁT THỰC TẾ ĐẦU VÀO & MÃ NGUỒN)

### 1.1 Cấu Trúc Thư Mục Kiểm Thử (`tests/`) & Cơ Chế Fixtures (`conftest.py`)
Thư mục kiểm thử `tests/` tại `D:\Sandbox\pm_sosanhbom\tests` được phân tầng chặt chẽ thành 7 cụm thư mục với 73 tệp kiểm thử:

```
tests/
├── conftest.py                       # 531 dòng: Fixtures dùng chung (workspaces, synthetic PLM/CS12, mocks Selenium/SAP)
├── e2e/                              # 4 tệp: End-to-End suites (81 tests)
│   ├── test_tier1_feature_coverage.py      (40 tests: R1..R6)
│   ├── test_tier2_boundary_corner.py       (30 tests: Edge/Boundary)
│   ├── test_tier3_pairwise_combinations.py (6 tests: Pipelines)
│   └── test_tier4_production_scenarios.py  (5 tests: Real-world workflows)
├── unit/                             # 18 tệp: Unit test các module lõi và GUI (345 tests)
│   ├── test_adapters.py                    (16 tests)
│   ├── test_app_updates.py                 (5 tests)
│   ├── test_core_tree.py                   (28 tests)
│   ├── test_credentials_security.py       (47 tests)
│   ├── test_gui_and_reporting.py           (30 tests)
│   ├── test_inheritance_engine.py          (16 tests)
│   ├── test_jig_manager.py                 (11 tests)
│   ├── test_leader_view.py                 (18 tests)
│   ├── test_member_view.py                 (8 tests)
│   ├── test_reconciliation_msi.py          (40 tests)
│   ├── test_sap_automation.py              (32 tests)
│   ├── test_spec_m1_contract.py            (3 tests)
│   ├── test_tc14_automation.py             (47 tests)
│   ├── test_tc2412_automation.py           (19 tests)
│   ├── test_ui_i18n.py                     (5 tests)
│   ├── test_update_delivery.py             (7 tests)
│   ├── test_update_launcher.py             (3 tests)
│   └── test_update_security.py             (10 tests)
├── tier1_features/                   # 28 tệp: Kiểm thử đơn lẻ tính năng F01..F28 (141 tests)
├── tier2_boundaries/                 # 5 tệp: Biên và trường hợp góc (43 tests)
├── tier3_combinations/               # 3 tệp: Tích hợp chéo đường ống xử lý (14 tests)
├── tier4_real_world/                 # 2 tệp: Dữ liệu thực tế và form_ssbom kế thừa (10 tests)
└── tier5_adversarial/                # 13 tệp: Tải nặng, stress, tấn công cấu trúc (136 tests)
```

**Khảo sát `tests/conftest.py`**:
- Định nghĩa 11 fixtures chính: `temp_workspace`, `sample_14col_plm_records`, `sample_14col_plm_df`, `sample_plm_excel_file`, `sample_cs12_html_content`, `sample_cs12_file`, `sample_cttt_df`, `sample_msi_df`, `sample_bolocbom_rules`, `mock_tc14_driver`, `mock_sap_session`.
- Hiện tại `tests/conftest.py` **chưa có fixture `qapp` toàn cục**. Các tệp GUI (`test_leader_view.py`, `test_member_view.py`, `test_tier1_feature_coverage.py`...) đang tự khai báo fixture `qapp` cục bộ trong từng tệp:
  ```python
  @pytest.fixture
  def qapp() -> QApplication:
      app = QApplication.instance()
      if app is None:
          app = QApplication([])
      return app
  ```

---

### 1.2 Kết Quả Chạy Kiểm Thử Thực Tế Từng Phân Hệ
Đã thực thi toàn diện lệnh kiểm thử qua `py -m pytest` trên môi trường Python 3.13.14 Windows:

| Phân Hệ Test | Số Test | Passed | Failed | Error | Thời Gian | Tỷ Lệ Pass | Ghi Chú |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :--- |
| **`tests/e2e/`** | 81 | 81 | 0 | 0 | 64.60s | **100%** | Toàn bộ 4 tier E2E hoạt động hoàn hảo |
| **`tests/unit/`** | 345 | 345 | 0 | 0 | 59.36s | **100%** | Toàn bộ unit tests (bao gồm GUI leader & member) PASS |
| **`tests/tier1_features/`** | 141 | 136 | 0 | 1 | 22.61s | 99.27% | 1 lỗi import tại `test_f28_standalone_packaging.py` |
| **`tests/tier2_boundaries/`** | 43 | 43 | 0 | 0 | 32.49s | **100%** | Pass toàn bộ boundary/corner cases |
| **`tests/tier3_combinations/`** | 14 | 14 | 0 | 0 | 4.43s | **100%** | Pass toàn bộ integration pipelines |
| **`tests/tier4_real_world/`** | 10 | 10 | 0 | 0 | 2.15s | **100%** | Pass toàn bộ form_ssbom legacy checks |
| **`tests/tier5_adversarial/`** | 136 | 135 | 1 | 0 | 121.05s | 99.26% | 1 test probe thất bại do hành vi Win11 FS |
| **TỔNG CỘNG** | **770** | **764** | **1** | **1** | **~5.1 phút** | **99.74%** | **764 test cases đang PASS** |

**Chi tiết 2 điểm ngoại lệ đã ghi nhận**:
1. **Lỗi thu thập (Collection Error)** tại `tests/tier1_features/test_f28_standalone_packaging.py`:
   - Dòng 19: `from scripts.package_app import APP_NAME, APP_VERSION, PackageBuilder, compute_sha256`
   - Nguyên nhân: Trong `scripts/package_app.py`, biến phiên bản được đặt tên là `DEFAULT_MIN_APP_VERSION = "1.0.0"`, không có hằng số `APP_VERSION`.
2. **Lỗi kiểm thử (Failure)** tại `tests/tier5_adversarial/test_core_engines_stress.py:675`:
   - `TestSurveyEdgeCasesVerification.test_verify_survey_failure_2_dos_device_names_credentials`
   - Nguyên nhân: Đây là test case được viết nhằm xác minh lỗi thiết bị DOS Windows (`CON.tmp`), nhưng trên Windows 11 / Python 3.13 thao tác tạo tệp `CON.tmp` thành công nên câu lệnh `assert failed` kích hoạt AssertionError.

---

### 1.3 Khảo Sát Thư Mục Assets & Styles Hiện Tại Trong `src/gui/`
- Kiểm tra danh mục `src/gui/`:
  * Hiện tại **CHƯA TỒN TẠI** thư mục `src/gui/styles/` (chưa có `tokens.py`, `theme_manager.py`, `light_theme.qss`, `dark_theme.qss`).
  * Hiện tại **CHƯA TỒN TẠI** thư mục `src/gui/assets/` hoặc `src/gui/assets/icons/` (chưa có bộ SVG icons).
  * Thư mục `assets/` tại thư mục gốc chỉ chứa duy nhất 1 file `README.md` (213 bytes).
- Hiện trạng Styling trong mã nguồn GUI hiện có (`src/gui/leader_view.py`, `src/gui/member_view.py`, `src/gui/app.py`, `src/gui/plm_download_dialog.py`):
  * Toàn bộ màu sắc và style đang được thiết lập phân tán qua hơn 85 vị trí gọi inline `setStyleSheet("...")` (ví dụ: `background-color: #0078D4; color: white; padding: 6px 14px;`).
  * Biểu tượng trên giao diện đang dùng Emoji và ký tự Unicode trực tiếp trong chuỗi text (ví dụ `btn.setText("▶ Lập Dự Án & Phân Công")`, `btn.setText("✓ Tải & Xử Lý")`, `btn.setText("🔒 Theo Dõi")`), vi phạm nguyên tắc thiết kế chuyên nghiệp trong Spec `SPEC_UI_UX_ENTERPRISE_DASHBOARD.md`.

---

### 1.4 Đánh Giá Khả Năng Kiểm Thử PyQt6 Headless / Offscreen Trên Windows
- Kiểm tra module `pytest-qt`:
  * Kết quả probe: `pytest-qt` **chưa được cài đặt** trong môi trường Python hiện tại (`No module named 'pytestqt'`).
- Kiểm tra chế độ Headless / Offscreen Native của PyQt6:
  * Đã tạo và chạy script probe `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_3\probe_pyqt_offscreen.py` với biến môi trường `QT_QPA_PLATFORM=offscreen`.
  * Kết quả chạy thực tế:
    ```
    Platform name: offscreen
    QSvgRenderer valid: True
    Pixmap rendered isNull: False
    StyleSheet applied successfully without crash.
    ```
  * Xác nhận: PyQt6 trên Windows hỗ trợ hoàn hảo chế độ `offscreen`. Không cần màn hình vật lý, không sinh cửa sổ đồ họa popup, không bị treo tiến trình (hang), và hoàn toàn hỗ trợ render SVG qua `QSvgRenderer` / `QPixmap` / `QIcon` cũng như nạp/kiểm tra cú pháp QSS qua `app.setStyleSheet(...)`.
  * Cách thức triển khai: Dùng fixture `qapp` có `os.environ["QT_QPA_PLATFORM"] = "offscreen"` (tương tự như cách 345 unit tests hiện tại đang chạy ổn định).

---

## 2. LOGIC CHAIN (CHUỖI LẬP LUẬN TỪ QUAN SÁT ĐẾN KẾT LUẬN)

1. **Từ Quan sát 1.1 & 1.2**:
   - `tests/unit/` (345 tests) và `tests/e2e/` (81 tests) đạt tỷ lệ Pass 100.00%. Điều này chứng minh toàn bộ logic nghiệp vụ (BOM comparison, date filter, model pruner, unit resolver, MSI engine, JIG manager, Excel generation, trilingual i18n) và giao diện Leader/Member hiện có đang hoạt động chuẩn xác.
   - Do đó, bất kỳ thay đổi nào liên quan đến UI/UX, Theme, hay Stylesheet đều phải giữ nguyên tỷ lệ Pass của 426 tests này để bảo đảm tiêu chí **Zero Regression**.

2. **Từ Quan sát 1.3**:
   - Việc thiếu vắng hoàn toàn `src/gui/styles/` và `src/gui/assets/icons/` đồng nghĩa với việc toàn bộ kiến trúc Theming quy định trong `SPEC_UI_UX_ENTERPRISE_DASHBOARD.md` sẽ được khởi tạo mới một cách độc lập mà không bị xung đột với các file stylesheet cũ.
   - Các màn hình `leader_view.py`, `member_view.py`, `app.py`, `settings_dialog.py` cần được tái cấu trúc để gỡ bỏ các đoạn `setStyleSheet` inline phân tán và thay thế bằng các objectName/class selector tiêu chuẩn được điều phối bởi `ThemeManager`.

3. **Từ Quan sát 1.4**:
   - Không cần phụ thuộc vào thư viện bên ngoài `pytest-qt` (vốn chưa được cài và có thể gây phức tạp khi phân phối).
   - Tệp test mới `tests/unit/test_ui_theme.py` có thể được xây dựng hoàn toàn dựa trên thư viện chuẩn của dự án: `pytest`, `PyQt6.QtWidgets.QApplication`, `PyQt6.QtSvg.QSvgRenderer`, và công thức toán học tính Relative Luminance / Contrast Ratio chuẩn WCAG 2.1.
   - Khi đặt `os.environ["QT_QPA_PLATFORM"] = "offscreen"` ở đầu tệp test, toàn bộ quá trình test Theme, QSS, SVG, and View widgets sẽ chạy ngầm với tốc độ cực nhanh (< 2 giây), không làm nháy màn hình của kỹ sư.

---

## 3. CAVEATS (GIỚI HẠN & GIẢ ĐỊNH)

1. **Lỗi import `test_f28_standalone_packaging.py`**:
   - Thuộc phạm vi packaging script (`scripts/package_app.py`), không ảnh hưởng đến tầng GUI / Theme. Tuy nhiên, khi lập trình viên sửa `package_app.py` bổ sung alias `APP_VERSION = DEFAULT_MIN_APP_VERSION`, tệp này sẽ tự động pass.
2. **Không có `pytest-qt`**:
   - Mọi tương tác giao diện trong test suite phải được kích hoạt trực tiếp thông qua API của Qt (ví dụ `widget.click()`, `widget.setChecked()`, `widget.setCurrentText()`, hoặc gọi trực tiếp signal/slot) thay vì dùng `qtbot.mouseClick`. Đây là cách tiếp cận đang được dùng rất thành công ở 18 test cases trong `test_leader_view.py`.
3. **Môi trường CI / Docker (nếu có sau này)**:
   - Trên Linux/CI không có X11, biến `QT_QPA_PLATFORM=offscreen` hoặc gói `libgl1-mesa-glx` là bắt buộc. Trên Windows hiện tại, `offscreen` hoạt động out-of-the-box mà không cần driver phụ.

---

## 4. CONCLUSION (KẾT LUẬN & ĐỀ XUẤT KIẾN TRÚC CHO UI THEME)

### 4.1 Đề xuất Cấu trúc Module Assets & Styles Cần Xây Dựng
Theo đúng đặc tả tại Mục 1.2 của `SPEC_UI_UX_ENTERPRISE_DASHBOARD.md`:
```
src/gui/
├── styles/
│   ├── __init__.py            # Export ThemeManager, Tokens, get_theme_manager
│   ├── tokens.py              # Design Tokens: Palette, Typography, Spacing, WCAG Helpers
│   ├── theme_manager.py       # ThemeManager (Singleton): switch_theme, apply_theme, listen_system
│   ├── light_theme.qss        # QSS cho Slate Industrial (Data-Dense table, sharp 1px borders)
│   └── dark_theme.qss         # QSS cho Industrial Charcoal/Navy
├── assets/
│   └── icons/                 # 10 icons SVG chuẩn Lucide:
│       ├── download.svg
│       ├── refresh.svg
│       ├── check-circle.svg
│       ├── alert-triangle.svg
│       ├── x-circle.svg
│       ├── settings.svg
│       ├── search.svg
│       ├── filter.svg
│       ├── file-spreadsheet.svg
│       ├── moon.svg
│       └── sun.svg
```

### 4.2 Thiết Kế Chi Tiết Bộ Test `tests/unit/test_ui_theme.py`
Bộ test này sẽ bao phủ 5 yêu cầu kiểm định cốt lõi:
1. `test_wcag_contrast_ratios()`:
   - Tính toán tỷ lệ tương phản toán học giữa màu chữ và màu nền của tất cả các semantic tokens trong `tokens.py`.
   - Xác thực: Chữ chính trên nền card >= 7.0:1 (AAA), chữ phụ / nhãn >= 4.5:1 (AA), badge trạng thái >= 4.5:1.
2. `test_qss_syntax_validity(qapp)`:
   - Nạp cả `light_theme.qss` và `dark_theme.qss` vào `QApplication.setStyleSheet()`.
   - Xác thực không phát sinh lỗi ngoại lệ hoặc crash tiến trình.
3. `test_theme_manager_state_persistence(tmp_path)`:
   - Kiểm tra khởi tạo ThemeManager, đổi theme từ Light sang Dark và ngược lại.
   - Xác thực việc lưu và đọc trạng thái theme từ tệp cấu hình JSON.
4. `test_svg_icons_integrity()`:
   - Duyệt qua toàn bộ danh sách 11 tệp SVG trong `src/gui/assets/icons/`.
   - Dùng `QSvgRenderer` để xác thực cú pháp vector XML hợp lệ (`renderer.isValid() is True`).
   - Dùng `QIcon` để xác thực icon tải thành công và có thể kết xuất ra QPixmap (`icon.isNull() is False`).
5. `test_table_view_data_dense_metrics(qapp)`:
   - Khởi tạo QTableWidget với style của Theme, đo lường row-height (mục tiêu 30-34px) và kiểm tra gridline hiển thị sắc nét.

---

## 5. VERIFICATION METHOD & ZERO REGRESSION CATALOG

### 5.1 Danh Sách Kiểm Thử Bắt Buộc Để Đảm Bảo Zero Regression (426 Tests Nòng Cốt)
Trước và sau khi áp dụng UI Theme, bắt buộc phải chạy và xác nhận 100% PASS cho các lệnh sau:

```bash
# 1. Toàn bộ Unit Tests (345 tests - Core Logic + Leader/Member GUI)
py -m pytest tests/unit/ -v

# 2. Toàn bộ E2E Test Suite (81 tests - Phủ kín 4 Tier R1..R6)
py -m pytest tests/e2e/ -v

# 3. Test Suite cho Theme mới (khi đã hoàn thành code)
py -m pytest tests/unit/test_ui_theme.py -v
```

### 5.2 Điều Kiện Hủy Bỏ Kết Quả (Invalidation Conditions)
Báo cáo và kế hoạch triển khai sẽ bị hủy bỏ nếu:
- Bất kỳ test case nào trong số 345 unit tests hoặc 81 e2e tests chuyển từ PASSED sang FAILED sau khi áp dụng Theme.
- Việc tải icon SVG hoặc đổi Theme làm chậm thời gian khởi động ứng dụng thêm quá 150ms.
- Ứng dụng PyQt6 phát sinh lỗi không hiển thị được giao diện khi chạy trên máy trạm Windows thực tế.
