# E2E Test Infra: Data-Dense Enterprise Dashboard (`pm_sosanhbom`)

## Test Philosophy
- Opaque-box, requirement-driven, kết hợp kiểm định toán học WCAG 2.1 và kiểm tra tính toàn vẹn cú pháp QSS.
- Chế độ chạy: `QT_QPA_PLATFORM=offscreen` trên Windows headless (không cần display server).
- Đảm bảo Zero Regression cho toàn bộ 426+ test cases hiện hữu.

## Feature Inventory & Test Coverage
| # | Feature | Source (Requirement) | Test Case ID | Test Target File |
|---|---------|----------------------|:------------:|:-----------------|
| 1 | WCAG Contrast Engine & Tokens | SPEC §3.1, §3.2, §3.3 | TEST-UI-01 | `tests/unit/test_ui_theme.py` |
| 2 | QSS Stylesheet Validity (Qt6) | SPEC §3.5, R1 | TEST-UI-02 | `tests/unit/test_ui_theme.py` |
| 3 | Leader View Data-Dense Grid (32px) | SPEC §4.1, R2 | TEST-UI-03 | `tests/unit/test_leader_view_ui.py` |
| 4 | Theme Switch & Persistence | SPEC §4.4, R1 | TEST-UI-04 | `tests/unit/test_ui_theme.py` |
| 5 | SVG Icons Availability & Rendering | SPEC §5, R3 | TEST-UI-05 | `tests/unit/test_ui_theme.py` |

## Test Architecture
- **Runner**: `pytest -v tests/unit/test_ui_theme.py` và `pytest -v tests/unit/test_leader_view_ui.py`
- **Environment**: Windows 11, Python 3.13, PyQt6 with `QT_QPA_PLATFORM=offscreen`
- **Pass/Fail Semantics**:
  - TEST-UI-01: Tỉ lệ tương phản màu chữ chính >= 7:1 (AAA), văn bản phụ/CTA >= 4.5:1 (AA).
  - TEST-UI-02: QSS load vào `QApplication.setStyleSheet()` không ném ngoại lệ, không cảnh báo cú pháp.
  - TEST-UI-03: `table.verticalHeader().defaultSectionSize() == 32`, `showGrid() == True`.
  - TEST-UI-04: Chuyển đổi Light -> Dark -> Light mượt mà, lưu vào json và nạp lại chính xác.
  - TEST-UI-05: 16 tệp SVG tại `src/gui/assets/icons/` tồn tại, nội dung XML hợp lệ, `QIcon(path).isNull() is False`.
