# Project: Data-Dense Enterprise Dashboard UI/UX (`pm_sosanhbom`)

## Architecture
- **Design Tokens & Styling Engine**: `src/gui/styles/` (`tokens.py`, `theme_manager.py`, `light_theme.qss`, `dark_theme.qss`). Single Source of Truth cho toàn bộ mã màu, font, độ cao dòng bảng, và quản lý hot-reload theme.
- **Iconography**: `src/gui/assets/icons/` chứa 16+ vector SVG (Lucide/Fluent), hỗ trợ tinting/monochrome, không dùng emoji.
- **Main Shell**: `src/gui/app.py` quản lý QMainWindow, khởi tạo ThemeManager, gắn SVG icons vào Menu/Tabs/Toolbar, tích hợp Log Dock và Status Bar.
- **Workspaces**:
  - `src/gui/leader_view.py`: Quản lý dự án tuần tự 4 bước, bổ sung 4 KPI cards, chuẩn hóa 4 bảng biểu 32px viền sắc nét, sửa logic checkmark file.
  - `src/gui/member_view.py`: Không gian thành viên, Stepper 3 bước với icon SVG, 3 bảng nhập liệu 32px, khung Diff View trực quan chuẩn WCAG AAA.
- **Dialogs**:
  - `src/gui/plm_download_dialog.py`: Tải BOM PLM/R3 tự động, thanh tiến độ 14px tinh tế, log Consolas 11px.
  - `src/gui/settings_dialog.py`: Cấu hình hệ thống, tích hợp Tab chọn Theme (Light/Dark/System) và hot-reload.
- **Testing Architecture**:
  - `tests/unit/test_ui_theme.py`: Tự động kiểm tra toán học WCAG contrast (AA >= 4.5:1, AAA >= 7:1), QSS syntax validity, SVG availability, và Theme persistence.
  - Chế độ kiểm thử: `QT_QPA_PLATFORM=offscreen` trên Windows headless.

## Feature Inventory
| # | Feature | Description | Milestone | Source |
|---|---------|-------------|-----------|--------|
| 1 | Design Tokens (`tokens.py`) | Định nghĩa màu sắc Light/Dark, font size, row-height 32px, padding, hàm toán học WCAG | M1 | SPEC §3.1, §3.2, §3.3 |
| 2 | ThemeManager (`theme_manager.py`) | Quản lý chuyển đổi theme, hot-reload không restart app, persistence config.json | M1 | SPEC §1.2, §4.4 |
| 3 | Light QSS (`light_theme.qss`) | Stylesheet hoàn chỉnh cho Slate Industrial, viền ô #CBD5E1, row-height 32px, hover #F1F5F9 | M1 | SPEC §3.1, §3.5 |
| 4 | Dark QSS (`dark_theme.qss`) | Stylesheet hoàn chỉnh cho Industrial Dark Mode, viền ô #2A374A, hover #1A2436 | M1 | SPEC §3.2, §3.5 |
| 5 | SVG Icons Library (No Emoji) | 16 tệp SVG vector chuẩn Lucide/Fluent tại `src/gui/assets/icons/`, hàm tinting | M1 | SPEC §5, RES-UI-04 |
| 6 | Leader View: 4 KPI Cards | 4 thẻ thống kê tinh gọn: Tổng Model, Model đủ BOM, Tiến độ CTTT, Trạng thái đối soát | M2 | SPEC §4.1, R2 |
| 7 | Leader View: Table Grid 32px | 4 bảng (machine, staff, sourcing, submission) cao 32px, viền 1px, setShowGrid(True) | M2 | SPEC §3.5, RES-UI-01 |
| 8 | Leader View: Fix Checkmark Bug | Sửa logic L1349/L1371 từ chuỗi "✓" sang kiểm tra file m.plm_file.exists() | M2 | Explorer 1 forensic |
| 9 | Member View: Stepper 3 Bước | Thanh tiến trình quy trình 3 bước với icon SVG chỉ báo trạng thái | M2 | SPEC §4.2, R2 |
| 10 | Member View: Table Grid 32px | 3 bảng (cttt, msi, label) cao 32px, padding 4px 8px, viền sắc nét | M2 | SPEC §3.5, RES-UI-01 |
| 11 | Member View: WCAG Diff View | Tô màu tương phản cao cho dòng lệch (#FEE2E2/#7F1D1D) và khớp (#DCFCE7/#064E3B) | M2 | SPEC §4.2, §3.3 |
| 12 | PLM Download Dialog Modern | Thanh tiến độ 14px bo góc, log Consolas 11px auto-scroll, loại bỏ emoji | M3 | SPEC §4.3, R2 |
| 13 | Settings Dialog Theme Tab | Thêm tab chọn Theme (Light/Dark/System), lưu cấu hình, hot-reload tức thời | M3 | SPEC §4.4, R2 |
| 14 | App Shell Integration | Khởi tạo ThemeManager, gắn icon SVG vào menu/tabs/toolbar, xóa emoji | M3 | SPEC §1.2, R1 |
| 15 | Automated UI Theme Tests | tests/unit/test_ui_theme.py kiểm tra WCAG, QSS, SVG, Theme persistence | M4 | SPEC §7.1, R5 |
| 16 | Zero Regression Verification | Chạy toàn bộ test suite (426+ unit/e2e tests) đảm bảo 100% PASS | M4 | SPEC §7.1, R5 |

## Milestones
| # | Name | Scope | Dependencies | Status |
|---|------|-------|-------------|--------|
| M1 | Theme Architecture, Tokens & SVG Assets | `src/gui/styles/tokens.py`, `theme_manager.py`, `light_theme.qss`, `dark_theme.qss`, `src/gui/assets/icons/*.svg` | None | DONE |
| M2 | Leader View & Member View Data-Dense Upgrade | `src/gui/leader_view.py`, `src/gui/member_view.py` (KPI cards, 32px tables, Stepper, Diff View, fix L1349 bug, no emoji) | M1 | DONE |
| M3 | Dialogs & App Shell Theme Integration | `src/gui/plm_download_dialog.py`, `src/gui/settings_dialog.py`, `src/gui/app.py`, `src/gui/update_dialog.py` | M1, M2 | IN_PROGRESS |
| M4 | UI Theme Test Suite & Zero Regression Gate | `tests/unit/test_ui_theme.py`, `tests/unit/test_leader_view_ui.py`, full suite execution | M1, M2, M3 | PLANNED |

## Interface Contracts
### `src/gui/styles/tokens.py` ↔ Consumers
- Cung cấp:
  - Bảng màu: `COLOR_LIGHT_*`, `COLOR_DARK_*`
  - Kích thước: `TABLE_ROW_HEIGHT = 32`, `CELL_PADDING = (4, 8)`, `PROGRESS_BAR_HEIGHT = 14`, `BORDER_WIDTH = 1`
  - Tiện ích: `calculate_contrast_ratio(fg, bg) -> float`, `is_wcag_aa(fg, bg) -> bool`, `is_wcag_aaa(fg, bg) -> bool`

### `src/gui/styles/theme_manager.py` ↔ Consumers
- Cung cấp:
  - `ThemeManager(QObject)` với signal `theme_changed = pyqtSignal(str)`
  - `get_theme_manager() -> ThemeManager` (Singleton)
  - `set_theme(theme_name: str, save_preference: bool = True)`
  - `get_styled_icon(icon_name: str, color: str = None) -> QIcon`
  - `get_icon_path(icon_name: str) -> Path`

### `LeaderWorkspaceView` ↔ Unit Tests
- BẮT BUỘC giữ nguyên vẹn các delegation:
  - `view.model_combo`, `view.stage_combo`, `view.submission_table`
  - `view.create_project_folder_structure()`, `view.scan_member_submissions()`, `view.get_sub_unit_statuses()`
  - `view.trigger_batch_reconciliation()`, `view._open_plm_download_dialog()`

### `MemberWorkspaceView` ↔ Unit Tests
- BẮT BUỘC giữ nguyên vẹn:
  - `view.cttt_table`, `view.msi_table`, `view.label_table`
  - `view.add_cttt_row()`, `view.clear_cttt_table()`, `view.get_cttt_table_data()`
  - `view.set_reference_data()`, `view.run_preliminary_self_check()`, `view.submit_data()`, `view.unlock_submission()`
  - Signal `submission_completed`

## Code Layout
- `src/gui/styles/`: Theme tokens, manager, QSS files (Owned by M1)
- `src/gui/assets/icons/`: Vector SVG icons (Owned by M1)
- `src/gui/leader_view.py`: Leader Workspace (Owned by M2)
- `src/gui/member_view.py`: Member Workspace (Owned by M2)
- `src/gui/plm_download_dialog.py`: PLM Download Dialog (Owned by M3)
- `src/gui/settings_dialog.py`: Settings Dialog (Owned by M3)
- `src/gui/app.py`: Main Window Shell (Owned by M3)
- `tests/unit/test_ui_theme.py`: UI Theme unit tests (Owned by M1 & M4)
