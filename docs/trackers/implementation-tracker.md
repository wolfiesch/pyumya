# pyumya Implementation Tracker

## Phase 1: Core Cell R/W

| # | Task | Status | Notes |
|---|------|--------|-------|
| 1.1 | Rust: cell_ops.rs (read/write cell values - string, number, bool, blank, error) | done | Implemented read/write payloads on RustWorkbook. |
| 1.2 | Rust: cell_ops.rs (formula read/write) | done | Supports payload type "formula" + error-formula mapping. |
| 1.3 | Rust: cell_ops.rs (date/datetime serial <-> ISO conversion) | done | ISO <-> Excel serial conversion + date format heuristic. |
| 1.4 | Rust: worksheet.rs (cell access by A1 ref, row/col iteration) | done | Exposed sheet max row/col for Python iterators. |
| 1.5 | Rust: utils.rs (A1<->row/col, column letter conversion, color helpers) | done | Ported helpers from ExcelBench; will be used in Phase 2. |
| 1.6 | Python: Cell class with .value, .data_type properties | done | Implemented in `src/pyumya/cell.py`. |
| 1.7 | Python: Worksheet class with ['A1'], .cell(row,col), .iter_rows(), .iter_cols() | done | Implemented in `src/pyumya/worksheet.py`. |
| 1.8 | Python: Workbook['SheetName'] returns Worksheet (not str) | done | Updated `src/pyumya/workbook.py` to return Worksheet objects. |
| 1.9 | Tests: cell value roundtrip (all types) | done | `tests/test_cell_values.py`. |
| 1.10 | Tests: formula roundtrip | done | `tests/test_formulas.py`. |
| 1.11 | Tests: date/datetime roundtrip | done | Covered in `tests/test_cell_values.py`. |
| 1.12 | Tests: worksheet access patterns | done | `tests/test_worksheet.py`. |

## Phase 2: Formatting

| # | Task | Status | Notes |
|---|------|--------|-------|
| 2.1 | Rust: format_ops.rs (read/write font properties) | pending | |
| 2.2 | Rust: format_ops.rs (read/write fill/background) | pending | |
| 2.3 | Rust: format_ops.rs (read/write borders) | pending | |
| 2.4 | Rust: format_ops.rs (read/write alignment) | pending | |
| 2.5 | Rust: format_ops.rs (read/write number formats) | pending | |
| 2.6 | Rust: structural_ops.rs (row height, column width) | pending | |
| 2.7 | Rust: structural_ops.rs (merged cells) | pending | |
| 2.8 | Rust: structural_ops.rs (freeze panes via sheet views) | pending | |
| 2.9 | Python: Font class (name, size, bold, italic, underline, strikethrough, color) | pending | |
| 2.10 | Python: PatternFill class (fill_type, fgColor) | pending | |
| 2.11 | Python: Border + Side classes (left/right/top/bottom/diagonal, style, color) | pending | |
| 2.12 | Python: Alignment class (horizontal, vertical, wrap_text, text_rotation) | pending | |
| 2.13 | Python: cell.font, cell.fill, cell.border, cell.alignment, cell.number_format properties | pending | |
| 2.14 | Python: ws.row_dimensions[n].height, ws.column_dimensions['A'].width | pending | |
| 2.15 | Python: ws.merge_cells() / ws.unmerge_cells() | pending | |
| 2.16 | Python: ws.freeze_panes property | pending | |
| 2.17 | Tests: formatting roundtrip (all style types) | pending | |
| 2.18 | Tests: merge cells roundtrip | pending | |
| 2.19 | Tests: freeze panes roundtrip | pending | |
| 2.20 | Tests: row height / column width roundtrip | pending | |

## Session Log
(append entries as you complete tasks)

### 2026-02-13
- Added `rust/src/utils.rs`, `rust/src/cell_ops.rs`, and `rust/src/worksheet.rs`; wired into `rust/src/workbook.rs` and `rust/src/lib.rs`.
- `cargo check` passes (minor warnings; expected until Phase 2 uses color/border helpers).

### 2026-02-13
- Added Python `Cell` + `Worksheet` wrappers, updated `Workbook` to return `Worksheet`.
- Added roundtrip tests for values, formulas, dates, and worksheet iteration; `pytest` passes.
