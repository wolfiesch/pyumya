# CLAUDE.md — pyumya

## What This Project Is

pyumya is a fast Python library for reading and writing Excel (.xlsx) files with full formatting support, powered by Rust's umya-spreadsheet via PyO3.

## Quick Reference

```bash
# Install for development (builds Rust extension)
maturin develop

# Run tests
pytest

# Build wheel
maturin build --release

# Lint
ruff check src/ tests/
```

## Architecture

```
src/pyumya/             # Python API layer (thin wrappers)
  __init__.py           # Public API: load_workbook, Workbook
  workbook.py           # Workbook class
  worksheet.py          # Worksheet class (Phase 1)
  cell.py               # Cell class (Phase 1)
  styles.py             # Font, PatternFill, Border, etc. (Phase 2)

rust/src/               # PyO3 Rust backend
  lib.rs                # Module entry → pyumya._rust
  workbook.rs           # UmyaBook wrapper
  worksheet.rs          # Sheet-level ops (Phase 1)
  cell_ops.rs           # Cell value R/W (Phase 1)
  format_ops.rs         # Formatting (Phase 2)
  structural_ops.rs     # Freeze panes, filters (Phase 3)

Cargo.toml              # Rust deps (root-level, maturin convention)
pyproject.toml          # Python packaging (maturin backend)
```

## Key Patterns

### Adding a New Feature

1. Implement Rust side in `rust/src/` (add methods to `RustWorkbook` or new structs)
2. Expose via PyO3 `#[pymethods]` or `#[pyclass]`
3. Create/update Python wrapper in `src/pyumya/`
4. Add roundtrip test in `tests/`
5. Validate against ExcelBench fixtures if applicable

### API Design Rule

Mirror openpyxl's API. Same property names, same access patterns.
- `cell.font`, `cell.fill`, `cell.border`, `cell.alignment`
- `ws['A1']`, `ws.cell(row=1, column=1)`
- `wb.sheetnames`, `wb.save("file.xlsx")`

### Rust ↔ Python Boundary

- Writes push to Rust immediately (eager write)
- Reads pull from Rust on access (lazy read)
- Style objects are mutable (unlike openpyxl's immutable pattern)

## Conventions

- **Python 3.9+**, type hints on all function signatures
- **Rust**: edition 2021, standard formatting (`cargo fmt`)
- **Linter**: ruff (line-length 100)
- **Test runner**: pytest
- **Build**: maturin (PyO3 extension module)

## Reference

- Design doc: `docs/plans/pyumya-design.md`
- ExcelBench umya bindings (reference impl): `~/Projects/ExcelBench/rust/excelbench_rust/src/umya_backend.rs`
- ExcelBench Python adapter: `~/Projects/ExcelBench/src/excelbench/harness/adapters/umya_adapter.py`

## Gotchas

- `maturin develop` must be rerun after any Rust code changes
- PyO3 `#[pyclass(unsendable)]` needed because umya's Spreadsheet isn't Send
- umya uses 1-based (col, row) tuples internally — all A1 parsing must convert
- Color formats: umya uses ARGB ("FFRRGGBB"), Python API uses "#RRGGBB"
