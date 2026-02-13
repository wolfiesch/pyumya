# pyumya — Design & Implementation Plan

Created: 02/12/2026
Status: **Planning**

## Vision

A fast, ergonomic Python library for reading and writing Excel (.xlsx) files with full formatting support, powered by Rust's [umya-spreadsheet](https://github.com/MathNya/umya-spreadsheet) via PyO3.

**Tagline**: "openpyxl speed, Rust performance."

### Why This Exists

| Library | Read | Write | Modify | Formatting | Speed |
|---------|:----:|:-----:|:------:|:----------:|:-----:|
| openpyxl | Yes | Yes | Yes | Full | Slow (pure Python) |
| python-calamine | Yes | — | — | — | Fast (Rust) |
| fastexcel | Yes | — | — | — | Fast (Rust) |
| xlsxwriter | — | Yes | — | Full | Medium (C) |
| **pyumya** | Yes | Yes | Yes | Full | Fast (Rust) |

No Python library offers fast R+W with formatting. pyumya fills that gap.

### Non-Goals (v1)

- Charts, SmartArt, DrawingML rendering
- Pivot tables / pivot caches
- VBA / macro support
- Encrypted / password-protected workbooks
- Conditional formatting (defer to Phase 3+)
- .xls (legacy binary) format

## Architecture

```
pyumya/                          # Python package (pure Python API layer)
  __init__.py                    # Public API: load_workbook, Workbook()
  workbook.py                   # Workbook class (wraps Rust UmyaBook)
  worksheet.py                  # Worksheet class (cell access, row/col ops)
  cell.py                       # Cell class (value + format access)
  styles.py                     # Font, PatternFill, Border, Side, Alignment, ...
  utils.py                      # Column letter ↔ number, A1 parsing

rust/                           # PyO3 crate (Rust ↔ Python bridge)
  Cargo.toml
  src/
    lib.rs                      # PyO3 module entry
    workbook.rs                 # UmyaBook wrapper (open, save, sheets)
    worksheet.rs                # Sheet-level ops (cells, rows, cols, merges)
    cell_ops.rs                 # Cell value R/W (types, formulas, dates)
    format_ops.rs               # Font, fill, border, alignment, number format
    structural_ops.rs           # Freeze panes, autofilters, gridlines, row/col dims
    utils.rs                    # A1↔(row,col), color conversion, date serial math

tests/                          # pytest suite
benchmarks/                     # Performance comparisons vs openpyxl
```

### Design Principles

1. **Openpyxl-compatible API** — Users should be able to switch with minimal code changes. Same property names (`cell.font`, `cell.fill`, `cell.border`, `cell.alignment`), same patterns (`ws['A1']`, `ws.cell(row=1, column=1)`).

2. **Rust does the heavy lifting** — All XML parsing, cell storage, style resolution, and serialization happen in Rust. Python layer is thin: type conversion, API sugar, iteration protocols.

3. **Lazy by default** — Don't parse formatting until accessed. Don't build full style table until save. This matches umya's internal lazy loading.

4. **Fail loud** — No silent data loss. If a feature isn't supported, raise `NotImplementedError` rather than silently dropping data.

## Phase Plan

### Phase 0: Scaffolding [Target: 1 day]

- [x] Create repo, git init
- [ ] Maturin project setup (`Cargo.toml` + `pyproject.toml`)
- [ ] Minimal PyO3 module that compiles and imports
- [ ] `pip install -e .` via maturin develop works
- [ ] CI: GitHub Actions (build + test on Linux/macOS/Windows)
- [ ] Basic `.gitignore`, `CLAUDE.md`

### Phase 1: Core Cell R/W [Target: 3-5 days]

Port from ExcelBench's `umya_backend.rs` (795 lines), restructured into clean modules.

- [ ] `Workbook.open(path)` / `Workbook()` / `workbook.save(path)`
- [ ] `workbook.sheetnames` / `workbook[name]` / `workbook.create_sheet(name)`
- [ ] `worksheet['A1']` / `worksheet.cell(row, column)` → `Cell` object
- [ ] Cell value types: string, number, boolean, blank, date, datetime, error
- [ ] Formula read/write: `cell.value = "=SUM(A1:A10)"` (preserves formula text)
- [ ] Date serial ↔ ISO conversion (with 1900 leap year bug handling)
- [ ] Sheet iteration: `for row in ws.iter_rows()`, `for col in ws.iter_cols()`
- [ ] Test: roundtrip — write cells, save, reopen, verify values match
- [ ] Test: read ExcelBench fixtures, compare cell values to openpyxl output

### Phase 2: Formatting [Target: 5-7 days]

- [ ] `Font(name, size, bold, italic, underline, strikethrough, color)`
- [ ] `PatternFill(fill_type, fgColor)` — solid fills, pattern fills
- [ ] `Border(left, right, top, bottom, diagonal)` + `Side(style, color)`
- [ ] `Alignment(horizontal, vertical, wrap_text, text_rotation)`
- [ ] `NumberFormat` — format code strings (accounting, date, percentage, custom)
- [ ] `cell.font = Font(bold=True, color="FF0000")`
- [ ] `cell.fill = PatternFill(fgColor="4472C4")` (accept hex with/without #)
- [ ] Theme color resolution (ARGB → hex, tint/shade transforms)
- [ ] Test: roundtrip formatting — write styled cells, save, reopen, verify
- [ ] Test: read ExcelBench formatting fixtures, compare to openpyxl

### Phase 3: Structural Features [Target: 3-5 days]

- [ ] Row height: `ws.row_dimensions[1].height = 20`
- [ ] Column width: `ws.column_dimensions['A'].width = 15`
- [ ] Merged cells: `ws.merge_cells('A1:C1')` / `ws.unmerge_cells('A1:C1')`
- [ ] Freeze panes: `ws.freeze_panes = 'A2'`
- [ ] Autofilters: `ws.auto_filter.ref = 'A1:G100'`
- [ ] Gridlines: `ws.sheet_view.showGridLines = False`
- [ ] Row insertion: `ws.insert_rows(idx, amount)`
- [ ] Row deletion: `ws.delete_rows(idx, amount)`
- [ ] Column insertion/deletion
- [ ] Test: structural features roundtrip

### Phase 4: Advanced Cell Operations [Target: 3-5 days]

- [ ] Hyperlinks: `cell.hyperlink = "https://example.com"`
- [ ] Comments: `cell.comment = Comment("text", "author")`
- [ ] Images: `ws.add_image(Image("logo.png"), "A1")`
- [ ] Data validation (basic): dropdown lists, number ranges
- [ ] Named ranges (if straightforward in umya)
- [ ] Test: each feature roundtrips correctly

### Phase 5: API Polish & Performance [Target: 3-5 days]

- [ ] Context manager: `with pyumya.load_workbook("file.xlsx") as wb:`
- [ ] Pythonic iteration: `for cell in row:`, `for row in ws.rows:`
- [ ] `ws.append([val1, val2, val3])` for row-at-a-time writes
- [ ] Copy cell formatting: `copy_style(source_cell, target_cell)`
- [ ] Performance benchmarks vs openpyxl (read, write, modify, large files)
- [ ] Memory benchmarks (RSS comparison)
- [ ] ExcelBench adapter: register pyumya as an adapter, run full benchmark

### Phase 6: Packaging & Release [Target: 2-3 days]

- [ ] PyPI publishing via maturin (`maturin publish`)
- [ ] Pre-built wheels: Linux (manylinux), macOS (x86 + ARM), Windows
- [ ] `pip install pyumya` works on all platforms
- [ ] Minimal docs (README with examples)
- [ ] ExcelBench integration: add pyumya adapter to benchmark suite
- [ ] Migration guide: openpyxl → pyumya (common patterns)

## API Design — Key Decisions

### Decision 1: Mirror openpyxl or clean-slate API?

**Chose: Mirror openpyxl** — because:
- Minimizes switching cost for the massive openpyxl user base
- Property names (`cell.font`, `cell.fill`, `cell.border`) are already good
- Indexing patterns (`ws['A1']`, `ws.cell(row=1, column=1)`) are intuitive
- The "fast drop-in replacement" pitch is stronger than "better API"

### Decision 2: Style objects — mutable or immutable?

**Chose: Mutable (unlike openpyxl)** — because:
- openpyxl's immutable styles (must create new Font to change bold) are a major pain point
- `cell.font.bold = True` is more intuitive than `cell.font = Font(bold=True, **existing_attrs)`
- Rust-side, umya uses mutable style access, so this aligns naturally
- Trade-off: slightly more complex sync logic between Python and Rust layers

### Decision 3: When does Rust state sync with Python?

**Chose: Eager write, lazy read** — because:
- Writes (`cell.value = x`, `cell.font.bold = True`) push to Rust immediately
- Reads (`cell.font.bold`) pull from Rust on access (cached until next write)
- This avoids the "forgot to flush" problem while keeping reads fast
- Trade-off: more FFI calls on write-heavy paths, but writes are typically less frequent than reads

## Testing Strategy

### Layer 1: Unit tests (pure Python + Rust)
- Cell value type conversion (Python ↔ Rust)
- Color format normalization (#RGB, #RRGGBB, ARGB)
- A1 ↔ (row, col) parsing
- Date serial ↔ datetime conversion

### Layer 2: Roundtrip tests
- Write → save → reopen → read → assert match
- For every feature: values, formulas, formatting, structural

### Layer 3: ExcelBench cross-validation
- Run ExcelBench fixtures through pyumya adapter
- Compare scores to openpyxl (target: match on all supported features)
- This is the definitive fidelity test — ExcelBench was built for exactly this

### Layer 4: Performance benchmarks
- Wall clock: open, read all cells, save (vs openpyxl, vs python-calamine)
- Memory: peak RSS during large workbook operations
- Throughput: cells/second for read and write
- Use ExcelBench perf runner for standardized comparison

## Reference Implementation

ExcelBench already has a working umya PyO3 bridge:

- **Rust bindings**: `ExcelBench/rust/excelbench_rust/src/umya_backend.rs` (795 lines)
- **Python adapter**: `ExcelBench/src/excelbench/harness/adapters/umya_adapter.py` (178 lines)
- **Shared utilities**: `ExcelBench/rust/excelbench_rust/src/util.rs`

These cover: cell values, formulas, dates, font styling, background colors, borders, alignment, number formats, row height, column width. This is the starting point for Phase 1-2.

## Competitive Positioning

```
                    Formatting Coverage →
                    None          Basic         Full
                ┌─────────────┬─────────────┬─────────────┐
    Fast (Rust) │ calamine    │             │ pyumya      │
                │ fastexcel   │             │ (target)    │
  Speed ↓       ├─────────────┼─────────────┼─────────────┤
    Med (C/mix) │             │ xlsxwriter  │             │
                │             │ (write only)│             │
                ├─────────────┼─────────────┼─────────────┤
    Slow (Py)   │ pylightxl   │ pandas      │ openpyxl    │
                │             │ tablib      │             │
                └─────────────┴─────────────┴─────────────┘
```

pyumya's unique position: **top-right corner** — fast AND full formatting, for both read AND write.

## Risk Register

| Risk | Likelihood | Impact | Mitigation |
|------|:----------:|:------:|------------|
| umya-spreadsheet has round-trip bugs on complex files | Medium | High | ExcelBench catches these; file bugs upstream or patch locally |
| PyO3 FFI overhead negates Rust speed for small ops | Low | Medium | Batch operations where possible; benchmark early |
| openpyxl API surface is too large to mirror fully | Medium | Medium | Phase delivery — core features first, expand based on demand |
| umya crate development slows or stops | Low | High | Fork if needed; 86 releases over 5 years suggests active maintenance |
| Multi-platform wheel building is painful | Medium | Low | maturin + GitHub Actions handles this well (proven pattern) |

## Session Log (append-only)

### 02/12/2026 — Initial Planning
- Decision: Separate repo (not submodule of ExcelBench)
- Decision: Mirror openpyxl API for minimal switching cost
- Decision: Mutable styles (unlike openpyxl's immutable pattern)
- Decision: Eager write / lazy read for Rust state sync
- Created: Project directory, git repo, planning document
- Reference: ExcelBench umya_backend.rs (795 lines) as starting point
- Next: Phase 0 scaffolding (maturin setup, minimal compilable module)
