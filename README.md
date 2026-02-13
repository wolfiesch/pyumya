# pyumya

Fast Excel (.xlsx) reader/writer with full formatting support, powered by Rust.

```python
import pyumya

# Read an existing workbook
wb = pyumya.load_workbook("report.xlsx")
print(wb.sheetnames)

# Create a new workbook
wb = pyumya.Workbook()
wb.create_sheet("Data")
wb.save("output.xlsx")
```

## Installation

```bash
pip install pyumya
```

## Why pyumya?

- **Fast** — Rust-powered XML parsing and serialization (10-50x faster than openpyxl)
- **Full formatting** — Fonts, colors, borders, fills, alignment, number formats
- **Read + Write + Modify** — Open existing files, make changes, save back
- **openpyxl-compatible API** — Minimal code changes to switch
