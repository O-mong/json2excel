# json2excel

Desktop tool to convert nested JSON files to flat Excel spreadsheets with hierarchical depth columns.

Two implementations available: Python (PyQt6) and Rust (egui).

## Features

- Drag & drop JSON files
- Flattens nested objects/arrays into `Depth_1`, `Depth_2`, ..., `value` columns
- Single or batch conversion
- Works offline
- Supports semi-JSON input (e.g., with trailing commas)

## Python Version

**Requirements:**
```bash
pip install pandas openpyxl PyQt6
```

**Run:**
```bash
python json2excel.py
```

**Build executable:**
```bash
pip install pyinstaller
pyinstaller Json2Excel.spec
```

## Rust Version

**Requirements:**
- Rust 1.70+ ([install](https://rustup.rs/))

**Run:**
```bash
cd json2excel
cargo run --release
```

**Build:**
```bash
cd json2excel
cargo build --release
```

Binary output: `json2excel/target/release/json2excel[.exe]`

## Example

**Input JSON:**
```json
{
  "user": {
    "name": "Alice",
    "age": 30,
    "scores": [85, 92, 78]
  },
  "status": "active"
}
```

**Output Excel:**
```
┌─────────┬─────────┬─────────┬────────┐
│ Depth_1 │ Depth_2 │ Depth_3 │ value  │
├─────────┼─────────┼─────────┼────────┤
│ user    │ name    │         │ Alice  │
│ user    │ age     │         │ 30     │
│ user    │ scores  │ 0       │ 85     │
│ user    │ scores  │ 1       │ 92     │
│ user    │ scores  │ 2       │ 78     │
│ status  │         │         │ active │
└─────────┴─────────┴─────────┴────────┘
```

## Usage

1. Launch app
2. Drag & drop `.json` file(s)
3. Click "Convert"
4. Click "Download" and choose save location

## Rust vs Python

**Rust:**
- Faster (2-5x) for large files
- Smaller binary
- No runtime dependencies

**Python:**
- Easier to modify
- Familiar for Python developers

## Updates
- Supporting semi-json file: For now... trailing commas are now supported