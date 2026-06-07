# CSV Paste Exporter

**A lightweight Windows tool for pasting, cleaning, previewing, and exporting scientific table data as CSV/TXT/TSV.**

CSV Paste Exporter is built for researchers, students, and engineers who often move experimental data between Origin, Excel, Python, MATLAB, instrument software, and plotting scripts.

It is especially useful for materials science workflows such as:

- tensile stress-strain data
- XRD patterns
- spectroscopy tables
- wide CSV files with many columns
- Origin worksheet columns copied from different projects
- quick column extraction before plotting or debugging scripts

中文简介：这是一个面向材料学实验数据处理的小工具。你可以从 Origin、Excel、CSV、TXT 或仪器软件中复制若干列，直接粘贴到程序里，预览整理后的表格，然后导出为 CSV、TXT 或 TSV 文件。

## Download

For most Windows users, download the ready-to-run `.exe` from:

**[Latest Release](https://github.com/D-sudoasd/csv-paste-exporter/releases/latest)**

Download and run the Windows asset:

```text
CSV-Paste-Exporter-Windows-v0.2.0.exe
```

No Python installation is required for the Windows release build.

## Why This Tool Exists

Scientific data is often stored in wide tables with many columns. During analysis, you may need to copy only a few columns, reformat them, and save them as a clean input file for another program.

Typical examples:

- Copy strain, stress, and time columns from a tensile test table.
- Copy `2theta` and intensity columns from an XRD worksheet.
- Split a wide Origin worksheet into smaller files for plotting scripts.
- Convert copied table data into a clean CSV/TXT file without manually editing delimiters.

This tool keeps that workflow simple:

```text
Copy columns -> Paste -> Preview -> Export
```

## Features

- Paste table data directly from clipboard.
- Auto-detect common delimiters: tab, comma, semicolon, and whitespace.
- Remove fully empty rows and columns.
- Preserve headers, units, scientific notation, and original text values.
- Pad uneven rows with empty cells so the exported table is rectangular.
- Preview data before export.
- Treat the first row as a header for easier preview.
- Delete selected columns.
- Move columns left or right.
- Restore the original parsed table after column edits.
- Export as CSV, TXT, or TSV.
- Choose encoding: UTF-8 BOM, UTF-8, or GBK.
- Remember the last export directory and export settings.
- Open the output folder after export.
- Runs as a single Windows `.exe`.

## Format Guide

Different scientific tools prefer different text table formats. The defaults are chosen for compatibility rather than strict minimalism.

| Target software | Recommended format | Recommended encoding | Notes |
| --- | --- | --- | --- |
| Excel on Windows | CSV | UTF-8 BOM | Best chance of opening Chinese headers without mojibake. |
| Python / pandas | CSV or TSV | UTF-8 | UTF-8 BOM is also readable with `encoding="utf-8-sig"`. |
| Origin | TXT / tab-delimited | UTF-8 BOM or GBK | Tab-delimited TXT is common in scientific workflows. |
| MATLAB | CSV or TSV | UTF-8 | Use `readtable`, `readmatrix`, or import tools. |
| R | CSV or TSV | UTF-8 | Use `read.csv` or `read.delim`. |
| Instrument software | TXT / tab-delimited | UTF-8 BOM or GBK | Some older Chinese Windows tools may prefer GBK. |

## Quick Start for Users

1. Download `CSV列数据整理导出.exe` from the [Latest Release](https://github.com/D-sudoasd/csv-paste-exporter/releases/latest).
2. Double-click the EXE.
3. Copy columns from Origin, Excel, CSV, TXT, or an instrument data table.
4. Click **粘贴剪贴板** or paste into the text area.
5. Check the preview.
6. Choose export format and encoding.
7. Click **导出**.

## Quick Start for Developers

This project uses only the Python standard library at runtime.

Run from source:

```powershell
py csv_paste_exporter.py
```

Run tests:

```powershell
py -m pytest -q
```

Build the Windows EXE:

```powershell
py -m pip install -r requirements-dev.txt
py -m PyInstaller --noconfirm --clean --onefile --windowed --name "CSV列数据整理导出" --distpath "dist" --workpath "build" --specpath "build" "csv_paste_exporter.py"
```

The built EXE will be created at:

```text
dist/CSV列数据整理导出.exe
```

## Testing

The test suite covers:

- tab-delimited clipboard text
- quoted CSV cells
- empty row and empty column cleanup
- uneven row padding
- whitespace-delimited fallback parsing
- CSV/TXT/TSV export
- UTF-8 BOM, UTF-8, and GBK encoding behavior
- column delete and move operations
- first-row-as-header preview logic
- settings file round trip

Run:

```powershell
py -m pytest -q
py -m py_compile csv_paste_exporter.py tests/test_csv_paste_exporter.py
```

## Roadmap

- Drag-and-drop file import.
- Batch split selected column groups into multiple files.
- Optional curve-data preprocessing for stress-strain and XRD workflows.
- Better preview for very large tables.
- English UI option.
- Signed Windows releases if the project gains enough users.

## Contributing

Issues and suggestions are welcome. Helpful feedback includes:

- the software you copied data from
- an example of the pasted table structure
- the export format you need
- whether the result opens correctly in Origin, Excel, Python, MATLAB, or other tools

Please avoid sharing sensitive or unpublished research data in public issues. A small synthetic table with the same structure is enough.

## License

MIT License. See [LICENSE](LICENSE).
