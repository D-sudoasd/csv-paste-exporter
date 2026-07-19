<p align="center">
  <img src="assets/readme/hero.svg" width="100%" alt="CSV Paste Exporter: paste scientific table columns, preview, and export clean CSV, TXT, or TSV.">
</p>

# CSV Paste Exporter

**Paste columns. Preview. Export clean tables for Origin, Excel, Python, and MATLAB.**

A lightweight Windows tool for researchers, students, and engineers who move experimental data between instrument software, Origin, Excel, and scripts — without hand-editing delimiters.

中文：从 Origin / Excel / 仪器软件复制若干列 → 粘贴 → 预览 → 导出 CSV / TXT / TSV。

## Workflow

```text
Copy columns  →  Paste  →  Preview  →  Export
```

Typical materials-science uses: tensile stress–strain columns, XRD `2theta`/intensity pairs, spectroscopy tables, wide Origin worksheets split for plotting scripts.

## Download

**[Latest Release](https://github.com/D-sudoasd/csv-paste-exporter/releases/latest)**

Windows users: run the packaged `.exe` (for example `CSV-Paste-Exporter-Windows-v0.2.0.exe`). No Python install required for the release build.

## What it does

- Paste table data from the clipboard; auto-detect tab, comma, semicolon, or whitespace.
- Remove fully empty rows/columns; pad uneven rows so the export is rectangular.
- Preserve headers, units, scientific notation, and original text values.
- Preview the cleaned table; optional first-row-as-header; sample values beside column names.
- Delete or reorder columns; restore the original parsed table after edits.
- Quick point-line chart preview from selected X/Y columns.
- Target presets for Excel, Origin, Python/pandas, MATLAB, or legacy GBK instruments.
- Export CSV / TXT / TSV with UTF-8 BOM, UTF-8, or GBK.
- Export-readiness check (row/column counts, empty cells, cleanup warnings).
- Remember last export directory and settings; open the output folder after export.

## Format guide

| Target | Recommended format | Encoding | Notes |
| --- | --- | --- | --- |
| Excel (Windows) | CSV | UTF-8 BOM | Best chance for Chinese headers |
| Python / pandas | CSV or TSV | UTF-8 | `utf-8-sig` also works with BOM |
| Origin | TXT (tab) | UTF-8 BOM or GBK | Common lab workflow |
| MATLAB / R | CSV or TSV | UTF-8 | `readtable` / `read.csv` |
| Older instrument tools | TXT (tab) | UTF-8 BOM or GBK | Some Chinese Windows tools prefer GBK |

## Quick start (users)

1. Download the EXE from the [Latest Release](https://github.com/D-sudoasd/csv-paste-exporter/releases/latest).
2. Double-click the EXE.
3. Copy columns from Origin, Excel, CSV, TXT, or instrument software.
4. Click **粘贴剪贴板** or paste into the text area.
5. Check the preview; choose format and encoding.
6. Click **导出**.

## Develop from source

Runtime uses the Python standard library only.

```powershell
py csv_paste_exporter.py
py -m pytest -q
```

Build the Windows EXE:

```powershell
py -m pip install -r requirements-dev.txt
py -m PyInstaller --noconfirm --clean --onefile --windowed --name "CSV列数据整理导出" --distpath "dist" --workpath "build" --specpath "build" "csv_paste_exporter.py"
```

Output: `dist/CSV列数据整理导出.exe`

## Testing

```powershell
py -m pytest -q
py -m py_compile csv_paste_exporter.py tests/test_csv_paste_exporter.py
```

Coverage includes delimiter parsing, empty-row cleanup, uneven-row padding, encodings, presets, column ops, chart extraction, and settings round-trip.

## Roadmap

- Drag-and-drop file import
- Batch split of column groups into multiple files
- Optional preprocessing for stress–strain / XRD curves
- Better preview for very large tables
- English UI option
- Signed Windows releases if demand grows

## Contributing

Issues welcome. Useful reports include: source software, example table structure (synthetic is fine), target export format, and whether the result opens correctly downstream. Please do not share sensitive or unpublished research data.

## License

MIT — see [LICENSE](LICENSE).
