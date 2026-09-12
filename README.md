<p align="center">
  <img src="assets/readme/hero.svg" width="100%" alt="CSV Paste Exporter: paste scientific columns, preview, export clean tables.">
</p>

# CSV Paste Exporter

**Paste columns. Preview. Export clean tables for Origin, Excel, Python, and MATLAB.**

A lightweight Windows tool for researchers and engineers who move experimental data between instrument software, Origin, Excel, and scripts — without hand-editing delimiters.

中文：从 Origin / Excel / 仪器软件复制若干列 → 粘贴 → 预览 → 导出 CSV / TXT / TSV。

<p align="center">
  <img src="assets/readme/section-01-why.svg" width="100%" alt="01 Why: from clipboard chaos to a rectangular table.">
</p>

```text
Copy columns  →  Paste  →  Preview  →  Export
```

Typical uses: tensile stress–strain columns, XRD `2theta`/intensity pairs, spectroscopy tables, wide Origin worksheets split for plotting scripts.

### Capabilities

- Auto-detect tab, comma, semicolon, or whitespace  
- Remove empty rows/columns; pad uneven rows  
- Preserve headers, units, scientific notation, original text  
- Column delete/reorder; restore original parse  
- Quick X/Y chart preview · export readiness check  
- Target presets: Excel, Origin, Python/pandas, MATLAB, GBK instruments  
- Encoding: UTF-8 BOM, UTF-8, GBK  

### Format guide

Matches the **目标软件** presets in the app. TSV is available from **格式** (or **自定义**).

| Target | Format | Encoding |
| --- | --- | --- |
| Excel | CSV | UTF-8 BOM |
| Origin | TXT (tab) | UTF-8 BOM |
| Python / pandas | CSV | UTF-8 |
| MATLAB | CSV | UTF-8 |
| 旧仪器/GBK | TXT (tab) | GBK |

<p align="center">
  <img src="assets/readme/section-02-use.svg" width="100%" alt="02 Use: download the EXE and export in minutes.">
</p>

## Download

**[Latest Release](https://github.com/D-sudoasd/csv-paste-exporter/releases/latest)** — Windows EXE, no Python required. The tagged installer may lag `main`; run from source (below) for the current GUI, presets, and chart.

## Use the GUI

1. Copy columns from Origin / Excel / instrument software  
2. **粘贴剪贴板** (or paste into **粘贴区**)  
3. Optional: **第一行是表头**, **目标软件** (Excel / Origin / Python/pandas / MATLAB / 旧仪器/GBK)  
4. Check **预览**; use **整理预览** after editing the paste box  
5. Optional: **删除选中列** / **上移列** / **下移列** / **恢复原始数据**  
6. **导出** → CSV / TXT / TSV with UTF-8 BOM, UTF-8, or GBK  

**清空** clears the paste box and preview. Settings are stored at `%APPDATA%\CsvPasteExporter\settings.json`. Default: first row is data until **第一行是表头** is checked.

## Develop from source

Runtime is Python stdlib + Tkinter. Dev extras (pytest, PyInstaller) are in `requirements-dev.txt`:

```powershell
py -m pip install -r requirements-dev.txt
py csv_paste_exporter.py
py -m pytest -q
```

Build a windowed EXE:

```powershell
py -m PyInstaller --noconfirm --clean --onefile --windowed --name "CSV列数据整理导出" --distpath "dist" --workpath "build" --specpath "build" "csv_paste_exporter.py"
```

## License

MIT — see [LICENSE](LICENSE).
