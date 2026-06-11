from __future__ import annotations

import csv
import io
import json
import math
import os
import re
from collections import Counter
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable

import tkinter as tk
from tkinter import filedialog, messagebox, ttk


EXPLICIT_DELIMITERS = ("\t", ",", ";")
PREVIEW_ROW_LIMIT = 500
APP_DIR_NAME = "CsvPasteExporter"
EXPORT_FORMATS = {
    "csv": {
        "label": "CSV - Excel/Python 推荐",
        "delimiter": ",",
        "extension": ".csv",
        "filetype": "CSV 文件",
    },
    "txt": {
        "label": "TXT - Origin/仪器软件推荐",
        "delimiter": "\t",
        "extension": ".txt",
        "filetype": "TXT 文件",
    },
    "tsv": {
        "label": "TSV - 通用制表符表格",
        "delimiter": "\t",
        "extension": ".tsv",
        "filetype": "TSV 文件",
    },
}
ENCODINGS = {
    "UTF-8 BOM": "utf-8-sig",
    "UTF-8": "utf-8",
    "GBK": "gbk",
}
TARGET_PRESETS = {
    "excel": {
        "label": "Excel",
        "export_format": "csv",
        "encoding": "UTF-8 BOM",
    },
    "origin": {
        "label": "Origin",
        "export_format": "txt",
        "encoding": "UTF-8 BOM",
    },
    "pandas": {
        "label": "Python/pandas",
        "export_format": "csv",
        "encoding": "UTF-8",
    },
    "matlab": {
        "label": "MATLAB",
        "export_format": "csv",
        "encoding": "UTF-8",
    },
    "legacy_gbk": {
        "label": "旧仪器/GBK",
        "export_format": "txt",
        "encoding": "GBK",
    },
    "custom": {
        "label": "自定义",
        "export_format": "",
        "encoding": "",
    },
}
DEFAULT_SETTINGS = {
    "export_format": "csv",
    "encoding": "UTF-8 BOM",
    "last_export_dir": "",
    "first_row_is_header": False,
    "target_preset": "excel",
}


@dataclass(frozen=True)
class TableAnalysis:
    rows: list[list[str]]
    delimiter: str | None
    delimiter_name: str
    had_ragged_rows: bool
    empty_cell_count: int


@dataclass(frozen=True)
class ChartData:
    points: list[tuple[float, float]]
    x_range: tuple[float, float] | None
    y_range: tuple[float, float] | None
    skipped_rows: int


@dataclass(frozen=True)
class TableDiagnostics:
    row_count: int
    column_count: int
    empty_cell_count: int
    blank_heading_count: int
    duplicate_headings: tuple[str, ...]
    preview_truncated: bool
    requires_confirmation: bool
    warnings: tuple[str, ...]


def parse_table_text(text: str) -> list[list[str]]:
    """Parse pasted table text into a rectangular list of rows."""
    return analyze_table_text(text).rows


def analyze_table_text(text: str) -> TableAnalysis:
    if not text or not text.strip():
        return TableAnalysis([], None, "无", False, 0)

    delimiter = _choose_delimiter(text)
    if delimiter is None:
        rows = _parse_whitespace_table(text)
        delimiter_name = "空白"
    else:
        rows = _parse_delimited_table(text, delimiter)
        delimiter_name = _delimiter_name(delimiter)

    cleaned_rows, had_ragged_rows, empty_cell_count = _clean_table_with_metadata(rows)
    return TableAnalysis(
        cleaned_rows,
        delimiter,
        delimiter_name,
        had_ragged_rows,
        empty_cell_count,
    )


def write_csv(
    rows: list[list[str]],
    path: str | Path,
    encoding: str = "utf-8-sig",
) -> None:
    write_table(rows, path, delimiter=",", encoding=encoding)


def write_table(
    rows: list[list[str]],
    path: str | Path,
    delimiter: str,
    encoding: str = "utf-8-sig",
) -> None:
    with Path(path).open("w", encoding=encoding, newline="") as handle:
        writer = csv.writer(handle, delimiter=delimiter, lineterminator="\n")
        writer.writerows(rows)


def delete_columns(rows: list[list[str]], column_indices: set[int]) -> list[list[str]]:
    if not column_indices:
        return [row[:] for row in rows]

    return [
        [cell for index, cell in enumerate(row) if index not in column_indices]
        for row in rows
    ]


def move_column(rows: list[list[str]], column_index: int, direction: int) -> list[list[str]]:
    if not rows:
        return []

    column_count = max(len(row) for row in rows)
    target_index = column_index + direction
    if column_index < 0 or column_index >= column_count:
        return [row[:] for row in rows]
    if target_index < 0 or target_index >= column_count:
        return [row[:] for row in rows]

    moved_rows = []
    for row in rows:
        padded = row + [""] * (column_count - len(row))
        value = padded.pop(column_index)
        padded.insert(target_index, value)
        moved_rows.append(padded)
    return moved_rows


def get_preview_headings(rows: list[list[str]], first_row_is_header: bool) -> list[str]:
    if not rows:
        return []

    column_count = max(len(row) for row in rows)
    if not first_row_is_header:
        return [f"列 {index + 1}" for index in range(column_count)]

    first_row = rows[0] + [""] * (column_count - len(rows[0]))
    return [
        heading.strip() if heading.strip() else f"列 {index + 1}"
        for index, heading in enumerate(first_row)
    ]


def get_default_chart_column_indices(
    rows: list[list[str]],
) -> tuple[int | None, int | None]:
    column_count = max((len(row) for row in rows), default=0)
    if column_count < 2:
        return None, None
    return 0, 1


def get_chart_column_labels(
    rows: list[list[str]],
    first_row_is_header: bool,
) -> list[str]:
    return get_preview_headings(rows, first_row_is_header)


def build_column_list_labels(
    rows: list[list[str]],
    first_row_is_header: bool,
    sample_count: int = 3,
    max_sample_length: int = 18,
) -> list[str]:
    if not rows:
        return []

    column_count = max(len(row) for row in rows)
    headings = get_preview_headings(rows, first_row_is_header)
    data_rows = rows[1:] if first_row_is_header else rows
    labels = []
    for index in range(column_count):
        samples = []
        for row in data_rows:
            if index >= len(row):
                continue
            sample = _shorten_sample_text(row[index], max_sample_length)
            if sample:
                samples.append(sample)
            if len(samples) >= sample_count:
                break

        sample_suffix = f" | {', '.join(samples)}" if samples else " | 无样例"
        labels.append(f"{index + 1}. {headings[index]}{sample_suffix}")
    return labels


def calculate_axis_range(values: Iterable[float]) -> tuple[float, float] | None:
    numeric_values = list(values)
    if not numeric_values:
        return None

    minimum = min(numeric_values)
    maximum = max(numeric_values)
    if minimum != maximum:
        return minimum, maximum

    padding = abs(minimum) * 0.05
    if padding == 0:
        padding = 1.0
    return minimum - padding, maximum + padding


def zoom_axis_range(
    axis_range: tuple[float, float],
    anchor: float,
    scale: float,
    full_range: tuple[float, float] | None = None,
) -> tuple[float, float]:
    minimum, maximum = axis_range
    if scale <= 0 or minimum >= maximum:
        return axis_range

    new_minimum = anchor - (anchor - minimum) * scale
    new_maximum = anchor + (maximum - anchor) * scale
    if full_range is None:
        return new_minimum, new_maximum

    full_minimum, full_maximum = full_range
    full_width = full_maximum - full_minimum
    new_width = new_maximum - new_minimum
    if full_width <= 0 or new_width >= full_width:
        return full_range

    if new_minimum < full_minimum:
        shift = full_minimum - new_minimum
        new_minimum += shift
        new_maximum += shift
    if new_maximum > full_maximum:
        shift = new_maximum - full_maximum
        new_minimum -= shift
        new_maximum -= shift

    return max(new_minimum, full_minimum), min(new_maximum, full_maximum)


def build_chart_data(
    rows: list[list[str]],
    x_column: int,
    y_column: int,
    first_row_is_header: bool,
) -> ChartData:
    column_count = max((len(row) for row in rows), default=0)
    if (
        not rows
        or x_column < 0
        or y_column < 0
        or x_column >= column_count
        or y_column >= column_count
    ):
        return ChartData([], None, None, 0)

    data_rows = rows[1:] if first_row_is_header else rows
    points: list[tuple[float, float]] = []
    skipped_rows = 0
    for row in data_rows:
        try:
            x_text = row[x_column].strip()
            y_text = row[y_column].strip()
            x_value = float(x_text)
            y_value = float(y_text)
        except (IndexError, ValueError):
            skipped_rows += 1
            continue

        if not math.isfinite(x_value) or not math.isfinite(y_value):
            skipped_rows += 1
            continue
        points.append((x_value, y_value))

    return ChartData(
        points,
        calculate_axis_range(point[0] for point in points),
        calculate_axis_range(point[1] for point in points),
        skipped_rows,
    )


def build_table_diagnostics(
    rows: list[list[str]],
    analysis: TableAnalysis,
    first_row_is_header: bool,
    preview_row_limit: int = PREVIEW_ROW_LIMIT,
) -> TableDiagnostics:
    row_count = len(rows)
    column_count = max((len(row) for row in rows), default=0)
    empty_cell_count = sum(cell == "" for row in rows for cell in row)
    blank_heading_count = 0
    duplicate_headings: tuple[str, ...] = ()
    warnings = []

    if empty_cell_count:
        warnings.append(f"空单元格 {empty_cell_count} 个")
    if analysis.had_ragged_rows and empty_cell_count:
        warnings.append("行列不齐已补空")

    if first_row_is_header and rows:
        first_row = rows[0] + [""] * (column_count - len(rows[0]))
        headings = [heading.strip() for heading in first_row]
        blank_heading_count = sum(not heading for heading in headings)
        duplicate_counts = Counter(heading for heading in headings if heading)
        duplicate_headings = tuple(
            heading for heading, count in duplicate_counts.items() if count > 1
        )
        if blank_heading_count:
            warnings.append(f"存在空表头 {blank_heading_count} 个")
        if duplicate_headings:
            warnings.append(f"重复表头：{', '.join(duplicate_headings)}")

    return TableDiagnostics(
        row_count=row_count,
        column_count=column_count,
        empty_cell_count=empty_cell_count,
        blank_heading_count=blank_heading_count,
        duplicate_headings=duplicate_headings,
        preview_truncated=row_count > preview_row_limit,
        requires_confirmation=bool(warnings),
        warnings=tuple(warnings),
    )


def find_target_preset_key(export_format: str, encoding: str) -> str:
    for key, preset in TARGET_PRESETS.items():
        if key == "custom":
            continue
        if (
            preset["export_format"] == export_format
            and preset["encoding"] == encoding
        ):
            return key
    return "custom"


def build_default_filename(format_key: str, now: datetime | None = None) -> str:
    current_time = now or datetime.now()
    extension = EXPORT_FORMATS.get(format_key, EXPORT_FORMATS["csv"])["extension"]
    return f"整理数据_{current_time:%Y%m%d_%H%M%S}{extension}"


def get_settings_path() -> Path:
    appdata = os.environ.get("APPDATA")
    if appdata:
        return Path(appdata) / APP_DIR_NAME / "settings.json"
    return Path.home() / f".{APP_DIR_NAME}" / "settings.json"


def load_settings(path: str | Path | None = None) -> dict[str, object]:
    settings_path = Path(path) if path is not None else get_settings_path()
    settings = DEFAULT_SETTINGS.copy()
    try:
        with settings_path.open("r", encoding="utf-8") as handle:
            loaded = json.load(handle)
    except (OSError, json.JSONDecodeError):
        return settings

    if isinstance(loaded, dict):
        for key, default_value in DEFAULT_SETTINGS.items():
            value = loaded.get(key, default_value)
            if isinstance(value, type(default_value)):
                settings[key] = value
        if (
            "target_preset" not in loaded
            or settings["target_preset"] not in TARGET_PRESETS
        ):
            settings["target_preset"] = find_target_preset_key(
                str(settings["export_format"]),
                str(settings["encoding"]),
            )
    return settings


def save_settings(path: str | Path | None, settings: dict[str, object]) -> None:
    settings_path = Path(path) if path is not None else get_settings_path()
    settings_path.parent.mkdir(parents=True, exist_ok=True)
    cleaned = DEFAULT_SETTINGS.copy()
    for key, default_value in DEFAULT_SETTINGS.items():
        value = settings.get(key, default_value)
        if isinstance(value, type(default_value)):
            cleaned[key] = value
    with settings_path.open("w", encoding="utf-8") as handle:
        json.dump(cleaned, handle, ensure_ascii=False, indent=2)


def _shorten_sample_text(value: str, max_length: int) -> str:
    text = value.strip()
    if not text:
        return ""
    if max_length < 4:
        max_length = 4
    if len(text) <= max_length:
        return text
    return text[: max_length - 3] + "..."


def _copy_rows(rows: list[list[str]]) -> list[list[str]]:
    return [row[:] for row in rows]


def _choose_delimiter(text: str) -> str | None:
    scores = []
    for delimiter in EXPLICIT_DELIMITERS:
        score = _score_delimiter(text, delimiter)
        if score[0] > 0:
            scores.append((score, delimiter))

    if not scores:
        return None

    return max(scores, key=lambda item: item[0])[1]


def _delimiter_name(delimiter: str) -> str:
    names = {
        "\t": "Tab",
        ",": "逗号",
        ";": "分号",
    }
    return names.get(delimiter, delimiter)


def _score_delimiter(text: str, delimiter: str) -> tuple[int, int, int]:
    try:
        rows = list(csv.reader(io.StringIO(text), delimiter=delimiter, skipinitialspace=True))
    except csv.Error:
        return (0, 0, 0)

    counts = [len(row) for row in rows if any(cell.strip() for cell in row)]
    if not counts:
        return (0, 0, 0)

    multi_column_rows = sum(count > 1 for count in counts)
    most_common_width = Counter(counts).most_common(1)[0][1]
    widest_row = max(counts)
    return (multi_column_rows, most_common_width, widest_row)


def _parse_delimited_table(text: str, delimiter: str) -> list[list[str]]:
    reader = csv.reader(io.StringIO(text), delimiter=delimiter, skipinitialspace=True)
    return [[cell.strip() for cell in row] for row in reader]


def _parse_whitespace_table(text: str) -> list[list[str]]:
    rows: list[list[str]] = []
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped:
            rows.append([])
        else:
            rows.append(re.split(r"\s+", stripped))
    return rows


def _clean_table(rows: Iterable[list[str]]) -> list[list[str]]:
    return _clean_table_with_metadata(rows)[0]


def _clean_table_with_metadata(
    rows: Iterable[list[str]],
) -> tuple[list[list[str]], bool, int]:
    non_empty_rows = [
        [cell.strip() for cell in row]
        for row in rows
        if any(cell.strip() for cell in row)
    ]
    if not non_empty_rows:
        return [], False, 0

    original_widths = {len(row) for row in non_empty_rows}
    had_ragged_rows = len(original_widths) > 1
    width = max(len(row) for row in non_empty_rows)
    padded_rows = [row + [""] * (width - len(row)) for row in non_empty_rows]
    keep_columns = [
        index
        for index in range(width)
        if any(row[index] != "" for row in padded_rows)
    ]

    cleaned_rows = [[row[index] for index in keep_columns] for row in padded_rows]
    empty_cell_count = sum(cell == "" for row in cleaned_rows for cell in row)
    return cleaned_rows, had_ragged_rows, empty_cell_count


class CsvPasteExporterApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("CSV 列数据整理导出")
        self.root.geometry("1180x720")
        self.root.minsize(980, 600)

        self.settings_path = get_settings_path()
        self.settings = load_settings(self.settings_path)
        self.rows: list[list[str]] = []
        self.original_rows: list[list[str]] = []
        self.current_analysis = TableAnalysis([], None, "无", False, 0)
        self.preview_after_id: str | None = None
        self.text_dirty = False
        self.status_var = tk.StringVar(value="粘贴数据后会自动预览")
        target_preset = str(self.settings["target_preset"])
        if target_preset not in TARGET_PRESETS:
            target_preset = find_target_preset_key(
                str(self.settings["export_format"]),
                str(self.settings["encoding"]),
            )
        export_format = str(self.settings["export_format"])
        if target_preset != "custom":
            export_format = TARGET_PRESETS[target_preset]["export_format"]
        if export_format not in EXPORT_FORMATS:
            export_format = "csv"
        encoding = str(self.settings["encoding"])
        if target_preset != "custom":
            encoding = TARGET_PRESETS[target_preset]["encoding"]
        if encoding not in ENCODINGS:
            encoding = "UTF-8 BOM"
        self.target_preset_var = tk.StringVar(
            value=TARGET_PRESETS[target_preset]["label"]
        )
        self.export_format_var = tk.StringVar(value=EXPORT_FORMATS[export_format]["label"])
        self.encoding_var = tk.StringVar(value=encoding)
        self.first_row_is_header_var = tk.BooleanVar(
            value=bool(self.settings["first_row_is_header"])
        )
        self.chart_x_var = tk.StringVar(value="")
        self.chart_y_var = tk.StringVar(value="")
        self.chart_info_var = tk.StringVar(value="")
        self.chart_x_index: int | None = None
        self.chart_y_index: int | None = None
        self.chart_points: list[tuple[float, float]] = []
        self.chart_full_range: tuple[float, float, float, float] | None = None
        self.chart_view_range: tuple[float, float, float, float] | None = None
        self.chart_empty_message = "粘贴至少两列数值数据后显示图表"
        self.chart_drag_start: tuple[int, int] | None = None
        self.chart_drag_rect_id: int | None = None

        self._build_widgets()
        self._bind_events()

    def _build_widgets(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(2, weight=1)

        actions = ttk.Frame(self.root, padding=(10, 10, 10, 4))
        actions.grid(row=0, column=0, sticky="ew")
        actions.columnconfigure(0, weight=1)

        primary_actions = ttk.Frame(actions)
        primary_actions.grid(row=0, column=0, sticky="w")
        ttk.Button(primary_actions, text="粘贴剪贴板", command=self.paste_clipboard).grid(
            row=0, column=0, padx=(0, 8)
        )
        ttk.Button(primary_actions, text="清空", command=self.clear_all).grid(
            row=0, column=1, padx=(0, 8)
        )
        ttk.Button(primary_actions, text="整理预览", command=self.refresh_preview).grid(
            row=0, column=2, padx=(0, 8)
        )
        ttk.Button(primary_actions, text="导出", command=self.export_table).grid(
            row=0, column=3, padx=(0, 12)
        )
        ttk.Checkbutton(
            primary_actions,
            text="第一行是表头",
            variable=self.first_row_is_header_var,
            command=self._on_header_toggle,
        ).grid(row=0, column=4, padx=(0, 12))

        export_options = ttk.Frame(actions)
        export_options.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        export_options.columnconfigure(6, weight=1)

        ttk.Label(export_options, text="目标软件").grid(row=0, column=0, padx=(0, 5))
        target_box = ttk.Combobox(
            export_options,
            textvariable=self.target_preset_var,
            values=[config["label"] for config in TARGET_PRESETS.values()],
            state="readonly",
            width=14,
        )
        target_box.grid(row=0, column=1, padx=(0, 12))
        target_box.bind("<<ComboboxSelected>>", self._on_target_preset_changed)

        ttk.Label(export_options, text="格式").grid(row=0, column=2, padx=(0, 5))
        format_box = ttk.Combobox(
            export_options,
            textvariable=self.export_format_var,
            values=[config["label"] for config in EXPORT_FORMATS.values()],
            state="readonly",
            width=24,
        )
        format_box.grid(row=0, column=3, padx=(0, 12))
        format_box.bind("<<ComboboxSelected>>", self._on_export_option_changed)

        ttk.Label(export_options, text="编码").grid(row=0, column=4, padx=(0, 5))
        encoding_box = ttk.Combobox(
            export_options,
            textvariable=self.encoding_var,
            values=list(ENCODINGS),
            state="readonly",
            width=12,
        )
        encoding_box.grid(row=0, column=5, padx=(0, 12))
        encoding_box.bind("<<ComboboxSelected>>", self._on_export_option_changed)

        column_actions = ttk.Frame(self.root, padding=(10, 0, 10, 6))
        column_actions.grid(row=1, column=0, sticky="ew")
        ttk.Button(
            column_actions,
            text="删除选中列",
            command=self.delete_selected_columns,
        ).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(
            column_actions,
            text="上移列",
            command=lambda: self.move_selected_column(-1),
        ).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(
            column_actions,
            text="下移列",
            command=lambda: self.move_selected_column(1),
        ).grid(row=0, column=2, padx=(0, 8))
        ttk.Button(
            column_actions,
            text="恢复原始数据",
            command=self.restore_original_data,
        ).grid(row=0, column=3, padx=(0, 8))

        paned = ttk.PanedWindow(self.root, orient=tk.VERTICAL)
        paned.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 8))

        paste_frame = ttk.Labelframe(paned, text="粘贴区", padding=8)
        paste_frame.columnconfigure(0, weight=1)
        paste_frame.rowconfigure(0, weight=1)

        self.text = tk.Text(paste_frame, wrap="none", undo=True, height=10)
        text_y = ttk.Scrollbar(paste_frame, orient=tk.VERTICAL, command=self.text.yview)
        text_x = ttk.Scrollbar(paste_frame, orient=tk.HORIZONTAL, command=self.text.xview)
        self.text.configure(yscrollcommand=text_y.set, xscrollcommand=text_x.set)
        self.text.grid(row=0, column=0, sticky="nsew")
        text_y.grid(row=0, column=1, sticky="ns")
        text_x.grid(row=1, column=0, sticky="ew")

        preview_frame = ttk.Labelframe(paned, text="预览", padding=8)
        preview_frame.columnconfigure(1, weight=3)
        preview_frame.columnconfigure(3, weight=2)
        preview_frame.rowconfigure(0, weight=1)

        column_frame = ttk.Frame(preview_frame)
        column_frame.grid(row=0, column=0, sticky="ns", padx=(0, 8))
        column_frame.columnconfigure(0, weight=1)
        column_frame.rowconfigure(1, weight=1)
        ttk.Label(column_frame, text="列").grid(row=0, column=0, sticky="w")
        self.column_list = tk.Listbox(
            column_frame,
            selectmode=tk.EXTENDED,
            exportselection=False,
            width=38,
            height=12,
        )
        column_y = ttk.Scrollbar(column_frame, orient=tk.VERTICAL, command=self.column_list.yview)
        column_x = ttk.Scrollbar(
            column_frame,
            orient=tk.HORIZONTAL,
            command=self.column_list.xview,
        )
        self.column_list.configure(
            yscrollcommand=column_y.set,
            xscrollcommand=column_x.set,
        )
        self.column_list.grid(row=1, column=0, sticky="nsew")
        column_y.grid(row=1, column=1, sticky="ns")
        column_x.grid(row=2, column=0, sticky="ew")

        self.tree = ttk.Treeview(preview_frame, show="headings")
        tree_y = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.tree.yview)
        tree_x = ttk.Scrollbar(preview_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=tree_y.set, xscrollcommand=tree_x.set)
        self.tree.grid(row=0, column=1, sticky="nsew")
        tree_y.grid(row=0, column=2, sticky="ns")
        tree_x.grid(row=1, column=1, sticky="ew")

        chart_frame = ttk.Labelframe(preview_frame, text="快速图", padding=8)
        chart_frame.grid(row=0, column=3, rowspan=2, sticky="nsew", padx=(8, 0))
        chart_frame.columnconfigure(0, weight=1)
        chart_frame.rowconfigure(1, weight=1)

        chart_controls = ttk.Frame(chart_frame)
        chart_controls.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        chart_controls.columnconfigure(1, weight=1)
        chart_controls.columnconfigure(3, weight=1)
        ttk.Label(chart_controls, text="X").grid(row=0, column=0, padx=(0, 4))
        self.chart_x_box = ttk.Combobox(
            chart_controls,
            textvariable=self.chart_x_var,
            state="disabled",
            width=12,
        )
        self.chart_x_box.grid(row=0, column=1, sticky="ew", padx=(0, 8))
        ttk.Label(chart_controls, text="Y").grid(row=0, column=2, padx=(0, 4))
        self.chart_y_box = ttk.Combobox(
            chart_controls,
            textvariable=self.chart_y_var,
            state="disabled",
            width=12,
        )
        self.chart_y_box.grid(row=0, column=3, sticky="ew", padx=(0, 8))
        ttk.Button(
            chart_controls,
            text="自适应",
            command=self.reset_chart_view,
        ).grid(row=0, column=4)

        self.chart_canvas = tk.Canvas(
            chart_frame,
            background="white",
            highlightthickness=1,
            highlightbackground="#d0d0d0",
            width=320,
            height=240,
        )
        self.chart_canvas.grid(row=1, column=0, sticky="nsew")
        ttk.Label(
            chart_frame,
            textvariable=self.chart_info_var,
            anchor="w",
        ).grid(row=2, column=0, sticky="ew", pady=(6, 0))

        paned.add(paste_frame, weight=1)
        paned.add(preview_frame, weight=2)

        status = ttk.Label(
            self.root,
            textvariable=self.status_var,
            anchor="w",
            padding=(10, 0, 10, 8),
            justify="left",
            wraplength=1080,
        )
        status.grid(row=3, column=0, sticky="ew")
        self.status_label = status

    def _bind_events(self) -> None:
        self.text.bind("<KeyRelease>", self._schedule_preview)
        self.text.bind("<<Paste>>", self._schedule_preview)
        self.chart_x_box.bind("<<ComboboxSelected>>", self._on_chart_axis_changed)
        self.chart_y_box.bind("<<ComboboxSelected>>", self._on_chart_axis_changed)
        self.chart_canvas.bind("<Configure>", self._on_chart_canvas_configure)
        self.chart_canvas.bind("<MouseWheel>", self._on_chart_mousewheel)
        self.chart_canvas.bind("<ButtonPress-1>", self._on_chart_drag_start)
        self.chart_canvas.bind("<B1-Motion>", self._on_chart_drag_move)
        self.chart_canvas.bind("<ButtonRelease-1>", self._on_chart_drag_end)
        self.chart_canvas.bind("<Double-Button-1>", self.reset_chart_view)
        self.root.bind("<Configure>", self._on_root_configure, add="+")

    def _on_root_configure(self, event: tk.Event) -> None:
        if event.widget is self.root:
            self.status_label.configure(wraplength=max(420, event.width - 24))

    def _schedule_preview(self, _event: tk.Event | None = None) -> None:
        self.text_dirty = True
        if self.preview_after_id is not None:
            self.root.after_cancel(self.preview_after_id)
        self.preview_after_id = self.root.after(180, self.refresh_preview)

    def _on_target_preset_changed(self, _event: tk.Event | None = None) -> None:
        preset_key = self._current_target_preset_key()
        if preset_key != "custom":
            preset = TARGET_PRESETS[preset_key]
            self.export_format_var.set(
                EXPORT_FORMATS[preset["export_format"]]["label"]
            )
            self.encoding_var.set(preset["encoding"])
        self._save_current_settings()
        self._update_status()

    def _on_export_option_changed(self, _event: tk.Event | None = None) -> None:
        self._sync_target_preset_to_current_export_options()
        self._save_current_settings()
        self._update_status()

    def _on_header_toggle(self) -> None:
        self._save_current_settings()
        self._render_table(self.rows)
        self._update_status()

    def paste_clipboard(self) -> None:
        try:
            clipboard_text = self.root.clipboard_get()
        except tk.TclError:
            self.status_var.set("剪贴板里没有可读取的文本")
            return

        self.text.delete("1.0", tk.END)
        self.text.insert("1.0", clipboard_text)
        self.refresh_preview()

    def clear_all(self) -> None:
        self.text.delete("1.0", tk.END)
        self.rows = []
        self.original_rows = []
        self.current_analysis = TableAnalysis([], None, "无", False, 0)
        self.text_dirty = False
        self._render_table([], reset_chart_selection=True)
        self.status_var.set("已清空")

    def refresh_preview(self) -> None:
        if self.preview_after_id is not None:
            try:
                self.root.after_cancel(self.preview_after_id)
            except tk.TclError:
                pass
        self.preview_after_id = None
        text = self.text.get("1.0", tk.END)
        try:
            self.current_analysis = analyze_table_text(text)
        except csv.Error as exc:
            self.rows = []
            self.original_rows = []
            self.current_analysis = TableAnalysis([], None, "无", False, 0)
            self._render_table([], reset_chart_selection=True)
            self.status_var.set(f"解析失败：{exc}")
            return

        self.original_rows = _copy_rows(self.current_analysis.rows)
        self.rows = _copy_rows(self.current_analysis.rows)
        self.text_dirty = False
        self._render_table(self.rows, reset_chart_selection=True)
        self._update_status()

    def export_csv(self) -> None:
        self.export_table()

    def export_table(self) -> None:
        if self.text_dirty:
            self.refresh_preview()
        if not self.rows:
            messagebox.showwarning("没有数据", "请先粘贴需要导出的表格数据。")
            return

        diagnostics = self._current_table_diagnostics()
        if diagnostics.requires_confirmation and not self._confirm_risky_export(
            diagnostics
        ):
            return

        format_key = self._current_export_format_key()
        format_config = EXPORT_FORMATS[format_key]
        encoding_label = self.encoding_var.get()
        encoding = ENCODINGS.get(encoding_label, "utf-8-sig")
        last_export_dir = str(self.settings.get("last_export_dir", ""))
        initialdir = last_export_dir if last_export_dir and Path(last_export_dir).exists() else None

        path = filedialog.asksaveasfilename(
            title="导出表格",
            defaultextension=format_config["extension"],
            initialdir=initialdir,
            initialfile=build_default_filename(format_key),
            filetypes=[
                (format_config["filetype"], f"*{format_config['extension']}"),
                ("所有文件", "*.*"),
            ],
        )
        if not path:
            return

        try:
            write_table(
                self.rows,
                path,
                delimiter=str(format_config["delimiter"]),
                encoding=encoding,
            )
        except (OSError, UnicodeError) as exc:
            messagebox.showerror("导出失败", str(exc))
            self.status_var.set(f"导出失败：{exc}")
            return

        self._save_current_settings(last_export_dir=str(Path(path).parent))
        self.status_var.set(f"已导出：{path}")
        self._show_export_success(Path(path), diagnostics)

    def _confirm_risky_export(self, diagnostics: TableDiagnostics) -> bool:
        warning_text = "\n".join(f"- {warning}" for warning in diagnostics.warnings)
        return messagebox.askyesno(
            "导出前确认",
            "检测到以下需要确认的表格状态：\n\n"
            f"{warning_text}\n\n"
            "这些提示不会自动修改数据。仍然继续导出吗？",
        )

    def delete_selected_columns(self) -> None:
        selected = set(self.column_list.curselection())
        if not selected:
            messagebox.showinfo("未选择列", "请先在左侧列列表中选择需要删除的列。")
            return

        column_count = max((len(row) for row in self.rows), default=0)
        if len(selected) >= column_count:
            messagebox.showwarning("无法删除", "至少需要保留一列。")
            return

        self.rows = delete_columns(self.rows, selected)
        self._render_table(self.rows)
        self._update_status()

    def move_selected_column(self, direction: int) -> None:
        selected = list(self.column_list.curselection())
        if len(selected) != 1:
            messagebox.showinfo("选择一列", "请只选择一列进行移动。")
            return

        column_index = selected[0]
        column_count = max((len(row) for row in self.rows), default=0)
        target_index = column_index + direction
        if target_index < 0 or target_index >= column_count:
            return

        self.rows = move_column(self.rows, column_index, direction)
        self._render_table(self.rows, selected_indices={target_index})
        self._update_status()

    def restore_original_data(self) -> None:
        self.rows = _copy_rows(self.original_rows)
        self._render_table(self.rows)
        self._update_status()

    def _render_table(
        self,
        rows: list[list[str]],
        selected_indices: set[int] | None = None,
        reset_chart_selection: bool = False,
    ) -> None:
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = ()
        self.column_list.delete(0, tk.END)

        if not rows:
            self._update_chart_controls(reset_selection=True)
            return

        column_count = max(len(row) for row in rows)
        columns = [f"col_{index + 1}" for index in range(column_count)]
        self.tree["columns"] = columns
        headings = get_preview_headings(rows, self.first_row_is_header_var.get())
        column_labels = build_column_list_labels(
            rows,
            self.first_row_is_header_var.get(),
        )

        for index, column in enumerate(columns):
            heading = headings[index]
            self.tree.heading(column, text=heading)
            self.tree.column(column, width=120, minwidth=70, stretch=False)
            self.column_list.insert(tk.END, column_labels[index])

        for index in selected_indices or set():
            if 0 <= index < column_count:
                self.column_list.selection_set(index)

        display_rows = rows[1:] if self.first_row_is_header_var.get() else rows
        for row in display_rows[:PREVIEW_ROW_LIMIT]:
            values = row + [""] * (column_count - len(row))
            self.tree.insert("", tk.END, values=values)

        self._update_chart_controls(reset_selection=reset_chart_selection)

    def _update_chart_controls(self, reset_selection: bool = False) -> None:
        labels = get_chart_column_labels(self.rows, self.first_row_is_header_var.get())
        values = [f"{index + 1}. {label}" for index, label in enumerate(labels)]
        column_count = len(values)
        self.chart_x_box.configure(values=values)
        self.chart_y_box.configure(values=values)

        if column_count < 2:
            self.chart_x_index = None
            self.chart_y_index = None
            self.chart_x_var.set("")
            self.chart_y_var.set("")
            self.chart_x_box.configure(state="disabled")
            self.chart_y_box.configure(state="disabled")
            self.chart_info_var.set("")
            self._set_chart_empty("至少需要两列数据才能画图")
            return

        if reset_selection:
            self.chart_x_index, self.chart_y_index = get_default_chart_column_indices(
                self.rows
            )
        else:
            if self.chart_x_index is None or self.chart_x_index >= column_count:
                self.chart_x_index = 0
            if self.chart_y_index is None or self.chart_y_index >= column_count:
                self.chart_y_index = 1

        if self.chart_x_index is None or self.chart_y_index is None:
            self._set_chart_empty("至少需要两列数据才能画图")
            return

        self.chart_x_box.configure(state="readonly")
        self.chart_y_box.configure(state="readonly")
        self.chart_x_var.set(values[self.chart_x_index])
        self.chart_y_var.set(values[self.chart_y_index])
        self._refresh_chart(reset_view=True)

    def _on_chart_axis_changed(self, _event: tk.Event | None = None) -> None:
        x_index = self.chart_x_box.current()
        y_index = self.chart_y_box.current()
        if x_index >= 0:
            self.chart_x_index = x_index
        if y_index >= 0:
            self.chart_y_index = y_index
        self._refresh_chart(reset_view=True)

    def reset_chart_view(self, _event: tk.Event | None = None) -> str | None:
        if self.chart_full_range is None:
            return None

        self.chart_view_range = self.chart_full_range
        self._draw_chart()
        return "break"

    def _refresh_chart(self, reset_view: bool = False) -> None:
        if self.chart_x_index is None or self.chart_y_index is None:
            self.chart_points = []
            self.chart_full_range = None
            self.chart_view_range = None
            self._set_chart_empty("至少需要两列数据才能画图")
            return

        chart_data = build_chart_data(
            self.rows,
            self.chart_x_index,
            self.chart_y_index,
            self.first_row_is_header_var.get(),
        )
        self.chart_points = chart_data.points
        if (
            len(chart_data.points) < 2
            or chart_data.x_range is None
            or chart_data.y_range is None
        ):
            self.chart_full_range = None
            self.chart_view_range = None
            suffix = ""
            if chart_data.skipped_rows:
                suffix = f"，已跳过 {chart_data.skipped_rows} 行"
            self.chart_info_var.set("")
            self._set_chart_empty(f"有效数值点不足 2 个{suffix}")
            return

        self.chart_full_range = (
            chart_data.x_range[0],
            chart_data.x_range[1],
            chart_data.y_range[0],
            chart_data.y_range[1],
        )
        if reset_view or self.chart_view_range is None:
            self.chart_view_range = self.chart_full_range

        skipped = f"，跳过 {chart_data.skipped_rows} 行" if chart_data.skipped_rows else ""
        self.chart_info_var.set(f"{len(chart_data.points)} 个有效点{skipped}")
        self._draw_chart()

    def _set_chart_empty(self, message: str) -> None:
        self.chart_empty_message = message
        self.chart_points = []
        self.chart_full_range = None
        self.chart_view_range = None
        self._draw_chart_empty(message)

    def _chart_canvas_size(self) -> tuple[int, int]:
        width = self.chart_canvas.winfo_width()
        height = self.chart_canvas.winfo_height()
        if width <= 1:
            width = int(float(self.chart_canvas.cget("width")))
        if height <= 1:
            height = int(float(self.chart_canvas.cget("height")))
        return width, height

    def _chart_plot_bounds(self) -> tuple[int, int, int, int]:
        width, height = self._chart_canvas_size()
        left = 52
        top = 16
        right = max(left + 40, width - 16)
        bottom = max(top + 40, height - 38)
        return left, top, right, bottom

    def _draw_chart_empty(self, message: str) -> None:
        self.chart_canvas.delete("all")
        width, height = self._chart_canvas_size()
        self.chart_canvas.create_text(
            width / 2,
            height / 2,
            text=message,
            fill="#666666",
            width=max(120, width - 30),
            justify=tk.CENTER,
        )

    def _draw_chart(self) -> None:
        if self.chart_view_range is None or not self.chart_points:
            self._draw_chart_empty(self.chart_empty_message)
            return

        self.chart_canvas.delete("all")
        left, top, right, bottom = self._chart_plot_bounds()
        x_min, x_max, y_min, y_max = self.chart_view_range
        width = right - left
        height = bottom - top

        for index in range(6):
            t = index / 5
            x = left + width * t
            y = top + height * t
            self.chart_canvas.create_line(x, top, x, bottom, fill="#eeeeee")
            self.chart_canvas.create_line(left, y, right, y, fill="#eeeeee")
            x_value = x_min + (x_max - x_min) * t
            y_value = y_max - (y_max - y_min) * t
            self.chart_canvas.create_text(
                x,
                bottom + 7,
                text=self._format_axis_label(x_value),
                anchor="n",
                fill="#555555",
            )
            self.chart_canvas.create_text(
                left - 6,
                y,
                text=self._format_axis_label(y_value),
                anchor="e",
                fill="#555555",
            )

        visible_points = [
            self._chart_data_to_canvas(x_value, y_value)
            for x_value, y_value in self.chart_points
            if x_min <= x_value <= x_max and y_min <= y_value <= y_max
        ]
        if len(visible_points) > 2500:
            step = math.ceil(len(visible_points) / 2500)
            visible_points = visible_points[::step]

        if len(visible_points) >= 2:
            flattened = [coord for point in visible_points for coord in point]
            self.chart_canvas.create_line(
                *flattened,
                fill="#2563eb",
                width=2,
            )
        if len(visible_points) <= 800:
            for x, y in visible_points:
                self.chart_canvas.create_oval(
                    x - 2,
                    y - 2,
                    x + 2,
                    y + 2,
                    fill="#2563eb",
                    outline="",
                )

        self.chart_canvas.create_rectangle(left, top, right, bottom, outline="#777777")

    def _format_axis_label(self, value: float) -> str:
        return f"{value:.4g}"

    def _chart_data_to_canvas(self, x_value: float, y_value: float) -> tuple[float, float]:
        if self.chart_view_range is None:
            return 0.0, 0.0

        left, top, right, bottom = self._chart_plot_bounds()
        x_min, x_max, y_min, y_max = self.chart_view_range
        x = left + (x_value - x_min) / (x_max - x_min) * (right - left)
        y = top + (y_max - y_value) / (y_max - y_min) * (bottom - top)
        return x, y

    def _chart_canvas_to_data(self, x: int, y: int) -> tuple[float, float]:
        if self.chart_view_range is None:
            return 0.0, 0.0

        left, top, right, bottom = self._chart_plot_bounds()
        x_min, x_max, y_min, y_max = self.chart_view_range
        x_value = x_min + (x - left) / (right - left) * (x_max - x_min)
        y_value = y_max - (y - top) / (bottom - top) * (y_max - y_min)
        return x_value, y_value

    def _is_in_chart_plot(self, x: int, y: int) -> bool:
        left, top, right, bottom = self._chart_plot_bounds()
        return left <= x <= right and top <= y <= bottom

    def _clamp_chart_canvas_point(self, x: int, y: int) -> tuple[int, int]:
        left, top, right, bottom = self._chart_plot_bounds()
        return min(max(x, left), right), min(max(y, top), bottom)

    def _on_chart_canvas_configure(self, _event: tk.Event | None = None) -> None:
        if self.chart_view_range is None:
            self._draw_chart_empty(self.chart_empty_message)
            return
        self._draw_chart()

    def _on_chart_mousewheel(self, event: tk.Event) -> str | None:
        if (
            self.chart_view_range is None
            or self.chart_full_range is None
            or not self._is_in_chart_plot(event.x, event.y)
        ):
            return None

        x_anchor, y_anchor = self._chart_canvas_to_data(event.x, event.y)
        scale = 0.8 if event.delta > 0 else 1.25
        x_min, x_max, y_min, y_max = self.chart_view_range
        full_x_min, full_x_max, full_y_min, full_y_max = self.chart_full_range
        new_x_min, new_x_max = zoom_axis_range(
            (x_min, x_max),
            x_anchor,
            scale,
            (full_x_min, full_x_max),
        )
        new_y_min, new_y_max = zoom_axis_range(
            (y_min, y_max),
            y_anchor,
            scale,
            (full_y_min, full_y_max),
        )
        self.chart_view_range = new_x_min, new_x_max, new_y_min, new_y_max
        self._draw_chart()
        return "break"

    def _on_chart_drag_start(self, event: tk.Event) -> str | None:
        if self.chart_view_range is None or not self._is_in_chart_plot(event.x, event.y):
            return None

        self.chart_drag_start = self._clamp_chart_canvas_point(event.x, event.y)
        if self.chart_drag_rect_id is not None:
            self.chart_canvas.delete(self.chart_drag_rect_id)
        x, y = self.chart_drag_start
        self.chart_drag_rect_id = self.chart_canvas.create_rectangle(
            x,
            y,
            x,
            y,
            outline="#2563eb",
            dash=(3, 2),
        )
        return "break"

    def _on_chart_drag_move(self, event: tk.Event) -> str | None:
        if self.chart_drag_start is None or self.chart_drag_rect_id is None:
            return None

        x, y = self._clamp_chart_canvas_point(event.x, event.y)
        start_x, start_y = self.chart_drag_start
        self.chart_canvas.coords(self.chart_drag_rect_id, start_x, start_y, x, y)
        return "break"

    def _on_chart_drag_end(self, event: tk.Event) -> str | None:
        if self.chart_drag_start is None:
            return None

        start_x, start_y = self.chart_drag_start
        end_x, end_y = self._clamp_chart_canvas_point(event.x, event.y)
        if self.chart_drag_rect_id is not None:
            self.chart_canvas.delete(self.chart_drag_rect_id)
        self.chart_drag_start = None
        self.chart_drag_rect_id = None

        if abs(end_x - start_x) < 6 or abs(end_y - start_y) < 6:
            self._draw_chart()
            return "break"

        x1, y1 = self._chart_canvas_to_data(start_x, start_y)
        x2, y2 = self._chart_canvas_to_data(end_x, end_y)
        self.chart_view_range = (
            min(x1, x2),
            max(x1, x2),
            min(y1, y2),
            max(y1, y2),
        )
        self._draw_chart()
        return "break"

    def _current_table_diagnostics(self) -> TableDiagnostics:
        return build_table_diagnostics(
            self.rows,
            self.current_analysis,
            self.first_row_is_header_var.get(),
        )

    def _update_status(self) -> None:
        if not self.rows:
            self.status_var.set("未检测到有效表格数据")
            return

        diagnostics = self._current_table_diagnostics()
        format_key = self._current_export_format_key()
        format_name = EXPORT_FORMATS[format_key]["label"].split(" - ")[0]
        target_name = TARGET_PRESETS[self._current_target_preset_key()]["label"]
        health_label = "需确认" if diagnostics.requires_confirmation else "可导出"
        parts = [
            health_label,
            f"共 {diagnostics.row_count} 行、{diagnostics.column_count} 列",
            f"分隔符 {self.current_analysis.delimiter_name}",
            f"空单元格 {diagnostics.empty_cell_count}",
            f"目标 {target_name}",
            f"导出 {format_name}/{self.encoding_var.get()}",
        ]
        status_warnings = [
            warning
            for warning in diagnostics.warnings
            if not warning.startswith("空单元格")
        ]
        parts.extend(status_warnings[:3])
        if diagnostics.preview_truncated:
            parts.append(f"仅预览前 {PREVIEW_ROW_LIMIT} 行，导出包含全部数据")
        self.status_var.set("；".join(parts))

    def _current_export_format_key(self) -> str:
        label = self.export_format_var.get()
        for key, config in EXPORT_FORMATS.items():
            if label == config["label"]:
                return key
        return "csv"

    def _current_target_preset_key(self) -> str:
        label = self.target_preset_var.get()
        for key, config in TARGET_PRESETS.items():
            if label == config["label"]:
                return key
        return "custom"

    def _sync_target_preset_to_current_export_options(self) -> None:
        preset_key = find_target_preset_key(
            self._current_export_format_key(),
            self.encoding_var.get(),
        )
        self.target_preset_var.set(TARGET_PRESETS[preset_key]["label"])

    def _save_current_settings(self, last_export_dir: str | None = None) -> None:
        settings = {
            "export_format": self._current_export_format_key(),
            "encoding": self.encoding_var.get()
            if self.encoding_var.get() in ENCODINGS
            else "UTF-8 BOM",
            "last_export_dir": last_export_dir
            if last_export_dir is not None
            else str(self.settings.get("last_export_dir", "")),
            "first_row_is_header": self.first_row_is_header_var.get(),
            "target_preset": self._current_target_preset_key(),
        }
        self.settings = settings
        try:
            save_settings(self.settings_path, settings)
        except OSError as exc:
            self.status_var.set(f"设置保存失败：{exc}")

    def _show_export_success(self, path: Path, diagnostics: TableDiagnostics) -> None:
        dialog = tk.Toplevel(self.root)
        dialog.title("导出成功")
        dialog.transient(self.root)
        dialog.resizable(False, False)

        content = ttk.Frame(dialog, padding=14)
        content.grid(row=0, column=0, sticky="nsew")
        format_key = self._current_export_format_key()
        format_name = EXPORT_FORMATS[format_key]["label"].split(" - ")[0]
        target_name = TARGET_PRESETS[self._current_target_preset_key()]["label"]
        summary = (
            f"已导出：{path}\n"
            f"目标软件：{target_name}\n"
            f"格式/编码：{format_name} / {self.encoding_var.get()}\n"
            f"数据规模：{diagnostics.row_count} 行、{diagnostics.column_count} 列"
        )
        ttk.Label(content, text=summary, justify="left", wraplength=720).grid(
            row=0,
            column=0,
            columnspan=2,
            sticky="w",
            pady=(0, 12),
        )
        ttk.Button(
            content,
            text="打开文件夹",
            command=lambda: self._open_export_folder(path),
        ).grid(row=1, column=0, padx=(0, 8), sticky="e")
        ttk.Button(content, text="关闭", command=dialog.destroy).grid(
            row=1,
            column=1,
            sticky="w",
        )

    def _open_export_folder(self, path: Path) -> None:
        try:
            os.startfile(path.parent)
        except OSError as exc:
            messagebox.showerror("打开失败", str(exc))


def main() -> None:
    root = tk.Tk()
    app = CsvPasteExporterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
