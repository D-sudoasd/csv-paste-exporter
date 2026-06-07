from __future__ import annotations

import csv
import io
import json
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
DEFAULT_SETTINGS = {
    "export_format": "csv",
    "encoding": "UTF-8 BOM",
    "last_export_dir": "",
    "first_row_is_header": False,
}


@dataclass(frozen=True)
class TableAnalysis:
    rows: list[list[str]]
    delimiter: str | None
    delimiter_name: str
    had_ragged_rows: bool
    empty_cell_count: int


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
        self.root.geometry("980x700")
        self.root.minsize(760, 520)

        self.settings_path = get_settings_path()
        self.settings = load_settings(self.settings_path)
        self.rows: list[list[str]] = []
        self.original_rows: list[list[str]] = []
        self.current_analysis = TableAnalysis([], None, "无", False, 0)
        self.preview_after_id: str | None = None
        self.text_dirty = False
        self.status_var = tk.StringVar(value="粘贴数据后会自动预览")
        export_format = str(self.settings["export_format"])
        if export_format not in EXPORT_FORMATS:
            export_format = "csv"
        encoding = str(self.settings["encoding"])
        if encoding not in ENCODINGS:
            encoding = "UTF-8 BOM"
        self.export_format_var = tk.StringVar(value=EXPORT_FORMATS[export_format]["label"])
        self.encoding_var = tk.StringVar(value=encoding)
        self.first_row_is_header_var = tk.BooleanVar(
            value=bool(self.settings["first_row_is_header"])
        )

        self._build_widgets()
        self._bind_events()

    def _build_widgets(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(2, weight=1)

        actions = ttk.Frame(self.root, padding=(10, 10, 10, 4))
        actions.grid(row=0, column=0, sticky="ew")
        actions.columnconfigure(12, weight=1)

        ttk.Button(actions, text="粘贴剪贴板", command=self.paste_clipboard).grid(
            row=0, column=0, padx=(0, 8)
        )
        ttk.Button(actions, text="清空", command=self.clear_all).grid(
            row=0, column=1, padx=(0, 8)
        )
        ttk.Button(actions, text="整理预览", command=self.refresh_preview).grid(
            row=0, column=2, padx=(0, 8)
        )
        ttk.Button(actions, text="导出", command=self.export_table).grid(
            row=0, column=3, padx=(0, 18)
        )
        ttk.Label(actions, text="格式").grid(row=0, column=4, padx=(0, 5))
        format_box = ttk.Combobox(
            actions,
            textvariable=self.export_format_var,
            values=[config["label"] for config in EXPORT_FORMATS.values()],
            state="readonly",
            width=24,
        )
        format_box.grid(row=0, column=5, padx=(0, 12))
        format_box.bind("<<ComboboxSelected>>", self._on_export_option_changed)

        ttk.Label(actions, text="编码").grid(row=0, column=6, padx=(0, 5))
        encoding_box = ttk.Combobox(
            actions,
            textvariable=self.encoding_var,
            values=list(ENCODINGS),
            state="readonly",
            width=12,
        )
        encoding_box.grid(row=0, column=7, padx=(0, 12))
        encoding_box.bind("<<ComboboxSelected>>", self._on_export_option_changed)
        ttk.Checkbutton(
            actions,
            text="第一行是表头",
            variable=self.first_row_is_header_var,
            command=self._on_header_toggle,
        ).grid(row=0, column=8, padx=(0, 12))

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
        preview_frame.columnconfigure(1, weight=1)
        preview_frame.rowconfigure(0, weight=1)

        column_frame = ttk.Frame(preview_frame)
        column_frame.grid(row=0, column=0, sticky="ns", padx=(0, 8))
        column_frame.rowconfigure(1, weight=1)
        ttk.Label(column_frame, text="列").grid(row=0, column=0, sticky="w")
        self.column_list = tk.Listbox(
            column_frame,
            selectmode=tk.EXTENDED,
            exportselection=False,
            width=24,
            height=12,
        )
        column_y = ttk.Scrollbar(column_frame, orient=tk.VERTICAL, command=self.column_list.yview)
        self.column_list.configure(yscrollcommand=column_y.set)
        self.column_list.grid(row=1, column=0, sticky="ns")
        column_y.grid(row=1, column=1, sticky="ns")

        self.tree = ttk.Treeview(preview_frame, show="headings")
        tree_y = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.tree.yview)
        tree_x = ttk.Scrollbar(preview_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=tree_y.set, xscrollcommand=tree_x.set)
        self.tree.grid(row=0, column=1, sticky="nsew")
        tree_y.grid(row=0, column=2, sticky="ns")
        tree_x.grid(row=1, column=1, sticky="ew")

        paned.add(paste_frame, weight=1)
        paned.add(preview_frame, weight=2)

        status = ttk.Label(
            self.root,
            textvariable=self.status_var,
            anchor="w",
            padding=(10, 0, 10, 8),
        )
        status.grid(row=3, column=0, sticky="ew")

    def _bind_events(self) -> None:
        self.text.bind("<KeyRelease>", self._schedule_preview)
        self.text.bind("<<Paste>>", self._schedule_preview)

    def _schedule_preview(self, _event: tk.Event | None = None) -> None:
        self.text_dirty = True
        if self.preview_after_id is not None:
            self.root.after_cancel(self.preview_after_id)
        self.preview_after_id = self.root.after(180, self.refresh_preview)

    def _on_export_option_changed(self, _event: tk.Event | None = None) -> None:
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
        self._render_table([])
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
            self._render_table([])
            self.status_var.set(f"解析失败：{exc}")
            return

        self.original_rows = _copy_rows(self.current_analysis.rows)
        self.rows = _copy_rows(self.current_analysis.rows)
        self.text_dirty = False
        self._render_table(self.rows)
        self._update_status()

    def export_csv(self) -> None:
        self.export_table()

    def export_table(self) -> None:
        if self.text_dirty:
            self.refresh_preview()
        if not self.rows:
            messagebox.showwarning("没有数据", "请先粘贴需要导出的表格数据。")
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
        self._show_export_success(Path(path))

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
    ) -> None:
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = ()
        self.column_list.delete(0, tk.END)

        if not rows:
            return

        column_count = max(len(row) for row in rows)
        columns = [f"col_{index + 1}" for index in range(column_count)]
        self.tree["columns"] = columns
        headings = get_preview_headings(rows, self.first_row_is_header_var.get())

        for index, column in enumerate(columns):
            heading = headings[index]
            self.tree.heading(column, text=heading)
            self.tree.column(column, width=120, minwidth=70, stretch=False)
            self.column_list.insert(tk.END, f"{index + 1}. {heading}")

        for index in selected_indices or set():
            if 0 <= index < column_count:
                self.column_list.selection_set(index)

        display_rows = rows[1:] if self.first_row_is_header_var.get() else rows
        for row in display_rows[:PREVIEW_ROW_LIMIT]:
            values = row + [""] * (column_count - len(row))
            self.tree.insert("", tk.END, values=values)

    def _update_status(self) -> None:
        if not self.rows:
            self.status_var.set("未检测到有效表格数据")
            return

        row_count = len(self.rows)
        column_count = max((len(row) for row in self.rows), default=0)
        empty_cell_count = sum(cell == "" for row in self.rows for cell in row)
        format_key = self._current_export_format_key()
        format_name = EXPORT_FORMATS[format_key]["label"].split(" - ")[0]
        parts = [
            f"共 {row_count} 行、{column_count} 列",
            f"空单元格 {empty_cell_count}",
            f"分隔符 {self.current_analysis.delimiter_name}",
            f"导出 {format_name}",
        ]
        if self.current_analysis.had_ragged_rows:
            parts.append("行列不齐已补空")
        suffix = ""
        if row_count > PREVIEW_ROW_LIMIT:
            suffix = f"；仅预览前 {PREVIEW_ROW_LIMIT} 行，导出包含全部数据"
        self.status_var.set("；".join(parts) + suffix)

    def _current_export_format_key(self) -> str:
        label = self.export_format_var.get()
        for key, config in EXPORT_FORMATS.items():
            if label == config["label"]:
                return key
        return "csv"

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
        }
        self.settings = settings
        try:
            save_settings(self.settings_path, settings)
        except OSError as exc:
            self.status_var.set(f"设置保存失败：{exc}")

    def _show_export_success(self, path: Path) -> None:
        dialog = tk.Toplevel(self.root)
        dialog.title("导出成功")
        dialog.transient(self.root)
        dialog.resizable(False, False)

        content = ttk.Frame(dialog, padding=14)
        content.grid(row=0, column=0, sticky="nsew")
        ttk.Label(content, text=f"已导出：{path}").grid(
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
