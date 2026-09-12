import csv
import json
from datetime import datetime

import pytest
import tkinter as tk

from csv_paste_exporter import (
    PREVIEW_ROW_LIMIT,
    TARGET_PRESETS,
    CsvPasteExporterApp,
    analyze_table_text,
    build_chart_data,
    build_column_list_labels,
    build_default_filename,
    build_table_diagnostics,
    calculate_axis_range,
    delete_columns,
    find_target_preset_key,
    remap_column_index_after_delete,
    remap_column_index_after_move,
    get_chart_column_labels,
    get_default_chart_column_indices,
    get_preview_headings,
    resolve_chart_column_indices,
    load_settings,
    move_column,
    parse_table_text,
    save_settings,
    write_csv,
    write_table,
    zoom_axis_range,
)


def test_parse_tab_delimited_clipboard_text_preserves_headers_and_values():
    text = "Strain\tStress MPa\tNote\n0\t0\tstart\n1.2E-3\t345.6\tpeak\n"

    rows = parse_table_text(text)

    assert rows == [
        ["Strain", "Stress MPa", "Note"],
        ["0", "0", "start"],
        ["1.2E-3", "345.6", "peak"],
    ]


def test_parse_csv_text_respects_quoted_commas():
    text = 'Name,Value,Comment\n"alpha,beta",12.5,"kept, together"\n'

    rows = parse_table_text(text)

    assert rows == [
        ["Name", "Value", "Comment"],
        ["alpha,beta", "12.5", "kept, together"],
    ]


def test_parse_removes_empty_rows_and_empty_columns():
    text = "\tA\t\tB\t\n\t1\t\t2\t\n\t\t\t\t\n\t3\t\t4\t\n"

    rows = parse_table_text(text)

    assert rows == [
        ["A", "B"],
        ["1", "2"],
        ["3", "4"],
    ]


def test_parse_pads_ragged_rows_after_cleanup():
    text = "A\tB\tC\n1\t2\n3\t4\t5\n"

    rows = parse_table_text(text)

    assert rows == [
        ["A", "B", "C"],
        ["1", "2", ""],
        ["3", "4", "5"],
    ]


def test_parse_prefers_tab_when_cells_contain_thousands_commas():
    text = "1,234.5\t2,345.6\n3,456.7\t4,567.8\n"

    analysis = analyze_table_text(text)

    assert analysis.delimiter_name == "Tab"
    assert analysis.rows == [
        ["1,234.5", "2,345.6"],
        ["3,456.7", "4,567.8"],
    ]


def test_parse_prefers_semicolon_when_cells_contain_decimal_commas():
    text = "1,5;2,3\n3,7;4,1\n"

    analysis = analyze_table_text(text)

    assert analysis.delimiter_name == "分号"
    assert analysis.rows == [
        ["1,5", "2,3"],
        ["3,7", "4,1"],
    ]


def test_parse_strips_nul_bytes_from_pasted_text():
    assert parse_table_text("A\tB\x00\n1\t2\n") == [["A", "B"], ["1", "2"]]


def test_parse_semicolon_delimited_text():
    text = "A;B;C\n1;2;3\n4;5;6\n"

    rows = parse_table_text(text)

    assert rows == [
        ["A", "B", "C"],
        ["1", "2", "3"],
        ["4", "5", "6"],
    ]


def test_parse_strips_utf8_bom_from_pasted_text():
    text = "\ufeffStrain\tStress\n0\t1\n"

    rows = parse_table_text(text)

    assert rows == [
        ["Strain", "Stress"],
        ["0", "1"],
    ]
    assert analyze_table_text("\ufeff").rows == []


def test_parse_empty_or_whitespace_text_returns_no_rows():
    assert parse_table_text("") == []
    assert parse_table_text("   \n\n") == []
    assert analyze_table_text("\n").delimiter_name == "无"


def test_parse_falls_back_to_whitespace_when_no_explicit_delimiter():
    text = "Q angle intensity\n10.0 120 5.5E+2\n20.0 240 6.0E+2\n"

    rows = parse_table_text(text)

    assert rows == [
        ["Q", "angle", "intensity"],
        ["10.0", "120", "5.5E+2"],
        ["20.0", "240", "6.0E+2"],
    ]


def test_write_csv_uses_utf8_bom_and_standard_csv_quoting(tmp_path):
    output = tmp_path / "out.csv"
    rows = [["应变", "备注"], ["0.01", "alpha,beta"]]

    write_csv(rows, output)

    data = output.read_bytes()
    assert data.startswith(b"\xef\xbb\xbf")
    with output.open("r", encoding="utf-8-sig", newline="") as handle:
        assert list(csv.reader(handle)) == rows


def test_write_table_exports_tab_delimited_txt_and_tsv(tmp_path):
    rows = [["角度", "强度"], ["10.0", "5.5E+2"]]
    txt_output = tmp_path / "out.txt"
    tsv_output = tmp_path / "out.tsv"

    write_table(rows, txt_output, delimiter="\t", encoding="utf-8-sig")
    write_table(rows, tsv_output, delimiter="\t", encoding="utf-8-sig")

    assert txt_output.read_text(encoding="utf-8-sig").splitlines() == [
        "角度\t强度",
        "10.0\t5.5E+2",
    ]
    assert tsv_output.read_text(encoding="utf-8-sig").splitlines() == [
        "角度\t强度",
        "10.0\t5.5E+2",
    ]


def test_write_table_utf8_without_bom(tmp_path):
    output = tmp_path / "out.csv"
    rows = [["应变", "值"], ["0.1", "1"]]

    write_table(rows, output, delimiter=",", encoding="utf-8")

    data = output.read_bytes()
    assert not data.startswith(b"\xef\xbb\xbf")
    assert data.decode("utf-8").splitlines() == ["应变,值", "0.1,1"]


def test_write_table_supports_gbk_encoding(tmp_path):
    output = tmp_path / "out.txt"
    rows = [["应力", "备注"], ["12.3", "中文"]]

    write_table(rows, output, delimiter="\t", encoding="gbk")

    assert output.read_bytes().decode("gbk").splitlines() == [
        "应力\t备注",
        "12.3\t中文",
    ]


def test_analyze_table_text_reports_delimiter_ragged_rows_and_empty_cells():
    text = "A\tB\tC\n1\t2\n3\t4\t5\n"

    analysis = analyze_table_text(text)

    assert analysis.rows == [
        ["A", "B", "C"],
        ["1", "2", ""],
        ["3", "4", "5"],
    ]
    assert analysis.delimiter_name == "Tab"
    assert analysis.had_ragged_rows is True
    assert analysis.empty_cell_count == 1


def test_column_operations_delete_and_move_selected_columns():
    rows = [["A", "B", "C"], ["1", "2", "3"], ["4", "5", "6"]]

    assert delete_columns(rows, {1}) == [["A", "C"], ["1", "3"], ["4", "6"]]
    assert move_column(rows, 2, -1) == [["A", "C", "B"], ["1", "3", "2"], ["4", "6", "5"]]
    assert move_column(rows, 0, -1) == rows


def test_delete_columns_drops_rows_that_become_entirely_empty():
    rows = [["A", "B", "C"], ["1", "2", "x"], ["", "", "onlyC"]]

    assert delete_columns(rows, {2}) == [["A", "B"], ["1", "2"]]


def test_chart_column_indices_remap_after_delete_and_move():
    assert remap_column_index_after_delete(1, {0}) == 0
    assert remap_column_index_after_delete(2, {0}) == 1
    assert remap_column_index_after_delete(0, {0}) is None
    assert remap_column_index_after_move(2, 0, 1) == 2
    assert remap_column_index_after_move(1, 0, 1) == 0
    assert remap_column_index_after_move(0, 0, 1) == 1


def test_resolve_chart_column_indices_avoids_plotting_a_column_against_itself():
    assert resolve_chart_column_indices(None, 0, 2) == (0, 1)
    assert resolve_chart_column_indices(0, None, 3) == (0, 1)
    assert resolve_chart_column_indices(0, 0, 2) == (0, 1)
    assert resolve_chart_column_indices(1, 1, 3) == (1, 0)
    assert resolve_chart_column_indices(None, None, 1) == (None, None)


def test_preview_headings_use_first_row_only_for_display():
    rows = [["应变", "应力"], ["0.01", "100"]]

    assert get_preview_headings(rows, first_row_is_header=True) == ["应变", "应力"]
    assert get_preview_headings(rows, first_row_is_header=False) == ["列 1", "列 2"]
    assert rows == [["应变", "应力"], ["0.01", "100"]]


def test_chart_defaults_to_first_two_columns_and_uses_header_labels():
    rows = [["Time", "Force", "Note"], ["0", "10", "start"], ["1", "12", "end"]]

    assert get_default_chart_column_indices(rows) == (0, 1)
    assert get_chart_column_labels(rows, first_row_is_header=True) == [
        "Time",
        "Force",
        "Note",
    ]

    chart_data = build_chart_data(rows, 0, 1, first_row_is_header=True)

    assert chart_data.points == [(0.0, 10.0), (1.0, 12.0)]
    assert chart_data.x_range == (0.0, 1.0)
    assert chart_data.y_range == (10.0, 12.0)
    assert chart_data.skipped_rows == 0


def test_chart_data_skips_non_numeric_rows_and_supports_scientific_notation():
    rows = [
        ["X", "Y"],
        ["0", "1"],
        ["bad", "2"],
        ["1.2E-3", "3.4E+2"],
        ["2", "NaN"],
    ]

    chart_data = build_chart_data(rows, 0, 1, first_row_is_header=True)

    assert chart_data.points == [(0.0, 1.0), (0.0012, 340.0)]
    assert chart_data.x_range == (0.0, 0.0012)
    assert chart_data.y_range == (1.0, 340.0)
    assert chart_data.skipped_rows == 2


def test_chart_data_skips_inf_and_negative_inf():
    rows = [
        ["X", "Y"],
        ["0", "1"],
        ["1", "inf"],
        ["2", "-inf"],
        ["3", "Infinity"],
        ["4", "2"],
    ]

    chart_data = build_chart_data(rows, 0, 1, first_row_is_header=True)

    assert chart_data.points == [(0.0, 1.0), (4.0, 2.0)]
    assert chart_data.skipped_rows == 3


def test_chart_axis_range_pads_constant_values():
    assert calculate_axis_range([10.0, 10.0, 10.0]) == (9.5, 10.5)
    assert calculate_axis_range([0.0, 0.0]) == (-1.0, 1.0)

    chart_data = build_chart_data(
        [["X", "Y"], ["5", "0"], ["5", "0"]],
        0,
        1,
        first_row_is_header=True,
    )

    assert chart_data.x_range == (4.75, 5.25)
    assert chart_data.y_range == (-1.0, 1.0)


def test_chart_helpers_disable_plot_when_table_has_fewer_than_two_columns():
    rows = [["A"], ["1"], ["2"]]

    assert get_default_chart_column_indices(rows) == (None, None)

    chart_data = build_chart_data(rows, 0, 1, first_row_is_header=False)

    assert chart_data.points == []
    assert chart_data.x_range is None
    assert chart_data.y_range is None


def test_zoom_axis_range_scales_around_anchor_and_clamps_to_full_range():
    assert zoom_axis_range((0.0, 10.0), anchor=5.0, scale=0.5) == (2.5, 7.5)
    assert zoom_axis_range(
        (2.0, 8.0),
        anchor=5.0,
        scale=2.0,
        full_range=(0.0, 10.0),
    ) == (0.0, 10.0)


def test_target_presets_map_to_existing_formats_and_encodings():
    assert TARGET_PRESETS["excel"]["export_format"] == "csv"
    assert TARGET_PRESETS["excel"]["encoding"] == "UTF-8 BOM"
    assert TARGET_PRESETS["origin"]["export_format"] == "txt"
    assert TARGET_PRESETS["legacy_gbk"]["encoding"] == "GBK"

    assert find_target_preset_key("csv", "UTF-8 BOM") == "excel"
    assert find_target_preset_key("txt", "GBK") == "legacy_gbk"
    assert find_target_preset_key("tsv", "GBK") == "custom"
    assert find_target_preset_key("csv", "UTF-8") == "pandas"
    assert find_target_preset_key("csv", "UTF-8", current_key="matlab") == "matlab"
    assert find_target_preset_key("csv", "UTF-8", current_key="pandas") == "pandas"


def test_table_diagnostics_identifies_confirmation_risks():
    analysis = analyze_table_text("A\tA\t\n1\t2\t\n3\t\t5\n")

    diagnostics = build_table_diagnostics(
        analysis.rows,
        analysis,
        first_row_is_header=True,
    )

    assert diagnostics.requires_confirmation is True
    assert diagnostics.row_count == 3
    assert diagnostics.column_count == 3
    assert diagnostics.empty_cell_count == 3
    assert diagnostics.blank_heading_count == 1
    assert diagnostics.duplicate_headings == ("A",)
    assert "空单元格 3 个" in diagnostics.warnings
    assert "存在空表头 1 个" in diagnostics.warnings
    assert "重复表头：A" in diagnostics.warnings


def test_table_diagnostics_marks_preview_truncation_above_limit():
    rows = [["X", "Y"]] + [[str(index), "1"] for index in range(PREVIEW_ROW_LIMIT)]
    analysis = analyze_table_text("X\tY\n0\t1\n")

    diagnostics = build_table_diagnostics(
        rows,
        analysis,
        first_row_is_header=True,
    )

    assert len(rows) == PREVIEW_ROW_LIMIT + 1
    assert diagnostics.preview_truncated is True
    assert diagnostics.row_count == PREVIEW_ROW_LIMIT + 1


def test_table_diagnostics_marks_clean_tables_as_ready():
    analysis = analyze_table_text("Strain\tStress\n0\t0\n0.1\t100\n")

    diagnostics = build_table_diagnostics(
        analysis.rows,
        analysis,
        first_row_is_header=True,
    )

    assert diagnostics.requires_confirmation is False
    assert diagnostics.warnings == ()


def test_table_diagnostics_does_not_keep_stale_ragged_warning_after_column_delete():
    analysis = analyze_table_text("A\tB\tC\n1\t2\n3\t4\t5\n")
    rows_without_padding = delete_columns(analysis.rows, {2})

    diagnostics = build_table_diagnostics(
        rows_without_padding,
        analysis,
        first_row_is_header=True,
    )

    assert diagnostics.empty_cell_count == 0
    assert "行列不齐已补空" not in diagnostics.warnings
    assert diagnostics.requires_confirmation is False


def test_column_list_labels_include_sample_values_and_truncate_long_text():
    rows = [
        ["Strain", "Stress MPa", "Comment"],
        ["0", "100", "short"],
        ["0.01", "200", "this-is-a-very-long-value"],
        ["0.02", "250", "ignored"],
    ]

    labels = build_column_list_labels(
        rows,
        first_row_is_header=True,
        sample_count=2,
        max_sample_length=10,
    )

    assert labels == [
        "1. Strain | 0, 0.01",
        "2. Stress MPa | 100, 200",
        "3. Comment | short, this-is...",
    ]


def test_settings_round_trip_with_defaults(tmp_path):
    settings_path = tmp_path / "settings.json"

    assert load_settings(settings_path) == {
        "export_format": "csv",
        "encoding": "UTF-8 BOM",
        "last_export_dir": "",
        "first_row_is_header": False,
        "target_preset": "excel",
    }

    save_settings(
        settings_path,
        {
            "export_format": "txt",
            "encoding": "GBK",
            "last_export_dir": "D:/data",
            "first_row_is_header": True,
            "target_preset": "legacy_gbk",
        },
    )

    assert load_settings(settings_path) == {
        "export_format": "txt",
        "encoding": "GBK",
        "last_export_dir": "D:/data",
        "first_row_is_header": True,
        "target_preset": "legacy_gbk",
    }


def test_load_settings_infers_target_preset_for_old_settings_file(tmp_path):
    settings_path = tmp_path / "settings.json"
    settings_path.write_text(
        json.dumps(
            {
                "export_format": "txt",
                "encoding": "GBK",
                "last_export_dir": "D:/data",
                "first_row_is_header": True,
            }
        ),
        encoding="utf-8",
    )

    settings = load_settings(settings_path)

    assert settings["target_preset"] == "legacy_gbk"
    assert settings["export_format"] == "txt"
    assert settings["encoding"] == "GBK"


def test_build_default_filename_uses_format_extension():
    now = datetime(2026, 6, 7, 15, 30, 5)

    assert build_default_filename("csv", now) == "整理数据_20260607_153005.csv"
    assert build_default_filename("txt", now) == "整理数据_20260607_153005.txt"
    assert build_default_filename("tsv", now) == "整理数据_20260607_153005.tsv"


def test_load_settings_returns_defaults_for_corrupt_or_invalid_values(tmp_path):
    corrupt_path = tmp_path / "corrupt.json"
    corrupt_path.write_text("{not json", encoding="utf-8")
    assert load_settings(corrupt_path)["export_format"] == "csv"

    invalid_path = tmp_path / "invalid.json"
    invalid_path.write_text(
        json.dumps(
            {
                "export_format": "xlsx",
                "encoding": "latin-1",
                "last_export_dir": "D:/data",
                "first_row_is_header": True,
                "target_preset": "excel",
            }
        ),
        encoding="utf-8",
    )
    settings = load_settings(invalid_path)
    assert settings["export_format"] == "csv"
    assert settings["encoding"] == "UTF-8 BOM"
    assert settings["first_row_is_header"] is True
    assert settings["last_export_dir"] == "D:/data"


def _reset_tk_root(root: tk.Tk) -> None:
    try:
        root.grab_release()
    except tk.TclError:
        pass
    for child in root.winfo_children():
        child.destroy()
    root.unbind("<Configure>")


@pytest.fixture(scope="module")
def tk_root():
    try:
        root = tk.Tk()
    except tk.TclError as exc:
        pytest.skip(f"Tk unavailable: {exc}")
    root.withdraw()
    try:
        yield root
    finally:
        try:
            root.destroy()
        except tk.TclError:
            pass


@pytest.fixture
def exporter_app(tk_root, tmp_path, monkeypatch):
    monkeypatch.setenv("APPDATA", str(tmp_path))
    _reset_tk_root(tk_root)
    app = CsvPasteExporterApp(tk_root)
    try:
        yield app
    finally:
        app._cancel_scheduled_preview()
        _reset_tk_root(tk_root)


def _load_preview_text(app: CsvPasteExporterApp, text: str) -> None:
    app.text.delete("1.0", tk.END)
    app.text.insert("1.0", text)
    app.text_dirty = True
    app.refresh_preview()


def test_app_preview_empty_state_and_clear(exporter_app):
    app = exporter_app
    assert app.status_var.get() == "粘贴数据后会自动预览"
    app.first_row_is_header_var.set(True)
    app._on_header_toggle()
    assert app.status_var.get() == "粘贴数据后会自动预览"
    app.refresh_preview()
    assert app.rows == []
    assert app.status_var.get() == "粘贴数据后会自动预览"
    assert app.tree.heading("empty")["text"] == "预览"
    assert "尚未检测到表格" in app.tree.item(app.tree.get_children()[0])["values"][0]
    assert app.chart_empty_message == "粘贴至少两列数值数据后显示图表"

    _load_preview_text(app, "A\tB\n1\t2\n")
    app.clear_all()
    assert app.rows == []
    assert app.status_var.get() == "已清空"
    assert app.text_dirty is False


def test_app_column_delete_survives_preview_and_is_exported(
    exporter_app, tmp_path, monkeypatch
):
    app = exporter_app
    _load_preview_text(app, "Strain\tStress\tNote\n0\t0\tstart\n1.2E-3\t345.6\tpeak\n")
    app.column_list.selection_set(2)
    app.delete_selected_columns()
    assert app.rows == [
        ["Strain", "Stress"],
        ["0", "0"],
        ["1.2E-3", "345.6"],
    ]
    assert "已调整列" in app.status_var.get()
    assert "首行作数据" in app.status_var.get()

    app.refresh_preview()
    assert app.rows == [
        ["Strain", "Stress"],
        ["0", "0"],
        ["1.2E-3", "345.6"],
    ]

    app._schedule_preview()
    assert app.preview_after_id is None
    assert app.rows[0] == ["Strain", "Stress"]

    output = tmp_path / "export-sample.csv"
    monkeypatch.setattr(
        "csv_paste_exporter.filedialog.asksaveasfilename",
        lambda **_kwargs: str(output),
    )
    monkeypatch.setattr(app, "_show_export_success", lambda *_args, **_kwargs: None)
    app.export_table()

    data = output.read_bytes()
    assert data.startswith(b"\xef\xbb\xbf")
    with output.open("r", encoding="utf-8-sig", newline="") as handle:
        assert list(csv.reader(handle)) == [
            ["Strain", "Stress"],
            ["0", "0"],
            ["1.2E-3", "345.6"],
        ]
    assert "已导出" in app.status_var.get()


def test_app_chart_axes_follow_deleted_leading_column(exporter_app):
    app = exporter_app
    _load_preview_text(app, "A\tB\tC\n0\t1\t2\n3\t4\t5\n")
    app.first_row_is_header_var.set(True)
    app._on_header_toggle()
    app.chart_x_index = 1
    app.chart_y_index = 2
    app._update_chart_controls()
    app.column_list.selection_set(0)
    app.delete_selected_columns()
    assert app.rows[0] == ["B", "C"]
    assert app.chart_x_index == 0
    assert app.chart_y_index == 1
    assert app.chart_points == [(1.0, 2.0), (4.0, 5.0)]


def test_app_default_chart_axes_stay_distinct_after_deleting_column_zero(
    exporter_app,
):
    app = exporter_app
    _load_preview_text(app, "A\tB\tC\n0\t1\t2\n3\t4\t5\n")
    app.first_row_is_header_var.set(True)
    app._on_header_toggle()
    assert app.chart_x_index == 0
    assert app.chart_y_index == 1
    app.column_list.selection_set(0)
    app.delete_selected_columns()
    assert app.rows[0] == ["B", "C"]
    assert app.chart_x_index != app.chart_y_index
    assert app.chart_x_index == 0
    assert app.chart_y_index == 1
    assert app.chart_points == [(1.0, 2.0), (4.0, 5.0)]


def test_app_restore_original_data_after_column_move(exporter_app):
    app = exporter_app
    _load_preview_text(app, "A\tB\tC\n1\t2\t3\n")
    app.column_list.selection_set(0)
    app.move_selected_column(1)
    assert app.rows == [["B", "A", "C"], ["2", "1", "3"]]
    app.restore_original_data()
    assert app.rows == [["A", "B", "C"], ["1", "2", "3"]]
    assert "已调整列" not in app.status_var.get()


def test_app_header_toggle_updates_preview_and_status(exporter_app):
    app = exporter_app
    _load_preview_text(app, "Time\tForce\n0\t10\n1\t12\n")
    app.first_row_is_header_var.set(True)
    app._on_header_toggle()
    assert app.tree.heading("col_1")["text"] == "Time"
    assert "含表头" in app.status_var.get()
    assert app.chart_info_var.get().startswith("2 个有效点")


def test_app_matlab_preset_survives_matching_export_options(exporter_app):
    app = exporter_app
    app.target_preset_var.set("MATLAB")
    app._on_target_preset_changed()
    assert app.export_format_var.get() == "CSV - Excel/Python 推荐"
    assert app.encoding_var.get() == "UTF-8"
    app._on_export_option_changed()
    assert app.target_preset_var.get() == "MATLAB"


def test_app_paste_clipboard_and_missing_clipboard(exporter_app):
    app = exporter_app
    app.root.clipboard_clear()
    app.root.clipboard_append("A,B\n1,2\n")
    app.paste_clipboard()
    assert app.rows == [["A", "B"], ["1", "2"]]
    assert app.current_analysis.delimiter_name == "逗号"

    app.root.clipboard_clear()
    app.paste_clipboard()
    assert app.rows == [["A", "B"], ["1", "2"]]
    assert app.status_var.get() == "剪贴板里没有可读取的文本"
