import csv
from datetime import datetime

from csv_paste_exporter import (
    analyze_table_text,
    build_chart_data,
    build_default_filename,
    calculate_axis_range,
    delete_columns,
    get_chart_column_labels,
    get_default_chart_column_indices,
    get_preview_headings,
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


def test_settings_round_trip_with_defaults(tmp_path):
    settings_path = tmp_path / "settings.json"

    assert load_settings(settings_path) == {
        "export_format": "csv",
        "encoding": "UTF-8 BOM",
        "last_export_dir": "",
        "first_row_is_header": False,
    }

    save_settings(
        settings_path,
        {
            "export_format": "txt",
            "encoding": "GBK",
            "last_export_dir": "D:/data",
            "first_row_is_header": True,
        },
    )

    assert load_settings(settings_path) == {
        "export_format": "txt",
        "encoding": "GBK",
        "last_export_dir": "D:/data",
        "first_row_is_header": True,
    }


def test_build_default_filename_uses_format_extension():
    now = datetime(2026, 6, 7, 15, 30, 5)

    assert build_default_filename("csv", now) == "整理数据_20260607_153005.csv"
    assert build_default_filename("txt", now) == "整理数据_20260607_153005.txt"
    assert build_default_filename("tsv", now) == "整理数据_20260607_153005.tsv"
