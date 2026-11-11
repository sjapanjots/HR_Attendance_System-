from __future__ import annotations

from io import BytesIO
import os
import zipfile

import polars as pl
import pytest

from attendance.layout import DayBlock, check_blocks, expand_blocks
from attendance.parsing import parse_clock_minutes, parse_shift_hours, parse_worked_minutes
from attendance.pipeline import ACTION_REVIEW, ACTION_SAME, ACTION_UPDATE, classify_month, edits_from
from attendance.rules import RULES, classify_worked_minutes
from attendance.workbook import WorkbookSafetyError, read_sheet, verify_patch, write_attendance

from conftest import make_workbook


# ------------------------------------------------------------------ parsing


@pytest.mark.parametrize(
    "value, expected",
    [("07:38", 458), ("7:38", 458), ("19:51", 1191), ("07:38:40", 459), ("7:38 PM", 1178), ("12:05 am", 5), (0.5, 720), (46205.25, 360), ("", None), (None, None), ("25:00", None), ("abc", None)],
)
def test_parse_clock(value, expected):
    assert parse_clock_minutes(value) == expected


def test_parse_worked_formats():
    assert parse_worked_minutes(12.13, "hmm_dot") == 733
    assert parse_worked_minutes(12.5, "hmm_dot") == 770  # 12.50 = 12 h 50 min
    assert parse_worked_minutes("24.29", "hmm_dot") == 1469
    assert parse_worked_minutes(12.75, "hmm_dot") is None  # 75 minutes is not valid
    assert parse_worked_minutes(12.5, "decimal") == 750
    assert parse_worked_minutes(0.5, "excel_fraction") == 720
    assert parse_worked_minutes("11:45", "decimal") == 705


@pytest.mark.parametrize("value, expected", [("12 HRS", 12), ("09 HRS", 9), (12, 12), ("10 hours", 10), ("12.0", 12), ("", None), ("twelve", None), (9.5, None)])
def test_parse_shift(value, expected):
    assert parse_shift_hours(value) == expected


# ------------------------------------------------------------------ rules


def _code(shift: int, minutes: int) -> str:
    frame = pl.DataFrame({"shift_hours": [shift], "worked_minutes": [minutes]}, schema={"shift_hours": pl.Int64, "worked_minutes": pl.Int64})
    return classify_worked_minutes(frame)["code"][0]


@pytest.mark.parametrize(
    "shift, a_below, sp_from, p_from, p_through",
    [(9, 300, 480, 525, 1020), (10, 300, 540, 585, 1080), (11, 300, 600, 645, 1200), (12, 360, 600, 705, 1320)],
)
def test_rule_boundaries(shift, a_below, sp_from, p_from, p_through):
    assert _code(shift, a_below - 1) == "A"
    assert _code(shift, a_below) == "HF"
    assert _code(shift, sp_from - 1) == "HF"
    assert _code(shift, sp_from) == "SP"
    assert _code(shift, p_from - 1) == "SP"
    assert _code(shift, p_from) == "P"  # 15-minute grace: e.g. 11:45 on a 12-hour shift
    assert _code(shift, p_through) == "P"
    assert _code(shift, p_through + 1) == "PP"


def test_grace_is_fifteen_minutes():
    assert all(shift * 60 - rule["p_from"] == 15 for shift, rule in RULES.items())


# ------------------------------------------------------------------ month pipeline


def _blocks(sheet):
    first = DayBlock(1, 4, 5, 7, 8)
    blocks = expand_blocks(first=first, second_in_col=9, days=2, max_col=sheet.max_col)
    assert check_blocks(blocks, sheet.headers) == []
    return blocks


def test_classify_month(month_workbook):
    sheet = read_sheet(month_workbook, "Working", 1)
    result = classify_month(sheet, shift_col=3, blocks=_blocks(sheet), worked_format="hmm_dot")
    got = {cell: (proposed, action) for cell, proposed, action in result.select("cell", "proposed", "action").iter_rows()}
    assert got["H2"] == ("P", ACTION_SAME)
    assert got["M2"] == ("MP", ACTION_UPDATE)          # single punch
    assert got["H3"] == ("P", ACTION_UPDATE)           # night shift crosses midnight
    assert got["M3"] == ("PP", ACTION_UPDATE)          # 24:29 recognised through the worked-hours cell
    assert got["H4"] == ("SP", ACTION_UPDATE)          # 8:44 on 9 h
    assert got["M4"] == ("P", ACTION_UPDATE)           # 8:45 on 9 h (grace)
    assert got["H5"] == (None, ACTION_REVIEW)          # no punches: keep WO
    assert got["M5"] == (None, ACTION_REVIEW)          # 8 h shift is not in the rules
    assert got["H6"] == ("A", ACTION_UPDATE)           # 5:59 on 12 h
    assert got["M6"] == ("HF", ACTION_SAME)            # existing 'Hf' is left as it is
    assert edits_from(result) == {"M2": "MP", "H3": "P", "M3": "PP", "H4": "SP", "M4": "P", "H6": "A"}


def test_without_worked_column_double_shift_is_short(month_workbook):
    sheet = read_sheet(month_workbook, "Working", 1)
    blocks = expand_blocks(first=DayBlock(1, 4, 5, None, 8), second_in_col=9, days=2, max_col=sheet.max_col)
    result = classify_month(sheet, shift_col=3, blocks=blocks, worked_format=None)
    row = result.filter(pl.col("cell") == "M3").row(0, named=True)
    assert row["worked_minutes"] == 29 and row["proposed"] == "A"


def test_block_overlap_and_header_mismatch(month_workbook):
    sheet = read_sheet(month_workbook, "Working", 1)
    with pytest.raises(ValueError):
        expand_blocks(first=DayBlock(1, 4, 5, 7, 8), second_in_col=6, days=2, max_col=sheet.max_col)
    with pytest.raises(ValueError):
        expand_blocks(first=DayBlock(1, 4, 5, 7, 8), second_in_col=9, days=3, max_col=sheet.max_col)
    shifted = expand_blocks(first=DayBlock(1, 4, 5, 7, 8), second_in_col=10, days=2, max_col=sheet.max_col)
    assert check_blocks(shifted, sheet.headers)


# ------------------------------------------------------------------ workbook writing


def _members(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(BytesIO(data)) as z:
        return {n: z.read(n) for n in z.namelist()}


def test_write_changes_only_target_cells(month_workbook):
    sheet = read_sheet(month_workbook, "Working", 1)
    edits = edits_from(classify_month(sheet, shift_col=3, blocks=_blocks(sheet), worked_format="hmm_dot"))
    out = write_attendance(month_workbook, sheet_name="Working", edits=edits)
    before, after = _members(month_workbook), _members(out)
    assert list(before) == list(after)
    for name in before:
        if name != "xl/worksheets/sheet1.xml":
            assert before[name] == after[name], name
    new_sheet = read_sheet(out, "Working", 1)
    old_cells = sheet.cells
    for row, values in new_sheet.cells.items():
        for col, value in values.items():
            ref = f"{'ABCDEFGHIJKLMN'[col - 1]}{row}"
            assert value == (edits[ref] if ref in edits else old_cells[row].get(col)), ref
    # style index kept; an existing shared string is reused for 'P', an inline string for 'MP'
    assert b'<c r="H3" s="1" t="s">' in after["xl/worksheets/sheet1.xml"]
    assert b'<c r="M2" s="1" t="inlineStr"><is><t>MP</t></is></c>' in after["xl/worksheets/sheet1.xml"]


def test_written_workbook_opens_in_openpyxl(month_workbook):
    openpyxl = pytest.importorskip("openpyxl")
    out = write_attendance(month_workbook, sheet_name="Working", edits={"M2": "MP", "H3": "P"})
    ws = openpyxl.load_workbook(BytesIO(out))["Working"]
    assert ws["M2"].value == "MP" and ws["H3"].value == "P" and ws["H2"].value == "P"


def test_refuses_formula_targets():
    data = make_workbook([["In", "Out", "Status"], ["08:00", "20:00", "=1+1"]])
    with pytest.raises(WorkbookSafetyError, match="formula"):
        write_attendance(data, sheet_name="Working", edits={"C2": "P"})


def test_refuses_cells_inside_merged_range():
    data = make_workbook([["In", "Out", "Status", "X"], ["08:00", "20:00", "A", "A"]], extra_sheet_xml='<mergeCells count="1"><mergeCell ref="C2:D2"/></mergeCells>')
    with pytest.raises(WorkbookSafetyError, match="merged"):
        write_attendance(data, sheet_name="Working", edits={"D2": "P"})
    write_attendance(data, sheet_name="Working", edits={"C2": "P"})  # top-left cell is fine


def test_unknown_code_is_rejected(month_workbook):
    with pytest.raises(WorkbookSafetyError):
        write_attendance(month_workbook, sheet_name="Working", edits={"H2": "XX"})


def test_missing_cell_is_inserted_in_column_order():
    data = make_workbook([["In", "Out", "Status", "Note"], ["08:00", "20:00", "<skip>", "n"]])
    out = write_attendance(data, sheet_name="Working", edits={"C2": "PP"})
    xml = _members(out)["xl/worksheets/sheet1.xml"]
    assert xml.index(b'r="B2"') < xml.index(b'r="C2"') < xml.index(b'r="D2"')
    assert b"<is><t>PP</t></is>" in xml  # 'PP' is not in shared strings, so an inline string is used


def test_append_column_opt_in():
    data = make_workbook([["SHIFT", "In", "Out"], ["12", "08:00", "20:00"]])
    out = write_attendance(data, sheet_name="Working", edits={"D2": "P"}, append_header=("D1", "Attendance Classification"))
    sheet = read_sheet(out, "Working", 1)
    assert sheet.headers[-1] == "D: Attendance Classification" and sheet.value(2, 4) == "P"
    assert b'<dimension ref="A1:D2"/>' in _members(out)["xl/worksheets/sheet1.xml"]


def test_verify_detects_unrelated_change(month_workbook):
    out = write_attendance(month_workbook, sheet_name="Working", edits={"M2": "MP"})
    members = _members(out)
    members["xl/worksheets/sheet1.xml"] = members["xl/worksheets/sheet1.xml"].replace(b"<v>3003139</v>", b"<v>1</v>")
    tampered = BytesIO()
    with zipfile.ZipFile(tampered, "w") as z:
        for name, payload in members.items():
            z.writestr(name, payload)
    with pytest.raises(WorkbookSafetyError, match="outside the classification cells"):
        verify_patch(month_workbook, tampered.getvalue(), sheet_name="Working", edits={"M2": "MP"})


@pytest.mark.skipif(not os.environ.get("ATTENDANCE_SAMPLE"), reason="set ATTENDANCE_SAMPLE=/path/to/workbook.xlsx to run")
def test_real_month_workbook():
    data = open(os.environ["ATTENDANCE_SAMPLE"], "rb").read()
    sheet = read_sheet(data, "Working", 1)
    first = DayBlock(1, 9, 10, 12, 13)  # I, J, L, M
    blocks = expand_blocks(first=first, second_in_col=14, days=30, max_col=sheet.max_col)
    assert check_blocks(blocks, sheet.headers) == []
    result = classify_month(sheet, shift_col=7, blocks=blocks, worked_format="hmm_dot")
    edits = edits_from(result)
    out = write_attendance(data, sheet_name="Working", edits=edits)
    assert read_sheet(out, "Working", 1).cells.keys() == sheet.cells.keys()
