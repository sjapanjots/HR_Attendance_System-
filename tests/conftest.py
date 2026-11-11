"""Build small .xlsx workbooks by hand so tests control every byte of the package."""

from __future__ import annotations

from io import BytesIO
import sys
from pathlib import Path
import zipfile
from xml.sax.saxutils import escape

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from attendance.workbook import column_letters  # noqa: E402

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _cell(ref: str, value, strings: list[str], style: int | None) -> str:
    s = f' s="{style}"' if style is not None else ""
    if value is None:
        return f'<c r="{ref}"{s}/>'
    if isinstance(value, str) and value.startswith("="):
        return f'<c r="{ref}"{s}><f>{escape(value[1:])}</f><v>0</v></c>'
    if isinstance(value, (int, float)):
        return f'<c r="{ref}"{s}><v>{value}</v></c>'
    if value not in strings:
        strings.append(value)
    return f'<c r="{ref}"{s} t="s"><v>{strings.index(value)}</v></c>'


def make_workbook(rows: list[list], *, extra_sheet_xml: str = "", style: int | None = 1, other_sheet: bool = True) -> bytes:
    strings: list[str] = []
    row_xml = []
    for r, values in enumerate(rows, start=1):
        cells = "".join(_cell(f"{column_letters(c)}{r}", v, strings, style) for c, v in enumerate(values, start=1) if v != "<skip>")
        row_xml.append(f'<row r="{r}" spans="1:{len(values)}" x14ac:dyDescent="0.3">{cells}</row>')
    width = max(len(r) for r in rows)
    sheet = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<worksheet xmlns="{MAIN}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" '
        'xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'
        f'<dimension ref="A1:{column_letters(width)}{len(rows)}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews>'
        f'<sheetData>{"".join(row_xml)}</sheetData>{extra_sheet_xml}'
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>'
    )
    sst = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="{MAIN}" count="{len(strings)}" uniqueCount="{len(strings)}">'
        + "".join(f"<si><t>{escape(s)}</t></si>" for s in strings)
        + "</sst>"
    )
    other = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="{MAIN}"><sheetData><row r="1"><c r="A1"><v>42</v></c></row></sheetData></worksheet>'
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        "</Types>"
    )
    workbook = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="{MAIN}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="Working" sheetId="1" r:id="rId1"/><sheet name="Other" sheetId="2" r:id="rId2"/></sheets></workbook>'
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
        '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        "</Relationships>"
    )
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        "</Relationships>"
    )
    styles = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="{MAIN}">'
        '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
        "</styleSheet>"
    )
    out = BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", rels)
        z.writestr("xl/worksheets/sheet1.xml", sheet)
        z.writestr("xl/worksheets/sheet2.xml", other)
        z.writestr("xl/sharedStrings.xml", sst)
        z.writestr("xl/styles.xml", styles)
    return out.getvalue()


HEADER = ["S.NO.", "Emp Code", "SHIFT \nHOURS",
          "In Time 26", "Out Time 26", "OT Hr 26", "Duration Hr 26", "26 Fr",
          "In Time 27", "Out Time 27", "OT Hr 27", "Duration Hr 27", "27 Sa",
          "REMARKS"]


@pytest.fixture
def month_workbook() -> bytes:
    rows = [
        HEADER,
        # day shift: 12:13 worked -> P (already P); day 27 single punch -> MP (currently A)
        [1, 3003139, "12 HRS", "07:38", "19:51", "", 12.13, "P", "07:40", "", "", "", "A", None],
        # night shift 19:48 -> 08:20 = 12:32 -> P (currently blank); 24 h double shift -> PP (currently P)
        [2, 3000494, "12 HRS", "19:48", "08:20", "", 12.32, None, "07:34", "08:03", 12.43, 24.29, "P", None],
        # 9 h shift: 8:44 -> SP, 8:45 -> P (grace); existing 'Hf' style codes stay when they already match
        [3, 3000777, "09 HRS", "09:00", "17:44", "", 8.44, "P", "09:00", "17:45", "", 8.45, "SP", None],
        # no punches -> left unchanged (keeps WO); unsupported shift -> left unchanged
        [4, 3000888, "8 HRS", "", "", "", "", "WO", "08:00", "16:00", "", 8.0, "P", None],
        # 12 h shift: 5:59 -> A, 6:00 -> HF ('Hf' already there counts as correct)
        [5, 3000999, "12 HRS", "08:00", "13:59", "", 5.59, "HF", "08:00", "14:00", "", 6.0, "Hf", None],
    ]
    return make_workbook(rows)
