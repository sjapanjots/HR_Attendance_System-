"""Read an .xlsx worksheet and write attendance codes into chosen cells without touching anything else.

The output workbook is produced by splicing new bytes into the selected worksheet's XML only at the
target cells. Every other ZIP member is copied byte-for-byte, and every byte of the worksheet outside the
target cells stays identical. `verify_patch` re-checks this before the output is offered for download.
"""

from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
import posixpath
import re
import zipfile
from typing import Any
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

from .rules import ATTENDANCE_CODES

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


class WorkbookSafetyError(ValueError):
    """The workbook cannot be changed while keeping the promise that only target cells change."""


# ---------------------------------------------------------------------------------------------- helpers


def _tag(name: str) -> str:
    return f"{{{MAIN_NS}}}{name}"


def column_number(letters: str) -> int:
    n = 0
    for char in letters.upper():
        if not "A" <= char <= "Z":
            raise WorkbookSafetyError(f"Invalid column reference: {letters}")
        n = n * 26 + ord(char) - 64
    return n


def column_letters(number: int) -> str:
    if number < 1:
        raise WorkbookSafetyError(f"Invalid column number: {number}")
    result = ""
    while number:
        number, rem = divmod(number - 1, 26)
        result = chr(65 + rem) + result
    return result


_REF = re.compile(r"([A-Z]{1,3})(\d+)")


def split_ref(ref: str) -> tuple[int, int]:
    """'M2' -> (column 13, row 2)."""
    match = _REF.fullmatch(ref.replace("$", ""))
    if not match:
        raise WorkbookSafetyError(f"Invalid cell reference: {ref}")
    return column_number(match.group(1)), int(match.group(2))


def _range_contains(range_ref: str, col: int, row: int) -> bool:
    first, _, last = range_ref.partition(":")
    c1, r1 = split_ref(first)
    c2, r2 = split_ref(last or first)
    return min(c1, c2) <= col <= max(c1, c2) and min(r1, r2) <= row <= max(r1, r2)


def _open_zip(file_bytes: bytes) -> zipfile.ZipFile:
    try:
        return zipfile.ZipFile(BytesIO(file_bytes))
    except zipfile.BadZipFile as exc:
        raise WorkbookSafetyError("This is not a readable .xlsx workbook (it may be password-protected or an old .xls file).") from exc


def sheet_paths(file_bytes: bytes) -> dict[str, str]:
    """Worksheet name -> ZIP path, in workbook order."""
    with _open_zip(file_bytes) as book:
        try:
            workbook = ET.fromstring(book.read("xl/workbook.xml"))
            rels = ET.fromstring(book.read("xl/_rels/workbook.xml.rels"))
        except (KeyError, ET.ParseError) as exc:
            raise WorkbookSafetyError("This file is not a supported Excel .xlsx workbook.") from exc
    if workbook.tag != _tag("workbook"):
        raise WorkbookSafetyError("Unsupported workbook format (only standard, non-Strict Office Open XML .xlsx is supported).")
    targets = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels.findall(f"{{{PKG_REL_NS}}}Relationship")}
    result: dict[str, str] = {}
    for sheet in workbook.findall(f".//{_tag('sheet')}"):
        target = targets.get(sheet.attrib.get(f"{{{REL_NS}}}id", ""))
        if not target or not target.endswith(".xml") or "worksheets/" not in target:
            continue  # chart sheets, dialog sheets and macro sheets are not worksheets
        path = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
        result[sheet.attrib["name"]] = path
    return result


def _shared_strings(book: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(book.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    # Direct <t> or rich-text runs <r><t>; phonetic guides (<rPh><t>) are not part of the displayed text.
    return ["".join(t.text or "" for t in item.findall(f"./{_tag('t')}") + item.findall(f"./{_tag('r')}/{_tag('t')}")) for item in root.findall(_tag("si"))]


def _cell_value(cell: ET.Element, strings: list[str]) -> Any:
    kind = cell.attrib.get("t")
    if kind == "inlineStr":
        inline = cell.find(_tag("is"))
        return "" if inline is None else "".join(t.text or "" for t in inline.iter(_tag("t")))
    value = cell.find(_tag("v"))
    if value is None:
        return None
    raw = value.text or ""
    if kind == "s":
        try:
            return strings[int(raw)]
        except (ValueError, IndexError):
            return None
    if kind == "b":
        return raw == "1"
    if kind in ("str", "e"):
        return raw
    try:
        number = float(raw)
    except ValueError:
        return raw
    return int(number) if number.is_integer() else number


# ---------------------------------------------------------------------------------------------- reading


@dataclass
class SheetData:
    headers: list[str]              # display names, index 0 = column A
    cells: dict[int, dict[int, Any]]  # excel row -> column number -> value (rows after the header row only)
    max_col: int
    header_row: int

    def value(self, row: int, col: int) -> Any:
        return self.cells.get(row, {}).get(col)


def read_sheet(file_bytes: bytes, sheet_name: str, header_row: int = 1) -> SheetData:
    paths = sheet_paths(file_bytes)
    if sheet_name not in paths:
        raise WorkbookSafetyError(f"Worksheet not found: {sheet_name}")
    with _open_zip(file_bytes) as book:
        try:
            root = ET.fromstring(book.read(paths[sheet_name]))
        except (KeyError, ET.ParseError) as exc:
            raise WorkbookSafetyError("Unable to read the selected worksheet.") from exc
        strings = _shared_strings(book)

    by_row: dict[int, dict[int, Any]] = {}
    max_col = 0
    previous_row = 0
    sheet_data = root.find(_tag("sheetData"))
    for row in [] if sheet_data is None else sheet_data.findall(_tag("row")):
        row_idx = int(row.attrib.get("r", previous_row + 1))
        previous_row = row_idx
        values: dict[int, Any] = {}
        previous_col = 0
        for cell in row.findall(_tag("c")):
            col = split_ref(cell.attrib["r"])[0] if "r" in cell.attrib else previous_col + 1
            previous_col = col
            value = _cell_value(cell, strings)
            values[col] = value
            if value is not None:
                max_col = max(max_col, col)
        by_row[row_idx] = values
    if header_row not in by_row or not any(v not in (None, "") for v in by_row[header_row].values()):
        raise WorkbookSafetyError(f"Header row {header_row} is empty or missing.")
    headers: list[str] = []
    seen: set[str] = set()
    for col in range(1, max_col + 1):
        raw = by_row[header_row].get(col)
        name = " ".join(str(raw).split()) if raw not in (None, "") else f"(blank header)"
        name = f"{column_letters(col)}: {name}"
        if name in seen:  # cannot happen with the letter prefix, kept as a guard
            name = f"{name} ({col})"
        seen.add(name)
        headers.append(name)
    cells = {r: v for r, v in by_row.items() if r > header_row}
    return SheetData(headers=headers, cells=cells, max_col=max_col, header_row=header_row)


# ---------------------------------------------------------------------------------------------- patching

_ROW_OPEN = re.compile(rb"<row\b([^>]*?)(/?)>")
_CELL = re.compile(rb"<c\b([^>]*?)(?:/>|>(.*?)</c>)", re.S)
_ATTR_R = re.compile(rb"""\sr\s*=\s*["']([^"']+)["']""")
_ATTR_T = re.compile(rb"""\st\s*=\s*["'][^"']*["']""")
_DIMENSION = re.compile(rb"""<dimension\b[^>]*?\sref\s*=\s*["']([^"']+)["'][^>]*?/>""")


@dataclass
class _RowSpan:
    number: int
    content_start: int  # index just after the opening <row ...>
    content_end: int    # index of the closing </row>
    self_closing: bool


def _check_sheet_xml(xml: bytes) -> tuple[int, int]:
    """Return the sheetData content span, refusing XML layouts the byte-level patch cannot handle."""
    head = xml[:4096]
    if re.search(rb"<\w+:worksheet\b", head):
        raise WorkbookSafetyError("This worksheet uses prefixed XML tags, which this tool cannot edit safely.")
    if f'xmlns="{MAIN_NS}"'.encode() not in head:
        raise WorkbookSafetyError("Unsupported worksheet format (only standard, non-Strict Office Open XML is supported).")
    start = xml.find(b"<sheetData")
    if start < 0:
        raise WorkbookSafetyError("Worksheet has no sheet data section.")
    open_end = xml.find(b">", start)
    if xml[open_end - 1:open_end + 1] == b"/>":
        return open_end + 1, open_end + 1
    end = xml.find(b"</sheetData>", open_end)
    if end < 0:
        raise WorkbookSafetyError("Worksheet sheet data is malformed.")
    return open_end + 1, end


def _scan_rows(xml: bytes, data_start: int, data_end: int) -> dict[int, _RowSpan]:
    rows: dict[int, _RowSpan] = {}
    pos = data_start
    while True:
        match = _ROW_OPEN.search(xml, pos, data_end)
        if not match:
            break
        r = _ATTR_R.search(match.group(1))
        if not r:
            raise WorkbookSafetyError("A worksheet row has no row number; this tool cannot edit it safely.")
        number = int(r.group(1))
        if match.group(2):
            rows[number] = _RowSpan(number, match.end(), match.end(), True)
            pos = match.end()
            continue
        close = xml.find(b"</row>", match.end(), data_end)
        if close < 0:
            raise WorkbookSafetyError("Worksheet row XML is malformed.")
        rows[number] = _RowSpan(number, match.end(), close, False)
        pos = close + len(b"</row>")
    return rows


def _scan_cells(xml: bytes, span: _RowSpan) -> list[tuple[int, int, int, bytes, bytes | None]]:
    """(column, start, end, attributes, inner) for each cell in the row."""
    cells = []
    for match in _CELL.finditer(xml, span.content_start, span.content_end):
        r = _ATTR_R.search(match.group(1))
        if not r:
            raise WorkbookSafetyError(f"Row {span.number} has a cell without a reference; this tool cannot edit it safely.")
        col, row = split_ref(r.group(1).decode())
        if row != span.number:
            raise WorkbookSafetyError(f"Cell {r.group(1).decode()} is stored in row {span.number}.")
        cells.append((col, match.start(), match.end(), match.group(1), match.group(2)))
    return cells


def _shared_string_indexes(book: zipfile.ZipFile) -> dict[str, int]:
    """Index of plain (unformatted) shared strings, so written codes look exactly like existing ones."""
    try:
        root = ET.fromstring(book.read("xl/sharedStrings.xml"))
    except KeyError:
        return {}
    result: dict[str, int] = {}
    for index, si in enumerate(root.findall(_tag("si"))):
        children = list(si)
        if len(children) == 1 and children[0].tag == _tag("t") and children[0].text in ATTENDANCE_CODES:
            result.setdefault(children[0].text, index)
    return result


def _new_cell(ref: str, attrs: bytes | None, text: str, shared: dict[str, int]) -> bytes:
    if attrs is None:
        base = f' r="{ref}"'.encode()
    else:
        if re.search(rb"""\s(vm|cm)\s*=""", attrs):
            raise WorkbookSafetyError(f"Cell {ref} holds rich-value metadata; refusing to overwrite it.")
        base = _ATTR_T.sub(b"", attrs).rstrip()
    if text in shared:
        return b"<c" + base + b' t="s"><v>' + str(shared[text]).encode() + b"</v></c>"
    space = b' xml:space="preserve"' if text != text.strip() else b""
    return b"<c" + base + b' t="inlineStr"><is><t' + space + b">" + escape(text).encode("utf-8") + b"</t></is></c>"


def _protected_ranges(xml: bytes) -> list[tuple[str, str]]:
    ranges = [("merged", m.group(1).decode()) for m in re.finditer(rb"""<mergeCell\b[^>]*?\sref\s*=\s*["']([^"']+)["']""", xml)]
    for m in re.finditer(rb"<f\b([^>]*)>", xml):
        attrs = m.group(1)
        kind = re.search(rb"""\st\s*=\s*["'](array|dataTable)["']""", attrs)
        ref = re.search(rb"""\sref\s*=\s*["']([^"']+)["']""", attrs)
        if kind and ref:
            ranges.append((kind.group(1).decode(), ref.group(1).decode()))
    return ranges


def _patch_sheet(xml: bytes, edits: dict[str, str], shared: dict[str, int], *, append_header: tuple[str, str] | None) -> bytes:
    data_start, data_end = _check_sheet_xml(xml)
    rows = _scan_rows(xml, data_start, data_end)
    protected = _protected_ranges(xml)
    all_edits = dict(edits)
    if append_header:
        all_edits[append_header[0]] = append_header[1]

    by_row: dict[int, dict[int, tuple[str, str]]] = {}
    for ref, value in all_edits.items():
        col, row = split_ref(ref)
        for kind, rng in protected:
            if _range_contains(rng, col, row) and not (kind == "merged" and rng.split(":")[0] == ref):
                raise WorkbookSafetyError(f"Cell {ref} is inside a {kind} range ({rng}); refusing to change it.")
        by_row.setdefault(row, {})[col] = (ref, value)

    splices: list[tuple[int, int, bytes]] = []
    for row_number, targets in by_row.items():
        span = rows.get(row_number)
        if span is None or span.self_closing:
            raise WorkbookSafetyError(f"Row {row_number} has no stored cells; refusing to create new rows.")
        cells = _scan_cells(xml, span)
        existing = {c[0]: c for c in cells}
        ext = xml.find(b"<extLst", span.content_start, span.content_end)
        row_tail = ext if ext >= 0 else span.content_end
        for col, (ref, value) in sorted(targets.items()):
            if col in existing:
                _, start, end, attrs, inner = existing[col]
                if inner is not None and re.search(rb"<f\b", inner):
                    raise WorkbookSafetyError(f"Cell {ref} contains a formula; refusing to overwrite it.")
                splices.append((start, end, _new_cell(ref, attrs, value, shared)))
            else:
                after = [c for c in cells if c[0] > col]
                at = after[0][1] if after else row_tail
                splices.append((at, at, _new_cell(ref, None, value, shared)))

    out = bytearray()
    pos = 0
    for start, end, payload in sorted(splices, key=lambda s: (s[0], s[1])):
        out += xml[pos:start] + payload
        pos = end
    out += xml[pos:]
    return bytes(out)


def _extend_dimension(xml: bytes, col: int, row: int) -> bytes:
    match = _DIMENSION.search(xml)
    if not match:
        return xml
    first, _, last = match.group(1).decode().partition(":")
    c1, r1 = split_ref(first)
    c2, r2 = split_ref(last or first)
    new_ref = f"{column_letters(min(c1, c2))}{min(r1, r2)}:{column_letters(max(c2, c1, col))}{max(r2, r1, row)}"
    start, end = match.span(1)
    return xml[:start] + new_ref.encode() + xml[end:]


def write_attendance(
    file_bytes: bytes,
    *,
    sheet_name: str,
    edits: dict[str, str],
    append_header: tuple[str, str] | None = None,
) -> bytes:
    """Return a new workbook with `edits` (cell ref -> code) written; raise if it cannot be done safely.

    `append_header` = (cell ref, heading) is only for the explicit opt-in to add one classification column.
    """
    for ref, code in edits.items():
        if code not in ATTENDANCE_CODES:
            raise WorkbookSafetyError(f"Unsupported attendance code {code!r} for {ref}.")
    sheet_path = sheet_paths(file_bytes)[sheet_name]
    with _open_zip(file_bytes) as book:
        infos = book.infolist()
        members = {info.filename: book.read(info.filename) for info in infos}
        shared = _shared_string_indexes(book)
    original_sheet = members[sheet_path]
    if append_header:
        if re.search(rb"<mergeCell\b", original_sheet):
            raise WorkbookSafetyError("Cannot safely add a column to a worksheet with merged cells.")
        if re.search(rb"<tablePart\b", original_sheet):
            raise WorkbookSafetyError("Cannot safely add a column to a worksheet that contains an Excel table.")
    patched = _patch_sheet(original_sheet, edits, shared, append_header=append_header)
    if append_header:
        refs = [split_ref(r) for r in [append_header[0], *edits]]
        patched = _extend_dimension(patched, max(c for c, _ in refs), max(r for _, r in refs))
    members[sheet_path] = patched

    output = BytesIO()
    with zipfile.ZipFile(output, "w") as out_zip:
        for info in infos:
            copy = zipfile.ZipInfo(info.filename, date_time=info.date_time)
            copy.compress_type = info.compress_type
            copy.external_attr = info.external_attr
            copy.create_system = info.create_system
            copy.comment = info.comment
            out_zip.writestr(copy, members[info.filename])
    result = output.getvalue()
    verify_patch(file_bytes, result, sheet_name=sheet_name, edits=edits, append_header=append_header)
    return result


# ---------------------------------------------------------------------------------------------- verification


def _strip_cells(xml: bytes, refs: set[str], *, drop_dimension: bool) -> bytes:
    data_start, data_end = _check_sheet_xml(xml)
    rows = _scan_rows(xml, data_start, data_end)
    wanted_rows: dict[int, set[int]] = {}
    for ref in refs:
        col, row = split_ref(ref)
        wanted_rows.setdefault(row, set()).add(col)
    cuts: list[tuple[int, int]] = []
    for row_number, cols in wanted_rows.items():
        span = rows.get(row_number)
        if span is None:
            continue
        cuts += [(start, end) for col, start, end, _, _ in _scan_cells(xml, span) if col in cols]
    out = bytearray()
    pos = 0
    for start, end in sorted(cuts):
        out += xml[pos:start]
        pos = end
    out += xml[pos:]
    result = bytes(out)
    if drop_dimension:
        result = _DIMENSION.sub(b"", result, count=1)
    return result


def verify_patch(
    original: bytes,
    output: bytes,
    *,
    sheet_name: str,
    edits: dict[str, str],
    append_header: tuple[str, str] | None = None,
) -> None:
    """Raise WorkbookSafetyError unless only the target cells differ between the two workbooks."""
    with _open_zip(original) as before, _open_zip(output) as after:
        if [i.filename for i in before.infolist()] != [i.filename for i in after.infolist()]:
            raise WorkbookSafetyError("Verification failed: the workbook package structure changed.")
        sheet_path = sheet_paths(original)[sheet_name]
        for name in before.namelist():
            if name != sheet_path and before.read(name) != after.read(name):
                raise WorkbookSafetyError(f"Verification failed: an unrelated workbook part changed ({name}).")
        refs = set(edits) | ({append_header[0]} if append_header else set())
        sheet_before, sheet_after = before.read(sheet_path), after.read(sheet_path)
        if _strip_cells(sheet_before, refs, drop_dimension=bool(append_header)) != _strip_cells(sheet_after, refs, drop_dimension=bool(append_header)):
            raise WorkbookSafetyError("Verification failed: worksheet content outside the classification cells changed.")
        try:
            root = ET.fromstring(sheet_after)
        except ET.ParseError as exc:
            raise WorkbookSafetyError("Verification failed: the edited worksheet is not valid XML.") from exc
        strings = _shared_strings(after)
    expected = dict(edits)
    if append_header:
        expected[append_header[0]] = append_header[1]
    found: dict[str, Any] = {}
    for cell in root.iter(_tag("c")):
        ref = cell.attrib.get("r")
        if ref in expected:
            if ref in found:
                raise WorkbookSafetyError(f"Verification failed: cell {ref} appears twice.")
            found[ref] = _cell_value(cell, strings)
    for ref, value in expected.items():
        if found.get(ref) != value:
            raise WorkbookSafetyError(f"Verification failed: {ref} does not contain {value!r}.")
