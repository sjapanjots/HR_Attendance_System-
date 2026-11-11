"""Monthly attendance marking: Streamlit UI.

Run with:  streamlit run app.py
"""

from __future__ import annotations

import hashlib
from pathlib import Path

import polars as pl
import streamlit as st

from attendance.layout import DayBlock, check_blocks, expand_blocks
from attendance.parsing import WORKED_FORMATS
from attendance.pipeline import ACTION_REVIEW, ACTION_SAME, ACTION_UPDATE, classify_month, edits_from
from attendance.rules import GRACE_MINUTES, rules_review_table
from attendance.workbook import SheetData, WorkbookSafetyError, column_letters, read_sheet, sheet_paths, write_attendance

NOT_AVAILABLE = "— not in this sheet —"


@st.cache_data(show_spinner=False, max_entries=4)
def _sheet_names(file_bytes: bytes) -> list[str]:
    return list(sheet_paths(file_bytes))


@st.cache_data(show_spinner="Reading worksheet…", max_entries=4)
def _read(file_bytes: bytes, sheet_name: str, header_row: int) -> SheetData:
    return read_sheet(file_bytes, sheet_name, header_row)


def _col(headers: list[str], choice: str | None) -> int | None:
    return None if choice in (None, NOT_AVAILABLE) else headers.index(choice) + 1


def _preview(sheet: SheetData, rows: int = 15) -> pl.DataFrame:
    numbers = sorted(sheet.cells)[:rows]
    data = {"Row": numbers}
    for col, header in enumerate(sheet.headers, start=1):
        data[header] = [None if sheet.value(r, col) is None else str(sheet.value(r, col)) for r in numbers]
    return pl.DataFrame(data, schema={"Row": pl.Int64, **{h: pl.String for h in sheet.headers}})


def _reset_on_change(signature: tuple) -> None:
    if st.session_state.get("signature") != signature:
        st.session_state["signature"] = signature
        for key in ("result", "output", "output_name"):
            st.session_state.pop(key, None)


def main() -> None:
    st.set_page_config(page_title="Monthly Attendance Marking", layout="wide")
    st.title("Monthly attendance marking")
    st.caption(
        "Marks every employee-day of the month from the punch-in and punch-out times. Only the mapped "
        "attendance classification cells are written, into a separate copy of the workbook."
    )

    uploaded = st.file_uploader("1. Upload the attendance workbook (.xlsx)", type=["xlsx"])
    if not uploaded:
        st.info("Upload a workbook to begin. The uploaded file is never modified.")
        return
    file_bytes = uploaded.getvalue()
    try:
        names = _sheet_names(file_bytes)
    except WorkbookSafetyError as exc:
        st.error(str(exc))
        return
    if not names:
        st.error("This workbook has no worksheets that can be read.")
        return

    left, right = st.columns([3, 1])
    sheet_name = left.selectbox("2. Worksheet", names, index=None, placeholder="Choose the worksheet to mark")
    header_row = int(right.number_input("Header row", min_value=1, value=1, step=1))
    if sheet_name is None:
        return
    try:
        sheet = _read(file_bytes, sheet_name, header_row)
    except WorkbookSafetyError as exc:
        st.error(str(exc))
        return

    with st.expander("Read-only preview (first 15 rows)", expanded=False):
        st.dataframe(_preview(sheet), hide_index=True, width="stretch")

    # ------------------------------------------------------------------ mapping
    st.subheader("3. Map the columns")
    st.write("Map the columns for the **first day** of the sheet. The same layout is repeated for every following day.")
    headers = sheet.headers
    shift_choice = st.selectbox("Shift length column", headers, index=None, placeholder="e.g. SHIFT HOURS")
    c1, c2, c3, c4 = st.columns(4)
    in_choice = c1.selectbox("Day 1 in-punch", headers, index=None)
    out_choice = c2.selectbox("Day 1 out-punch", headers, index=None)
    worked_choice = c3.selectbox(
        "Day 1 worked hours",
        [NOT_AVAILABLE, *headers],
        index=None,
        help="Used only to recognise punches that span 24 hours or more (e.g. in 07:34, out 08:03 the next day). "
        "The classification itself is calculated from the punch times.",
    )
    append = st.checkbox(
        "This sheet has no attendance classification column: add one (allowed for single-day sheets only)",
        value=False,
    )
    if append:
        status_choice = None
        new_heading = c4.text_input("New classification column heading", value="Attendance Classification")
    else:
        status_choice = c4.selectbox("Day 1 attendance classification", headers, index=None)
        new_heading = ""

    worked_format = None
    if worked_choice not in (None, NOT_AVAILABLE):
        fmt_label = st.radio(
            "How are worked hours written in that column?",
            list(WORKED_FORMATS.values()),
            index=None,
            horizontal=True,
        )
        worked_format = next((k for k, v in WORKED_FORMATS.items() if v == fmt_label), None)

    d1, d2 = st.columns(2)
    days = d1.number_input("Number of days to mark", min_value=1, max_value=31, value=None, step=1, placeholder="e.g. 30")
    second_choice = None
    if days and days > 1:
        second_choice = d2.selectbox("Day 2 in-punch (sets the width of one day's columns)", headers, index=None)

    missing = [
        label
        for label, value in (
            ("shift length", shift_choice),
            ("day 1 in-punch", in_choice),
            ("day 1 out-punch", out_choice),
            ("day 1 worked hours (or choose 'not in this sheet')", worked_choice),
            ("day 1 attendance classification", status_choice if not append else "ok"),
            ("worked-hours format", worked_format if worked_choice not in (None, NOT_AVAILABLE) else "ok"),
            ("number of days", days),
            ("day 2 in-punch", second_choice if days and days > 1 else "ok"),
        )
        if value is None
    ]
    if append and not new_heading.strip():
        missing.append("new classification column heading")
    if missing:
        st.warning("Still to choose: " + "; ".join(missing) + ".")
        return
    if append and days != 1:
        st.error("Adding a classification column is only supported when marking a single day. Map the existing day-wise classification columns instead.")
        return

    shift_col = _col(headers, shift_choice)
    status_col = sheet.max_col + 1 if append else _col(headers, status_choice)
    first = DayBlock(1, _col(headers, in_choice), _col(headers, out_choice), _col(headers, worked_choice), status_col)
    try:
        blocks = expand_blocks(first=first, second_in_col=_col(headers, second_choice), days=int(days), max_col=sheet.max_col)
    except ValueError as exc:
        st.error(str(exc))
        return
    problems = check_blocks(blocks, headers)
    block_cols = {c for b in blocks for c in (b.in_col, b.out_col, b.worked_col, b.status_col) if c}
    if shift_col in block_cols:
        problems.append("The shift length column is also mapped as a punch, worked-hours or classification column.")
    if problems:
        st.error("The column mapping does not look consistent, so nothing will be processed:\n\n- " + "\n- ".join(problems[:20]))
        return

    def label(col: int | None) -> str:
        if col is None:
            return "—"
        return headers[col - 1] if col <= len(headers) else f"{column_letters(col)}: {new_heading} (new)"

    mapping_table = pl.DataFrame(
        {
            "Day": [b.index for b in blocks],
            "In-punch": [label(b.in_col) for b in blocks],
            "Out-punch": [label(b.out_col) for b in blocks],
            "Worked hours": [label(b.worked_col) for b in blocks],
            "Classification (written)": [label(b.status_col) for b in blocks],
        }
    )
    st.dataframe(mapping_table, hide_index=True, width="stretch", height=min(38 + 35 * len(blocks), 420))
    mapping_ok = st.checkbox(f"I have checked the column mapping for all {len(blocks)} day(s) above")

    # ------------------------------------------------------------------ rules
    st.subheader("4. Review the attendance thresholds")
    st.dataframe(rules_review_table(), hide_index=True, width="stretch")
    st.caption(
        f"Times are worked time (h:mm) from punch-in to punch-out; an out-punch earlier than the in-punch counts as the next day. "
        f"P starts {GRACE_MINUTES} minutes before the full shift length (grace period). "
        "Exactly one punch → MP. No punches, an unreadable punch time, or a missing/unsupported shift length → the cell is left unchanged and listed for review. "
        "Cells that already hold the proposed code (ignoring letter case, e.g. 'Hf') are left as they are."
    )
    rules_ok = st.checkbox("I have reviewed these thresholds")
    if not (mapping_ok and rules_ok):
        st.info("Confirm the mapping and the thresholds to preview the changes.")
        return

    signature = (
        hashlib.sha256(file_bytes).hexdigest(), sheet_name, header_row, shift_col, tuple(blocks), worked_format, append, new_heading,
    )
    _reset_on_change(signature)

    # ------------------------------------------------------------------ preview
    st.subheader("5. Preview the proposed changes")
    if st.button("Calculate attendance for the whole sheet", type="primary"):
        try:
            st.session_state["result"] = classify_month(sheet, shift_col=shift_col, blocks=blocks, worked_format=worked_format)
        except Exception as exc:  # malformed data should produce a clear message, not a traceback
            st.error(f"Could not classify the sheet: {exc}")
            return
    result: pl.DataFrame | None = st.session_state.get("result")
    if result is None:
        return

    counts = dict(result.group_by("action").len().iter_rows())
    m1, m2, m3 = st.columns(3)
    m1.metric("Cells to update", counts.get(ACTION_UPDATE, 0))
    m2.metric("Already correct", counts.get(ACTION_SAME, 0))
    m3.metric("Left unchanged for review", counts.get(ACTION_REVIEW, 0))

    view = st.radio("Show", ["Cells that will change", "Left unchanged for review", "All employee-days"], horizontal=True)
    shown = {
        "Cells that will change": result.filter(pl.col("action") == ACTION_UPDATE),
        "Left unchanged for review": result.filter(pl.col("action") == ACTION_REVIEW),
        "All employee-days": result,
    }[view]
    st.dataframe(
        shown.select(
            pl.col("excel_row").alias("Excel row"),
            pl.col("day_label").alias("Day column"),
            pl.col("cell").alias("Cell"),
            pl.col("shift_raw").alias("Shift"),
            pl.col("in_raw").alias("In"),
            pl.col("out_raw").alias("Out"),
            pl.col("worked_from_punches").alias("Worked (from punches)"),
            pl.col("worked_raw").alias("Worked-hours cell"),
            pl.col("current").alias("Old code"),
            pl.col("proposed").alias("New code"),
            pl.col("action").alias("Action"),
            pl.col("note").alias("Note"),
        ),
        hide_index=True,
        width="stretch",
    )

    edits = edits_from(result)
    if not edits and not append:
        st.success("Every classification cell already matches the rules. No output workbook is needed.")
        return

    # ------------------------------------------------------------------ write
    st.subheader("6. Create the output workbook")
    approve = st.checkbox(f"I approve writing {len(edits)} classification cell(s) to a new copy of the workbook")
    if approve and st.button("Create output workbook"):
        append_header = (f"{column_letters(status_col)}{header_row}", new_heading) if append else None
        try:
            with st.spinner("Writing and verifying…"):
                st.session_state["output"] = write_attendance(file_bytes, sheet_name=sheet_name, edits=edits, append_header=append_header)
            st.session_state["output_name"] = f"{Path(uploaded.name).stem}_attendance_marked.xlsx"
        except WorkbookSafetyError as exc:
            st.session_state.pop("output", None)
            st.error(f"No output was created: {exc}")
            return
    if approve and st.session_state.get("output"):
        st.success("Verification passed: only the approved classification cells differ; every other cell, sheet and workbook part is unchanged.")
        st.download_button(
            "Download marked workbook",
            data=st.session_state["output"],
            file_name=st.session_state["output_name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )


if __name__ == "__main__":
    main()
