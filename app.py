"""Monthly attendance marking: Streamlit UI.

Run with:  streamlit run app.py
"""

from __future__ import annotations

import hashlib
from html import escape
from pathlib import Path

import polars as pl
import streamlit as st

from attendance.layout import DayBlock, check_blocks, expand_blocks
from attendance.parsing import WORKED_FORMATS
from attendance.pipeline import ACTION_REVIEW, ACTION_SAME, ACTION_UPDATE, classify_month, edits_from
from attendance.rules import GRACE_MINUTES, rules_review_table
from attendance.workbook import SheetData, WorkbookSafetyError, column_letters, read_sheet, sheet_paths, write_attendance

NOT_AVAILABLE = "— not in this sheet —"
STEPS = ["Upload", "Worksheet", "Map columns", "Thresholds", "Preview", "Export"]
CODE_MEANINGS = {"P": "Present", "PP": "Present plus", "SP": "Short present", "HF": "Half day", "A": "Absent", "MP": "Missed punch"}
STYLE = (Path(__file__).parent / "assets" / "style.css").read_text(encoding="utf-8")


# ---------------------------------------------------------------------------------------------- data


# Short-lived cache: uploaded HR data should not stay in server memory longer than needed.
@st.cache_data(show_spinner=False, max_entries=4, ttl=900)
def _sheet_names(file_bytes: bytes) -> list[str]:
    return list(sheet_paths(file_bytes))


@st.cache_data(show_spinner="Reading worksheet…", max_entries=4, ttl=900)
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


# ---------------------------------------------------------------------------------------------- presentation


def _html(markup: str) -> None:
    st.markdown(markup, unsafe_allow_html=True)


def _stepper(reached: int) -> str:
    items = []
    for i, name in enumerate(STEPS):
        state = "done" if i < reached else "active" if i == reached else ""
        mark = "✓" if i < reached else str(i + 1)
        items.append(f'<li class="{state}"><span class="dot">{mark}</span>{name}</li>')
    return f'<ul class="stepper">{"".join(items)}</ul>'


def _card_head(number: int, title: str, description: str, *, done: bool = False) -> None:
    _html(
        f'<div class="card-head{" done" if done else ""}"><span class="num">{"✓" if done else number}</span>'
        f"<div><h3>{escape(title)}</h3><p>{escape(description)}</p></div></div>"
    )


def _note(markup: str) -> None:
    _html(f'<div class="note">{markup}</div>')


def _chips(counts: dict[str, int] | None = None) -> str:
    chips = []
    for code, meaning in CODE_MEANINGS.items():
        if counts is not None and not counts.get(code):
            continue
        text = f"{counts[code]:,}" if counts is not None else meaning
        chips.append(f'<span class="chip"><b class="code-{code}">{code}</b>{text}</span>')
    return f'<div class="chips">{"".join(chips)}</div>'


def _metrics(items: list[tuple[str, int, str]]) -> None:
    tiles = "".join(
        f'<div class="metric"><div class="label"><i style="background:{color}"></i>{escape(label)}</div><div class="value">{value:,}</div></div>'
        for label, value, color in items
    )
    _html(f'<div class="metrics">{tiles}</div>')


# ---------------------------------------------------------------------------------------------- app


def main() -> None:
    st.set_page_config(page_title="Attendance Marking", page_icon="🗓️", layout="wide")
    _html(f"<style>{STYLE}</style>")
    _html(
        '<div class="hero"><div class="eyebrow">HR Attendance</div><h1>Monthly attendance marking</h1>'
        "<p>Mark every employee-day of the month from punch-in and punch-out times. Only the attendance cells you map are "
        "written, into a separate copy of your workbook.</p></div>"
    )
    stepper = st.empty()
    progress = {"reached": 0}
    try:
        _run(progress)
    finally:
        with stepper:
            _html(_stepper(progress["reached"]))


def _run(progress: dict[str, int]) -> None:
    # ---------------------------------------------------------------- 1. upload
    with st.container(border=True, key="card_upload"):
        _card_head(1, "Upload workbook", "An .xlsx attendance workbook. The uploaded file is never modified.")
        uploaded = st.file_uploader("Attendance workbook", type=["xlsx"], label_visibility="collapsed")
    if not uploaded:
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
    progress["reached"] = 1

    # ---------------------------------------------------------------- 2. worksheet
    with st.container(border=True, key="card_sheet"):
        _card_head(2, "Choose the worksheet", "Pick the sheet that holds the month's punches and the header row.")
        left, right = st.columns([3, 1])
        sheet_name = left.selectbox("Worksheet", names, index=None, placeholder="Choose a worksheet")
        header_row = int(right.number_input("Header row", min_value=1, value=1, step=1))
        sheet = None
        if sheet_name is not None:
            try:
                sheet = _read(file_bytes, sheet_name, header_row)
            except WorkbookSafetyError as exc:
                st.error(str(exc))
            if sheet is not None:
                with st.expander(f"Preview · {len(sheet.cells):,} rows · {sheet.max_col} columns"):
                    st.dataframe(_preview(sheet), hide_index=True, width="stretch", height=300)
    if sheet is None:
        return
    progress["reached"] = 2

    # ---------------------------------------------------------------- 3. mapping
    headers = sheet.headers
    with st.container(border=True, key="card_map"):
        _card_head(3, "Map the columns", "Map the first day's columns. The same layout is repeated for every day of the month.")
        shift_choice = st.selectbox("Shift length", headers, index=None, placeholder="e.g. G: SHIFT HOURS")

        _html('<div class="subhead">Day 1</div>')
        c1, c2, c3, c4 = st.columns(4)
        in_choice = c1.selectbox("In-punch", headers, index=None, placeholder="Select")
        out_choice = c2.selectbox("Out-punch", headers, index=None, placeholder="Select")
        worked_choice = c3.selectbox(
            "Worked hours",
            [NOT_AVAILABLE, *headers],
            index=None,
            placeholder="Select",
            help="Only used to recognise punches that span 24 hours or more (e.g. in 07:34, out 08:03 the next day). "
            "The code itself is calculated from the punch times.",
        )
        append = st.session_state.get("append", False)
        if append:
            status_choice = None
            new_heading = c4.text_input("New classification column", value="Attendance Classification")
        else:
            status_choice = c4.selectbox("Attendance code (written)", headers, index=None, placeholder="Select")
            new_heading = ""

        worked_format = None
        if worked_choice not in (None, NOT_AVAILABLE):
            fmt_label = st.radio("Worked-hours format", list(WORKED_FORMATS.values()), index=None, horizontal=True)
            worked_format = next((k for k, v in WORKED_FORMATS.items() if v == fmt_label), None)

        _html('<div class="subhead">Month</div>')
        d1, d2 = st.columns(2)
        days = d1.number_input("Number of days", min_value=1, max_value=31, value=None, step=1, placeholder="e.g. 30")
        second_choice = None
        if days and days > 1:
            second_choice = d2.selectbox("Day 2 in-punch", headers, index=None, placeholder="Sets the width of one day's columns")
        st.checkbox("No attendance code column in this sheet — add one (single-day sheets only)", key="append")

        missing = [
            label
            for label, value in (
                ("shift length", shift_choice),
                ("in-punch", in_choice),
                ("out-punch", out_choice),
                ("worked hours", worked_choice),
                ("attendance code column", status_choice if not append else "ok"),
                ("worked-hours format", worked_format if worked_choice not in (None, NOT_AVAILABLE) else "ok"),
                ("number of days", days),
                ("day 2 in-punch", second_choice if days and days > 1 else "ok"),
            )
            if value is None
        ]
        if append and not new_heading.strip():
            missing.append("new column heading")
        if missing:
            _note("<b>Still to choose:</b> " + escape(", ".join(missing)))
            return
        if append and days != 1:
            st.error("Adding a code column is only supported for a single day. Map the existing day-wise code columns instead.")
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
            problems.append("The shift length column is also mapped as a punch, worked-hours or code column.")
        if problems:
            st.error("The mapping does not look consistent, so nothing will be processed:\n\n- " + "\n- ".join(problems[:20]))
            return

        def label(col: int | None) -> str:
            if col is None:
                return "—"
            return headers[col - 1] if col <= len(headers) else f"{column_letters(col)}: {new_heading} (new)"

        with st.expander(f"Day-by-day mapping · {len(blocks)} day(s)", expanded=False):
            st.dataframe(
                pl.DataFrame(
                    {
                        "Day": [b.index for b in blocks],
                        "In-punch": [label(b.in_col) for b in blocks],
                        "Out-punch": [label(b.out_col) for b in blocks],
                        "Worked hours": [label(b.worked_col) for b in blocks],
                        "Attendance code (written)": [label(b.status_col) for b in blocks],
                    }
                ),
                hide_index=True,
                width="stretch",
                height=min(38 + 35 * len(blocks), 420),
            )
        mapping_ok = st.checkbox(f"I have checked the mapping for all {len(blocks)} day(s)")
    if not mapping_ok:
        return
    progress["reached"] = 3

    # ---------------------------------------------------------------- 4. thresholds
    with st.container(border=True, key="card_rules"):
        _card_head(4, "Review the thresholds", f"Worked time from punch-in to punch-out (h:mm). P starts {GRACE_MINUTES} minutes before the full shift.")
        st.dataframe(rules_review_table(), hide_index=True, width="stretch")
        _html(_chips())
        _note(
            "Out-punch earlier than in-punch counts as the next day. <b>One punch → MP.</b> No punches, an unreadable time or an "
            "unsupported shift → cell left unchanged and listed for review. Cells that already hold the code (e.g. <b>Hf</b>) are left as they are."
        )
        rules_ok = st.checkbox("I have reviewed these thresholds")
    if not rules_ok:
        return
    progress["reached"] = 4

    signature = (hashlib.sha256(file_bytes).hexdigest(), sheet_name, header_row, shift_col, tuple(blocks), worked_format, append, new_heading)
    _reset_on_change(signature)

    # ---------------------------------------------------------------- 5. preview
    with st.container(border=True, key="card_preview"):
        _card_head(5, "Preview changes", "Every employee-day is classified. Nothing is written yet.")
        if st.button("Calculate attendance", type="primary"):
            try:
                st.session_state["result"] = classify_month(sheet, shift_col=shift_col, blocks=blocks, worked_format=worked_format)
            except Exception as exc:  # malformed data should produce a clear message, not a traceback
                st.error(f"Could not classify the sheet: {exc}")
                return
        result: pl.DataFrame | None = st.session_state.get("result")
        if result is None:
            return

        counts = dict(result.group_by("action").len().iter_rows())
        _metrics(
            [
                ("Cells to update", counts.get(ACTION_UPDATE, 0), "#18181b"),
                ("Already correct", counts.get(ACTION_SAME, 0), "#16a34a"),
                ("Left unchanged for review", counts.get(ACTION_REVIEW, 0), "#f59e0b"),
            ]
        )
        new_codes = dict(result.filter(pl.col("action") == ACTION_UPDATE).group_by("proposed").len().iter_rows())
        if new_codes:
            _html('<div class="subhead">New codes</div>' + _chips(new_codes))

        views = {
            "Will change": result.filter(pl.col("action") == ACTION_UPDATE),
            "For review": result.filter(pl.col("action") == ACTION_REVIEW),
            "All": result,
        }
        view = st.segmented_control("Show", list(views), default="Will change", label_visibility="collapsed") or "Will change"
        st.dataframe(
            views[view].select(
                pl.col("excel_row").alias("Row"),
                pl.col("day_label").alias("Day"),
                pl.col("cell").alias("Cell"),
                pl.col("shift_raw").alias("Shift"),
                pl.col("in_raw").alias("In"),
                pl.col("out_raw").alias("Out"),
                pl.col("worked_from_punches").alias("Worked"),
                pl.col("current").alias("Old"),
                pl.col("proposed").alias("New"),
                pl.col("note").alias("Note"),
            ),
            hide_index=True,
            width="stretch",
            height=380,
            column_config={
                "Row": st.column_config.NumberColumn(width="small", format="%d"),
                "Cell": st.column_config.TextColumn(width="small"),
                "Old": st.column_config.TextColumn(width="small"),
                "New": st.column_config.TextColumn(width="small"),
                "Note": st.column_config.TextColumn(width="large"),
            },
        )

    edits = edits_from(result)
    if not edits and not append:
        _note("<b>Nothing to change.</b> Every attendance code already matches the rules.")
        return
    progress["reached"] = 5

    # ---------------------------------------------------------------- 6. export
    with st.container(border=True, key="card_export"):
        _card_head(6, "Create the output workbook", "A new copy with only the approved cells changed, verified before download.")
        approve = st.checkbox(f"I approve writing {len(edits):,} attendance cell(s) to a new copy of the workbook")
        if approve and st.button("Create workbook", type="primary"):
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
            progress["reached"] = 6
            _html(
                '<div class="verified">✓ Verified: only the approved attendance cells differ. Every other cell, sheet and '
                "workbook part is unchanged.</div>"
            )
            st.download_button(
                "Download workbook",
                data=st.session_state["output"],
                file_name=st.session_state["output_name"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                icon=":material/download:",
            )


if __name__ == "__main__":
    main()
