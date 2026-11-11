"""Turn a month-wide sheet into one record per employee-day and classify every record with Polars."""

from __future__ import annotations

from typing import Any

import polars as pl

from .layout import DayBlock
from .parsing import is_blank, parse_clock_minutes, parse_shift_hours, parse_worked_minutes
from .rules import RULES, classify_worked_minutes
from .workbook import SheetData, column_letters

# A worked-hours cell this close to "punch time + whole days" marks a punch pair spanning 24 h or more
# (e.g. in 07:34, out 08:03 the next day, duration 24.29). The punches alone cannot show that.
MULTI_DAY_TOLERANCE_MINUTES = 5

ACTION_UPDATE = "Update"
ACTION_SAME = "Already correct"
ACTION_REVIEW = "Review (left unchanged)"


def _text(value: Any) -> str | None:
    if value is None:
        return None
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value)


def build_records(sheet: SheetData, *, shift_col: int, blocks: list[DayBlock]) -> dict[str, list[Any]]:
    """One row per (worksheet row, day block). Values are read only; nothing is cleaned or changed."""
    cols: dict[str, list[Any]] = {k: [] for k in (
        "excel_row", "day", "day_label", "cell", "shift_raw", "in_raw", "out_raw", "worked_raw", "current",
        "shift_hours", "in_min", "out_min", "worked_cell_min", "has_in", "has_out",
    )}
    for row_number in sorted(sheet.cells):
        values = sheet.cells[row_number]
        if all(is_blank(v) for v in values.values()):
            continue  # an empty worksheet row is not a record
        shift_raw = values.get(shift_col)
        for block in blocks:
            in_raw, out_raw = values.get(block.in_col), values.get(block.out_col)
            worked_raw = values.get(block.worked_col) if block.worked_col else None
            current = values.get(block.status_col)
            cols["excel_row"].append(row_number)
            cols["day"].append(block.index)
            cols["day_label"].append(sheet.headers[block.status_col - 1] if block.status_col <= len(sheet.headers) else f"{column_letters(block.status_col)}: (new column)")
            cols["cell"].append(f"{column_letters(block.status_col)}{row_number}")
            cols["shift_raw"].append(_text(shift_raw))
            cols["in_raw"].append(_text(in_raw))
            cols["out_raw"].append(_text(out_raw))
            cols["worked_raw"].append(_text(worked_raw))
            cols["current"].append(_text(current))
            cols["shift_hours"].append(parse_shift_hours(shift_raw))
            cols["in_min"].append(parse_clock_minutes(in_raw))
            cols["out_min"].append(parse_clock_minutes(out_raw))
            cols["worked_cell_min"].append(worked_raw)  # parsed later, once the format is known
            cols["has_in"].append(not is_blank(in_raw))
            cols["has_out"].append(not is_blank(out_raw))
    return cols


def classify_month(
    sheet: SheetData,
    *,
    shift_col: int,
    blocks: list[DayBlock],
    worked_format: str | None,
) -> pl.DataFrame:
    cols = build_records(sheet, shift_col=shift_col, blocks=blocks)
    cols["worked_cell_min"] = [None if worked_format is None else parse_worked_minutes(v, worked_format) for v in cols["worked_cell_min"]]
    frame = pl.DataFrame(
        cols,
        schema={
            "excel_row": pl.Int64, "day": pl.Int64, "day_label": pl.String, "cell": pl.String,
            "shift_raw": pl.String, "in_raw": pl.String, "out_raw": pl.String, "worked_raw": pl.String, "current": pl.String,
            "shift_hours": pl.Int64, "in_min": pl.Int64, "out_min": pl.Int64, "worked_cell_min": pl.Int64,
            "has_in": pl.Boolean, "has_out": pl.Boolean,
        },
    )

    # Worked time from the punches only. An out-punch earlier than the in-punch is the next morning.
    punch = (pl.col("out_min") - pl.col("in_min") + 1440) % 1440
    frame = frame.with_columns(punch.alias("punch_min"))
    extra_days = ((pl.col("worked_cell_min") - pl.col("punch_min")) / 1440).round(0).cast(pl.Int64)
    frame = frame.with_columns(
        pl.when(
            pl.col("worked_cell_min").is_not_null()
            & pl.col("punch_min").is_not_null()
            & (extra_days >= 1)
            & ((pl.col("worked_cell_min") - pl.col("punch_min") - extra_days * 1440).abs() <= MULTI_DAY_TOLERANCE_MINUTES)
        )
        .then(extra_days)
        .otherwise(0)
        .alias("extra_days")
    )
    frame = frame.with_columns((pl.col("punch_min") + pl.col("extra_days") * 1440).alias("worked_minutes"))
    frame = classify_worked_minutes(frame)

    single = pl.col("has_in") != pl.col("has_out")
    none = ~pl.col("has_in") & ~pl.col("has_out")
    bad_punch = pl.col("in_min").is_null() | pl.col("out_min").is_null()
    bad_shift = pl.col("shift_hours").is_null() | ~pl.col("shift_hours").is_in(list(RULES))
    supported = ", ".join(str(s) for s in RULES)

    proposed = (
        pl.when(single).then(pl.lit("MP"))
        .when(none | bad_punch | bad_shift).then(pl.lit(None, dtype=pl.String))
        .otherwise(pl.col("code"))
    )
    note = (
        pl.when(single & pl.col("has_in")).then(pl.lit("Single punch: out-punch missing"))
        .when(single).then(pl.lit("Single punch: in-punch missing"))
        .when(none).then(pl.lit("No punches"))
        .when(pl.col("in_min").is_null()).then(pl.lit("In-punch is not a valid time: ") + pl.col("in_raw"))
        .when(pl.col("out_min").is_null()).then(pl.lit("Out-punch is not a valid time: ") + pl.col("out_raw"))
        .when(pl.col("shift_hours").is_null()).then(pl.lit("Missing or unreadable shift length: ") + pl.col("shift_raw").fill_null("(blank)"))
        .when(bad_shift).then(pl.lit(f"Unsupported shift length (supported: {supported} hours): ") + pl.col("shift_raw"))
        .when(pl.col("extra_days") > 0).then(pl.lit("Punches span more than 24 h (confirmed by worked-hours cell)"))
        .when(pl.col("worked_cell_min").is_not_null() & ((pl.col("worked_cell_min") - pl.col("worked_minutes")).abs() > MULTI_DAY_TOLERANCE_MINUTES))
        .then(pl.lit("Worked-hours cell differs from punches; punches used"))
        .otherwise(pl.lit(""))
    )
    frame = frame.with_columns(proposed.alias("proposed"), note.alias("note"))
    action = (
        pl.when(pl.col("proposed").is_null()).then(pl.lit(ACTION_REVIEW))
        # Same code apart from letter case or surrounding spaces (e.g. existing "Hf"): leave it as it is.
        .when(pl.col("proposed") == pl.col("current").fill_null("").str.strip_chars().str.to_uppercase()).then(pl.lit(ACTION_SAME))
        .otherwise(pl.lit(ACTION_UPDATE))
    )
    worked_text = (
        pl.when(pl.col("worked_minutes").is_null() | single).then(pl.lit(None, dtype=pl.String))
        .otherwise((pl.col("worked_minutes") // 60).cast(pl.String) + ":" + (pl.col("worked_minutes") % 60).cast(pl.String).str.zfill(2))
    )
    return frame.with_columns(action.alias("action"), worked_text.alias("worked_from_punches")).select(
        "excel_row", "day", "day_label", "cell", "shift_raw", "in_raw", "out_raw", "worked_raw",
        "worked_from_punches", "worked_minutes", "current", "proposed", "action", "note",
    )


def edits_from(result: pl.DataFrame) -> dict[str, str]:
    """Cell ref -> new code, for cells whose code actually changes."""
    updates = result.filter(pl.col("action") == ACTION_UPDATE)
    return dict(zip(updates["cell"].to_list(), updates["proposed"].to_list()))
