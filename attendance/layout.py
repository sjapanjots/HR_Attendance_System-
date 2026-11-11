"""Day-block column mapping for month-wide attendance sheets.

The user maps the columns of the first day explicitly (in-punch, out-punch, worked hours, classification)
and picks the in-punch column of the second day. The same column offsets are then repeated for the number
of days the user enters. Every generated mapping is shown for review and checked against the headers;
nothing is inferred from header names alone.
"""

from __future__ import annotations

from dataclasses import dataclass
import re

from .workbook import column_letters


@dataclass(frozen=True)
class DayBlock:
    index: int        # 1-based position in the sheet (not the calendar date)
    in_col: int
    out_col: int
    worked_col: int | None
    status_col: int


def _number_tokens(header: str) -> set[str]:
    # Headers are displayed as "M: 26 Fr"; ignore the column-letter prefix.
    text = header.split(": ", 1)[-1]
    return {str(int(n)) for n in re.findall(r"\d+", text)}


def _shape(header: str) -> str:
    return re.sub(r"\d+", "#", header.split(": ", 1)[-1]).strip().lower()


def expand_blocks(
    *,
    first: DayBlock,
    second_in_col: int | None,
    days: int,
    max_col: int,
) -> list[DayBlock]:
    """Repeat the first day's column offsets `days` times using the step between the two in-punch columns."""
    if days < 1:
        raise ValueError("Enter at least one day.")
    if days == 1:
        return [first]
    if second_in_col is None:
        raise ValueError("Select the second day's in-punch column so the day-block width is known.")
    step = second_in_col - first.in_col
    used = [first.in_col, first.out_col, first.status_col] + ([first.worked_col] if first.worked_col else [])
    width = max(used) - min(used) + 1
    if step < width:
        raise ValueError(
            f"The second day's in-punch column is only {step} column(s) after the first, but the first day's mapped "
            f"columns span {width} columns. Day blocks would overlap."
        )
    blocks = []
    for i in range(days):
        shift = i * step
        block = DayBlock(
            index=i + 1,
            in_col=first.in_col + shift,
            out_col=first.out_col + shift,
            worked_col=None if first.worked_col is None else first.worked_col + shift,
            status_col=first.status_col + shift,
        )
        if max(block.in_col, block.out_col, block.status_col, block.worked_col or 0) > max_col:
            raise ValueError(f"Day {i + 1} would use columns beyond the last used column ({column_letters(max_col)}). Reduce the number of days.")
        blocks.append(block)
    return blocks


def check_blocks(blocks: list[DayBlock], headers: list[str]) -> list[str]:
    """Problems found by comparing each generated day block's headers with the first day's headers."""
    problems: list[str] = []
    first = blocks[0]

    def h(col: int | None) -> str:
        return "" if col is None else headers[col - 1]

    for block in blocks[1:]:
        for role, a, b in (
            ("in-punch", first.in_col, block.in_col),
            ("out-punch", first.out_col, block.out_col),
            ("worked-hours", first.worked_col, block.worked_col),
        ):
            if a is not None and _shape(h(a)) != _shape(h(b)):
                problems.append(f"Day {block.index}: {role} header '{h(b)}' does not match day 1 ('{h(a)}').")
        first_tokens = [_number_tokens(h(c)) for c in (first.in_col, first.out_col, first.worked_col, first.status_col) if c]
        if first_tokens and set.intersection(*first_tokens):
            tokens = [_number_tokens(h(c)) for c in (block.in_col, block.out_col, block.worked_col, block.status_col) if c]
            if not set.intersection(*tokens):
                problems.append(f"Day {block.index}: the headers of this block do not share a day number ({', '.join(h(c) for c in (block.in_col, block.out_col, block.worked_col, block.status_col) if c)}).")
    all_cols = [c for b in blocks for c in (b.in_col, b.out_col, b.worked_col, b.status_col) if c]
    if len(all_cols) != len(set(all_cols)):
        problems.append("Two mapped fields use the same column.")
    return problems
