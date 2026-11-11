"""Attendance classification rules: the single place to review or change thresholds.

All thresholds are in minutes of worked time, calculated from the punch-in and punch-out times.
Each range includes its lower bound and excludes its upper bound, except that P includes its upper bound,
so the ranges never overlap and leave no gaps.

The P lower bound for every shift is 15 minutes short of the shift length (the 15-minute grace period),
e.g. 11 h 45 min worked on a 12-hour shift is P, not SP.
"""

from __future__ import annotations

import polars as pl

GRACE_MINUTES = 15


def _hm(hours: int, minutes: int = 0) -> int:
    return hours * 60 + minutes


# shift hours -> thresholds in minutes.
#   A   : worked <  hf_from
#   HF  : hf_from <= worked < sp_from
#   SP  : sp_from <= worked < p_from
#   P   : p_from  <= worked <= p_through
#   PP  : worked  >  p_through
RULES: dict[int, dict[str, int]] = {
    9: {"hf_from": _hm(5), "sp_from": _hm(8), "p_from": _hm(9) - GRACE_MINUTES, "p_through": _hm(17)},
    10: {"hf_from": _hm(5), "sp_from": _hm(9), "p_from": _hm(10) - GRACE_MINUTES, "p_through": _hm(18)},
    11: {"hf_from": _hm(5), "sp_from": _hm(10), "p_from": _hm(11) - GRACE_MINUTES, "p_through": _hm(20)},
    12: {"hf_from": _hm(6), "sp_from": _hm(10), "p_from": _hm(12) - GRACE_MINUTES, "p_through": _hm(22)},
}

ATTENDANCE_CODES = ("A", "HF", "SP", "P", "PP", "MP")


def format_minutes(total: int | None) -> str:
    if total is None:
        return ""
    return f"{total // 60}:{total % 60:02d}"


def rules_frame() -> pl.DataFrame:
    """The thresholds as a frame; used both for the review table and for the classification join."""
    return pl.DataFrame(
        {
            "shift_hours": list(RULES),
            "hf_from": [r["hf_from"] for r in RULES.values()],
            "sp_from": [r["sp_from"] for r in RULES.values()],
            "p_from": [r["p_from"] for r in RULES.values()],
            "p_through": [r["p_through"] for r in RULES.values()],
        },
        schema={"shift_hours": pl.Int64, "hf_from": pl.Int64, "sp_from": pl.Int64, "p_from": pl.Int64, "p_through": pl.Int64},
    )


def rules_review_table() -> pl.DataFrame:
    """Human-readable thresholds (h:mm worked) for the user to review before processing."""
    rows = []
    for shift, r in RULES.items():
        rows.append(
            {
                "Shift": f"{shift} hours",
                "A": f"< {format_minutes(r['hf_from'])}",
                "HF": f"{format_minutes(r['hf_from'])} to < {format_minutes(r['sp_from'])}",
                "SP": f"{format_minutes(r['sp_from'])} to < {format_minutes(r['p_from'])}",
                "P": f"{format_minutes(r['p_from'])} to {format_minutes(r['p_through'])} (inclusive)",
                "PP": f"> {format_minutes(r['p_through'])}",
            }
        )
    return pl.DataFrame(rows)


def classify_worked_minutes(frame: pl.DataFrame) -> pl.DataFrame:
    """Add a `code` column from `shift_hours` and `worked_minutes` (both Int64, possibly null).

    Rows whose shift is not in RULES or whose worked time is null get a null code.
    """
    joined = frame.join(rules_frame(), on="shift_hours", how="left", maintain_order="left")
    w = pl.col("worked_minutes")
    code = (
        pl.when(w.is_null() | pl.col("hf_from").is_null())
        .then(pl.lit(None, dtype=pl.String))
        .when(w < pl.col("hf_from"))
        .then(pl.lit("A"))
        .when(w < pl.col("sp_from"))
        .then(pl.lit("HF"))
        .when(w < pl.col("p_from"))
        .then(pl.lit("SP"))
        .when(w <= pl.col("p_through"))
        .then(pl.lit("P"))
        .otherwise(pl.lit("PP"))
    )
    return joined.with_columns(code.alias("code")).drop("hf_from", "sp_from", "p_from", "p_through")
