"""Read-only interpretation of cell values. Nothing here changes workbook data."""

from __future__ import annotations

import re
from typing import Any

WORKED_FORMATS = {
    "hmm_dot": "Hours.minutes (12.30 = 12 h 30 min)",
    "decimal": "Decimal hours (12.5 = 12 h 30 min)",
    "excel_fraction": "Excel time / day fraction (0.5 = 12 h)",
}

_CLOCK = re.compile(r"(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([ap]\.?m\.?)?", re.IGNORECASE)
_SHIFT = re.compile(r"(\d{1,2})(?:\.0+)?\s*(?:h|hr|hrs|hour|hours)?\.?", re.IGNORECASE)


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return not value.strip()
    return False


def parse_clock_minutes(value: Any) -> int | None:
    """Minutes after midnight for a punch time such as '07:38', '7:38 PM' or an Excel time fraction.

    Returns None when the value cannot be read as a time of day.
    """
    if is_blank(value) or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if value < 0:
            return None
        fraction = float(value) % 1  # an Excel date-time serial keeps the time in its fraction
        return round(fraction * 1440) % 1440
    match = _CLOCK.fullmatch(str(value).strip())
    if not match:
        return None
    hours, minutes, seconds = int(match.group(1)), int(match.group(2)), int(match.group(3) or 0)
    meridiem = (match.group(4) or "").lower().replace(".", "")
    if minutes >= 60 or seconds >= 60:
        return None
    if meridiem:
        if not 1 <= hours <= 12:
            return None
        hours = hours % 12 + (12 if meridiem == "pm" else 0)
    elif hours >= 24:
        return None
    return (hours * 60 + minutes + round(seconds / 60)) % 1440


def parse_worked_minutes(value: Any, fmt: str) -> int | None:
    """Worked time in minutes from a worked-hours/duration cell in the user-selected format."""
    if is_blank(value) or isinstance(value, bool):
        return None
    if isinstance(value, str):
        text = value.strip()
        clock = re.fullmatch(r"(\d+):(\d{2})(?::(\d{2}))?", text)
        if clock:
            minutes = int(clock.group(2))
            if minutes >= 60:
                return None
            return int(clock.group(1)) * 60 + minutes + round(int(clock.group(3) or 0) / 60)
        try:
            number = float(text)
        except ValueError:
            return None
    elif isinstance(value, (int, float)):
        number = float(value)
    else:
        return None
    if number < 0:
        return None
    if fmt == "hmm_dot":
        whole, _, frac = f"{number:.2f}".partition(".")
        minutes = int(frac)
        if minutes >= 60:
            return None
        return int(whole) * 60 + minutes
    if fmt == "decimal":
        return round(number * 60)
    if fmt == "excel_fraction":
        return round(number * 1440)
    raise ValueError(f"Unknown worked-hours format: {fmt}")


def parse_shift_hours(value: Any) -> int | None:
    """Shift length in whole hours from values such as 12, '12', '12 HRS' or '09 HRS'."""
    if is_blank(value) or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return int(value) if float(value).is_integer() else None
    match = _SHIFT.fullmatch(str(value).strip())
    return int(match.group(1)) if match else None
