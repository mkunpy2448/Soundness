"""Numeric helpers shared by parsing, defaults, and slide text."""

import math
import statistics
from decimal import Decimal, ROUND_HALF_UP


def is_number(value: object) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def as_float(value: object) -> float | None:
    """Return a finite float, or None when the cell is blank or not numeric."""
    if is_number(value):
        number = float(value)
        if math.isnan(number) or math.isinf(number):
            return None
        return number
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if text == "" or text == "------":
            return None
        try:
            number = float(text)
        except ValueError:
            return None
        if math.isnan(number) or math.isinf(number):
            return None
        return number
    return None


def is_missing_spec(value: object) -> bool:
    return as_float(value) is None


def format_number(value: float) -> str:
    """Print a stored Excel number without scientific notation."""
    number = float(value)
    if math.isnan(number) or math.isinf(number):
        return ""
    if number.is_integer() and abs(number) < 1e15:
        return str(int(number))
    text = f"{number:.6f}".rstrip("0").rstrip(".")
    if text in {"", "-0"}:
        return "0"
    return text


def round_half_up_2(value: float) -> str:
    """Round half away from zero and always show two decimal places."""
    quantized = Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return f"{quantized:.2f}"


def quartiles(values: list[float]) -> tuple[float, float, float]:
    """Return Q1, median, Q3 using the inclusive quartile definition."""
    ordered = sorted(values)
    if len(ordered) == 1:
        only = ordered[0]
        return only, only, only
    q1, median, q3 = statistics.quantiles(ordered, n=4, method="inclusive")
    return q1, median, q3
