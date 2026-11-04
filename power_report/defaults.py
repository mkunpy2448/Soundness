"""Default text for the two editable fields on the picker page."""

from power_report.model import SeqGroup
from power_report.numbers import round_half_up_2

NOT_PROVIDED = "Not provided"


def blank_to_not_provided(value: object) -> str:
    if value is None:
        return NOT_PROVIDED
    text = str(value).strip()
    return text if text else NOT_PROVIDED


def default_voltage(group: SeqGroup) -> str:
    return group.ext_name if group.ext_name else NOT_PROVIDED


def default_load(group: SeqGroup) -> str:
    """Upper Spec of every VarName containing Iout, left to right, two decimals."""
    parts: list[str] = []
    for variable in group.variables:
        if "iout" not in variable.name.casefold():
            continue
        if variable.upper is None:
            continue
        parts.append(round_half_up_2(variable.upper))
    if not parts:
        return NOT_PROVIDED
    return " / ".join(parts)
