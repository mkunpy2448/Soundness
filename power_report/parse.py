"""Read a process-capability workbook by the labels in column A."""

from collections import Counter
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from power_report.model import Report, SeqGroup, Variable
from power_report.numbers import as_float, format_number, is_number

REQUIRED_LABELS = (
    "SEQNo",
    "TIName",
    "Ext. Name",
    "VarName",
    "Upper Spec",
    "Lower Spec",
    "Samples",
    "Mean",
    "Ca",
    "Cpk",
    "Cp",
    "S/N",
)


class TemplateError(Exception):
    """The workbook does not match the expected column-A labels."""

    def __init__(self, message: str, missing: list[str] | None = None):
        super().__init__(message)
        self.missing = missing or []


def load_report(path: Path) -> Report:
    try:
        workbook = load_workbook(path, data_only=False, read_only=False)
    except InvalidFileException as exc:
        raise TemplateError(f"無法讀取 Excel：{path.name}") from exc
    try:
        sheet = _find_sheet(workbook)
        labels = _label_rows(sheet)
        serials = _serials(sheet, labels["S/N"])
        groups = _groups(sheet, labels, serials_rows=_serial_rows(sheet, labels["S/N"]))
    finally:
        workbook.close()
    return Report(source_name=path.name, serials=serials, groups=groups)


def _find_sheet(workbook: Workbook) -> Worksheet:
    for sheet in workbook.worksheets:
        for row in range(1, (sheet.max_row or 0) + 1):
            if _label_text(sheet.cell(row, 1).value) == "SEQNo":
                return sheet
    raise TemplateError("找不到含有 SEQNo 的工作表")


def _label_rows(sheet: Worksheet) -> dict[str, int]:
    found: dict[str, int] = {}
    for row in range(1, (sheet.max_row or 0) + 1):
        label = _label_text(sheet.cell(row, 1).value)
        if label and label not in found:
            found[label] = row
    missing = [label for label in REQUIRED_LABELS if label not in found]
    if missing:
        raise TemplateError("缺少標籤：" + "、".join(missing), missing)
    return found


def _label_text(value: object) -> str | None:
    if not isinstance(value, str):
        return None
    text = value.strip()
    return text or None


def _serial_rows(sheet: Worksheet, sn_row: int) -> list[int]:
    rows: list[int] = []
    for row in range(sn_row + 1, (sheet.max_row or 0) + 1):
        if _cell_text(sheet.cell(row, 1).value) is not None:
            rows.append(row)
    return rows


def _serials(sheet: Worksheet, sn_row: int) -> list[str]:
    serials: list[str] = []
    for row in _serial_rows(sheet, sn_row):
        text = _cell_text(sheet.cell(row, 1).value)
        if text is not None:
            serials.append(text)
    return serials


def _groups(sheet: Worksheet, labels: dict[str, int], serials_rows: list[int]) -> list[SeqGroup]:
    date_row = labels.get("Date")
    groups: list[SeqGroup] = []
    current: SeqGroup | None = None
    current_seq: str | None = None
    for column in range(2, (sheet.max_column or 1) + 1):
        seq_text = _cell_text(sheet.cell(labels["SEQNo"], column).value)
        if seq_text is None:
            current = None
            current_seq = None
            continue
        if current is None or seq_text != current_seq:
            current = SeqGroup(seq=seq_text, ti_name="", ext_name="")
            groups.append(current)
            current_seq = seq_text
        name = _cell_text(sheet.cell(labels["VarName"], column).value) or ""
        ti_name = _cell_text(sheet.cell(labels["TIName"], column).value) or ""
        ext_name = _cell_text(sheet.cell(labels["Ext. Name"], column).value) or ""
        if ti_name and not current.ti_name:
            current.ti_name = ti_name
        if ext_name and not current.ext_name:
            current.ext_name = ext_name
        values: list[float] = []
        for row in serials_rows:
            number = as_float(sheet.cell(row, column).value)
            if number is not None:
                values.append(number)
        current.variables.append(
            Variable(
                column=column,
                name=name,
                upper=as_float(sheet.cell(labels["Upper Spec"], column).value),
                lower=as_float(sheet.cell(labels["Lower Spec"], column).value),
                mean=sheet.cell(labels["Mean"], column).value,
                ca=sheet.cell(labels["Ca"], column).value,
                cpk=sheet.cell(labels["Cpk"], column).value,
                cp=sheet.cell(labels["Cp"], column).value,
                date=_cell_text(sheet.cell(date_row, column).value) if date_row else None,
                values=values,
            )
        )
    for group in groups:
        _assign_uids(group.variables)
    return groups


def _assign_uids(variables: list[Variable]) -> None:
    counts = Counter(variable.name for variable in variables if variable.name)
    seen: Counter[str] = Counter()
    for variable in variables:
        if not variable.name:
            variable.uid = f"column-{variable.column}"
            continue
        if counts[variable.name] == 1:
            variable.uid = variable.name
            continue
        seen[variable.name] += 1
        variable.uid = f"{variable.name}#{seen[variable.name]}"


def _cell_text(value: object) -> str | None:
    if value is None:
        return None
    if isinstance(value, str):
        text = value.strip()
        return text or None
    if isinstance(value, datetime):
        return value.strftime("%Y/%m/%d")
    if isinstance(value, date):
        return value.strftime("%Y/%m/%d")
    if is_number(value):
        return format_number(float(value))
    text = str(value).strip()
    return text or None
