"""Turn a workbook and the picker's choices into the slide contents."""

from dataclasses import dataclass, field

from power_report.defaults import blank_to_not_provided
from power_report.model import Report
from power_report.numbers import format_number, is_number


@dataclass
class SlideVariable:
    name: str
    upper: float | None
    lower: float | None
    mean: str
    ca: str
    cpk: str
    cp: str
    values: list[float]


@dataclass
class ContentSlide:
    seq: str
    ti_name: str
    input_voltage: str
    load_current: str
    variables: list[SlideVariable] = field(default_factory=list)

    @property
    def title(self) -> str:
        parts = [f"SEQ {self.seq}"]
        if self.ti_name:
            parts.append(self.ti_name)
        parts.append(self.input_voltage)
        return " ".join(parts)


@dataclass
class SummaryRow:
    seq: str
    ti_name: str
    input_voltage: str


@dataclass
class Deck:
    source_name: str
    dates: list[str]
    unit_count: int
    serials: list[str]
    rows: list[SummaryRow]
    slides: list[ContentSlide]


@dataclass
class GroupChoice:
    key: str
    input_voltage: str
    load_current: str
    checked_uids: set[str]


def assemble(report: Report, choices: list[GroupChoice]) -> Deck:
    by_key = {choice.key: choice for choice in choices}
    slides: list[ContentSlide] = []
    dates: list[str] = []
    seen_dates: set[str] = set()
    for group in report.picker_groups():
        choice = by_key.get(group.key)
        if choice is None:
            continue
        chosen = [
            variable
            for variable in group.plottable_variables
            if variable.uid in choice.checked_uids
        ]
        if not chosen:
            continue
        voltage = blank_to_not_provided(choice.input_voltage)
        load = blank_to_not_provided(choice.load_current)
        for variable in chosen:
            if variable.date and variable.date not in seen_dates:
                seen_dates.add(variable.date)
                dates.append(variable.date)
        slides.append(
            ContentSlide(
                seq=group.seq,
                ti_name=group.ti_name,
                input_voltage=voltage,
                load_current=load,
                variables=[_slide_variable(variable) for variable in chosen],
            )
        )
    rows = [
        SummaryRow(slide.seq, slide.ti_name, slide.input_voltage) for slide in slides
    ]
    return Deck(
        source_name=report.source_name,
        dates=dates,
        unit_count=len(report.serials),
        serials=list(report.serials),
        rows=rows,
        slides=slides,
    )


def _slide_variable(variable) -> SlideVariable:
    return SlideVariable(
        name=variable.name,
        upper=variable.upper,
        lower=variable.lower,
        mean=_display_stat(variable.mean),
        ca=_display_stat(variable.ca),
        cpk=_display_stat(variable.cpk),
        cp=_display_stat(variable.cp),
        values=list(variable.values),
    )


def _display_stat(value: object) -> str:
    if value is None:
        return ""
    if is_number(value):
        return format_number(float(value))
    if isinstance(value, str):
        return value.strip()
    return str(value).strip()
