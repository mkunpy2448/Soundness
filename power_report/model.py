"""Workbook data kept after the column-A labels have been located."""

from dataclasses import dataclass, field
import json


@dataclass
class Variable:
    column: int
    name: str
    upper: float | None
    lower: float | None
    mean: object
    ca: object
    cpk: object
    cp: object
    date: str | None
    values: list[float]
    uid: str = ""

    @property
    def plottable(self) -> bool:
        return bool(self.values)


@dataclass
class SeqGroup:
    seq: str
    ti_name: str
    ext_name: str
    variables: list[Variable] = field(default_factory=list)

    @property
    def key(self) -> str:
        return json.dumps(
            [self.seq, self.ti_name, self.ext_name],
            ensure_ascii=False,
            separators=(",", ":"),
        )

    @property
    def plottable_variables(self) -> list[Variable]:
        return [variable for variable in self.variables if variable.plottable]


@dataclass
class Report:
    source_name: str
    serials: list[str]
    groups: list[SeqGroup]

    def picker_groups(self) -> list[SeqGroup]:
        return [group for group in self.groups if group.plottable_variables]
