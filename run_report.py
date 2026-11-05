"""Double-click this file, or run_report.bat on Windows, to build a PowerPoint report.

This file is self-contained. Copy it together with run_report.bat; the power_report folder is not required.
"""

import argparse
import ctypes
import json
import math
import statistics
import subprocess
import sys
import traceback
import webbrowser
import zipfile
from collections import Counter
from collections.abc import Callable
from ctypes import wintypes
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from io import BytesIO
from pathlib import Path

ROOT = Path(__file__).resolve().parent


def ensure_packages() -> None:
    """Install Excel and PowerPoint libraries before they are imported."""
    try:
        import openpyxl  # noqa: F401
        import pptx  # noqa: F401
    except ImportError:
        requirements = ROOT / "requirements.txt"
        command = [sys.executable, "-m", "pip", "install"]
        if requirements.is_file():
            command.extend(["-r", str(requirements)])
        else:
            command.extend(["openpyxl>=3.1.5", "python-pptx>=1.0.2"])
        print("正在安裝需要的套件，請稍候...")
        subprocess.check_call(command)


if __name__ == "__main__":
    try:
        ensure_packages()
    except Exception as exc:
        print(f"發生錯誤：{exc}")
        try:
            input("\n按 Enter 關閉...")
        except EOFError:
            pass
        raise SystemExit(1)

from lxml import etree
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

try:
    import tkinter as tk
    from tkinter import filedialog
except ImportError:
    tk = None
    filedialog = None

PAGE = r'''<!DOCTYPE html>
<html lang="zh-Hant">
<head>
  <meta charset="utf-8">
  <title>勾選要放入 PPT 的項目</title>
  <style>
    body { margin: 0; color: #1a1a1a; background: #f4f4f4; font-family: Calibri, "Segoe UI", sans-serif; }
    header { position: sticky; top: 0; z-index: 2; background: #fff; border-bottom: 1px solid #ddd; padding: 12px 20px; display: flex; gap: 12px; align-items: center; }
    header h1 { margin: 0; font-size: 18px; font-weight: 600; }
    header .file { color: #555; font-size: 13px; flex: 1; }
    input[type="search"], input[type="text"] { border: 1px solid #bbb; padding: 6px 8px; font: inherit; }
    input[type="search"] { width: 220px; }
    button { font: inherit; cursor: pointer; }
    button.primary { background: #1f4e79; color: #fff; border: 0; padding: 8px 14px; }
    button.primary:disabled { opacity: 0.6; cursor: default; }
    button.ghost { background: #fff; border: 1px solid #bbb; padding: 4px 8px; }
    .hint { margin: 0; padding: 10px 20px 0; color: #333; font-size: 14px; }
    .report-fields { display: flex; flex-wrap: wrap; gap: 24px; padding: 12px 20px 0; background: #fff; }
    main { padding: 12px 20px 48px; }
    section { background: #fff; border: 1px solid #ddd; margin-bottom: 12px; padding: 12px 14px; }
    section h2 { margin: 0 0 8px; font-size: 16px; }
    .toolbar { display: flex; gap: 8px; margin-bottom: 8px; }
    .fields { display: flex; flex-wrap: wrap; gap: 16px; margin-bottom: 10px; }
    .field { display: flex; flex-direction: column; gap: 4px; font-size: 13px; }
    .field-row { display: flex; gap: 6px; align-items: center; }
    .field input { width: 320px; }
    .vars { display: flex; flex-direction: column; gap: 4px; }
    .vars label { display: flex; gap: 8px; align-items: center; }
    .gone { display: none; }
    #status { min-height: 1.2em; font-size: 13px; white-space: pre-wrap; }
    #status.error { color: #9b1c1c; }
  </style>
</head>
<body>
  <header>
    <h1>勾選要放入 PPT 的項目</h1>
    <div class="file" id="file-name"></div>
    <input id="filter" type="search" placeholder="篩選變數名稱">
    <button class="primary" id="generate" type="button">產生 PPT</button>
  </header>
  <div class="report-fields">
    <label class="field">輸入電壓範圍
      <input id="voltage-range" type="text">
    </label>
    <label class="field">滿載電流
      <input id="full-load" type="text">
    </label>
  </div>
  <p class="hint">勾選要出現在 PPT 的變數。輸入電壓與負載電流可以修改，產生時以格子裡的文字為準。</p>
  <main id="groups"></main>
  <script id="report-data" type="application/json">__REPORT_JSON__</script>
  <script>
    const report = JSON.parse(document.getElementById("report-data").textContent);
    document.getElementById("file-name").textContent = report.fileName;
    document.getElementById("voltage-range").value = report.voltageRange || "";
    document.getElementById("full-load").value = report.fullLoadCurrent || "";

    const root = document.getElementById("groups");
    for (const group of report.groups) {
      const section = document.createElement("section");
      section.dataset.key = group.key;
      const title = document.createElement("h2");
      title.textContent = ["SEQ " + group.seq, group.tiName, group.extName].filter(Boolean).join("  ");
      section.appendChild(title);

      const toolbar = document.createElement("div");
      toolbar.className = "toolbar";
      toolbar.appendChild(actionButton("全選", function () { setAll(section, true); }));
      toolbar.appendChild(actionButton("全不選", function () { setAll(section, false); }));
      section.appendChild(toolbar);

      const fields = document.createElement("div");
      fields.className = "fields";
      fields.appendChild(textField("輸入電壓", "voltage", group.inputVoltage, group.inputDefault));
      fields.appendChild(textField("負載電流", "load", group.loadCurrent, group.loadDefault));
      section.appendChild(fields);

      const vars = document.createElement("div");
      vars.className = "vars";
      for (const variable of group.variables) {
        const label = document.createElement("label");
        label.dataset.name = variable.name;
        const box = document.createElement("input");
        box.type = "checkbox";
        box.checked = variable.checked;
        box.dataset.uid = variable.uid;
        label.appendChild(box);
        label.appendChild(document.createTextNode(variable.name));
        vars.appendChild(label);
      }
      section.appendChild(vars);
      root.appendChild(section);
    }

    document.getElementById("filter").addEventListener("input", function (event) {
      const query = event.target.value.trim().toLowerCase();
      for (const section of root.querySelectorAll("section")) {
        let visible = 0;
        for (const label of section.querySelectorAll(".vars label")) {
          const matched = !query || label.dataset.name.toLowerCase().includes(query);
          label.classList.toggle("gone", !matched);
          if (matched) visible += 1;
        }
        section.classList.toggle("gone", visible === 0);
      }
    });

    const status = document.createElement("div");
    status.id = "status";
    root.before(status);

    document.getElementById("generate").addEventListener("click", async function () {
      const button = document.getElementById("generate");
      button.disabled = true;
      status.className = "";
      status.textContent = "產生中…";
      try {
        const response = await fetch("/api/generate", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            groups: collect(),
            voltage_range: document.getElementById("voltage-range").value,
            full_load_current: document.getElementById("full-load").value,
          }),
        });
        const payload = await response.json();
        if (!response.ok || !payload.ok) {
          throw new Error(payload.error || "產生失敗");
        }
        status.textContent = "已產生 PPT：" + payload.path + "\n已儲存勾選：" + payload.selectionPath;
      } catch (error) {
        status.className = "error";
        status.textContent = error.message;
      } finally {
        button.disabled = false;
      }
    });

    function actionButton(text, onClick) {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "ghost";
      button.textContent = text;
      button.addEventListener("click", onClick);
      return button;
    }

    function textField(labelText, role, value, fallback) {
      const wrapper = document.createElement("label");
      wrapper.className = "field";
      wrapper.appendChild(document.createTextNode(labelText));
      const row = document.createElement("div");
      row.className = "field-row";
      const input = document.createElement("input");
      input.type = "text";
      input.dataset.role = role;
      input.dataset.default = fallback;
      input.value = value;
      const restore = document.createElement("button");
      restore.type = "button";
      restore.className = "ghost";
      restore.textContent = "還原";
      restore.addEventListener("click", function () { input.value = input.dataset.default; });
      row.appendChild(input);
      row.appendChild(restore);
      wrapper.appendChild(row);
      return wrapper;
    }

    function setAll(section, checked) {
      for (const box of section.querySelectorAll('.vars input[type="checkbox"]')) {
        box.checked = checked;
      }
    }

    function collect() {
      return Array.from(root.querySelectorAll("section")).map(function (section) {
        const voltage = section.querySelector('input[data-role="voltage"]').value;
        const load = section.querySelector('input[data-role="load"]').value;
        const checked = Array.from(section.querySelectorAll('.vars input[type="checkbox"]'))
          .filter(function (box) { return box.checked; })
          .map(function (box) { return box.dataset.uid; });
        return { key: section.dataset.key, input_voltage: voltage, load_current: load, checked: checked };
      });
    }
  </script>
</body>
</html>
'''

"""Numeric helpers shared by parsing, defaults, and slide text."""



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


def round_half_up(value: float, places: int = 3) -> str:
    """Round half away from zero and drop trailing zeros."""
    quant = Decimal("1").scaleb(-places)
    quantized = Decimal(str(value)).quantize(quant, rounding=ROUND_HALF_UP)
    text = format(quantized, "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    if text in {"", "-0"}:
        return "0"
    return text


def format_load_number(value: float) -> str:
    """Two decimals, with .00 removed when the rounded value is an integer."""
    text = round_half_up_2(value)
    if text.endswith(".00"):
        return text[:-3]
    return text


def quartiles(values: list[float]) -> tuple[float, float, float]:
    """Return Q1, median, Q3 using the inclusive quartile definition."""
    ordered = sorted(values)
    if len(ordered) == 1:
        only = ordered[0]
        return only, only, only
    q1, median, q3 = statistics.quantiles(ordered, n=4, method="inclusive")
    return q1, median, q3


"""Workbook data kept after the column-A labels have been located."""



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


"""Read a process-capability workbook by the labels in column A."""




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


def workbook_location(path: Path) -> Path:
    """Return the selected path without asking Windows for a final OneDrive path.

    Path.resolve() opens the file to look through reparse points. On a OneDrive
    desktop that open is rejected with Errno 13 before the workbook is read.
    """
    expanded = path.expanduser()
    if expanded.is_absolute():
        return expanded
    return Path.cwd() / expanded


def read_workbook_bytes(path: Path) -> bytes:
    """Read an xlsx even when Excel still has it open."""
    try:
        return path.read_bytes()
    except PermissionError:
        if sys.platform != "win32":
            raise
        return _read_shared_windows(path)


def _read_shared_windows(path: Path) -> bytes:
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.CreateFileW.argtypes = [
        wintypes.LPCWSTR,
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.LPVOID,
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.HANDLE,
    ]
    kernel32.CreateFileW.restype = wintypes.HANDLE
    kernel32.ReadFile.argtypes = [
        wintypes.HANDLE,
        wintypes.LPVOID,
        wintypes.DWORD,
        ctypes.POINTER(wintypes.DWORD),
        wintypes.LPVOID,
    ]
    kernel32.ReadFile.restype = wintypes.BOOL
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
    kernel32.CloseHandle.restype = wintypes.BOOL
    handle = kernel32.CreateFileW(
        str(path),
        0x80000000,
        0x00000001 | 0x00000002 | 0x00000004,
        None,
        3,
        0x80,
        None,
    )
    invalid = wintypes.HANDLE(-1).value
    if handle is None or handle == invalid:
        raise PermissionError(13, "Permission denied", str(path))
    try:
        chunks: list[bytes] = []
        size = 1024 * 1024
        buffer = ctypes.create_string_buffer(size)
        read = wintypes.DWORD(0)
        while True:
            ok = kernel32.ReadFile(handle, buffer, size, ctypes.byref(read), None)
            if not ok or read.value == 0:
                if not ok:
                    raise PermissionError(13, "Permission denied", str(path))
                break
            chunks.append(buffer.raw[: read.value])
        return b"".join(chunks)
    finally:
        kernel32.CloseHandle(handle)


def load_report(path: Path) -> Report:
    try:
        data = read_workbook_bytes(path)
    except PermissionError as exc:
        raise TemplateError(
            "無法讀取這個 Excel。請先關閉 Excel 裡已打開的同一份檔案；"
            "若檔案在 OneDrive，請在檔案總管對它按右鍵，選「一律保留在此裝置上」，然後再執行一次。"
        ) from exc
    except OSError as exc:
        raise TemplateError(f"找不到檔案：{path}") from exc
    try:
        workbook = load_workbook(BytesIO(data), data_only=False, read_only=False)
    except (InvalidFileException, zipfile.BadZipFile) as exc:
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


"""Default text for the two editable fields on the picker page."""


NOT_PROVIDED = "Not provided"


def blank_to_not_provided(value: object) -> str:
    if value is None:
        return NOT_PROVIDED
    text = str(value).strip()
    return text if text else NOT_PROVIDED


def default_voltage(group: SeqGroup) -> str:
    return group.ext_name if group.ext_name else NOT_PROVIDED


# TIName prefix, then the VarNames that start checked. Anything else starts off.
# A trailing marker such as "(M)" or "_2" still matches the prefix.
_PRESET_CHECKS: tuple[tuple[str, frozenset[str]], ...] = (
    ("static load test", frozenset({"Vout_1", "Vout_2", "Vnpp_1", "Vnpp_2"})),
    ("turn on & sequence test", frozenset({"Ton_1", "Ton_2"})),
    ("hold up & sequence test", frozenset({"Thd_1", "Thd_2"})),
    (
        "line regulation test",
        frozenset({"Vout_1_1", "Vout_1_2", "Vout_2_1", "Vout_2_2", "Vout_3_1", "Vout_3_2"}),
    ),
    (
        "load regulation test",
        frozenset({"Vout_1_1", "Vout_1_2", "Vout_2_1", "Vout_2_2", "Vout_3_1", "Vout_3_2"}),
    ),
    (
        "short circuit protection test",
        frozenset({"Vdisable_1", "Vdisable_2", "Rec_Vout_1", "Rec_Vout_2"}),
    ),
    ("efficiency & power factor test", frozenset({"PF_1", "Eff_1"})),
    ("over load protection test", frozenset({"OLP_Trip_1"})),
    (
        "dynamic load regulation test",
        frozenset({"Vpk+_1", "Vpk+_2", "Vpk-_1", "Vpk-_2"}),
    ),
)


def default_checked(ti_name: str, variable_name: str) -> bool:
    folded = " ".join(ti_name.casefold().split())
    for prefix, names in _PRESET_CHECKS:
        if folded.startswith(prefix):
            return variable_name in names
    return False


def default_load(group: SeqGroup) -> str:
    """Upper Spec of every VarName containing Iout, then a trailing A."""
    parts: list[str] = []
    for variable in group.variables:
        if "iout" not in variable.name.casefold():
            continue
        if variable.upper is None:
            continue
        parts.append(format_load_number(variable.upper))
    if not parts:
        return NOT_PROVIDED
    return " / ".join(parts) + " A"


"""Turn a workbook and the picker's choices into the slide contents."""




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
    voltage_range: str = NOT_PROVIDED
    full_load_current: str = NOT_PROVIDED


@dataclass
class GroupChoice:
    key: str
    input_voltage: str
    load_current: str
    checked_uids: set[str]


def assemble(
    report: Report,
    choices: list[GroupChoice],
    voltage_range: str = "",
    full_load_current: str = "",
) -> Deck:
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
        voltage_range=_cover_text(voltage_range),
        full_load_current=_cover_text(full_load_current),
    )


def _cover_text(value: str) -> str:
    text = value.strip()
    return text if text else NOT_PROVIDED


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
        return round_half_up(float(value))
    if isinstance(value, str):
        return value.strip()
    return str(value).strip()


"""Persist checkbox and text-field choices beside the workbook."""




def selection_path(workbook: Path) -> Path:
    return workbook.with_suffix(".selection.json")


def load_saved_groups(path: Path) -> dict:
    if not path.is_file():
        return {}
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        print(f"無法讀取勾選檔，改用預設值：{path}", file=sys.stderr)
        return {}
    if not isinstance(payload, dict) or payload.get("version") != 2:
        return {}
    groups = payload.get("groups")
    if not isinstance(groups, dict):
        return {}
    return groups


def load_saved_cover(path: Path) -> tuple[str, str]:
    if not path.is_file():
        return "", ""
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return "", ""
    if not isinstance(payload, dict) or payload.get("version") != 2:
        return "", ""
    return _optional_text(payload.get("voltage_range")), _optional_text(payload.get("full_load_current"))


def _optional_text(value: object) -> str:
    if not isinstance(value, str):
        return ""
    return value.strip()


def page_model(
    report: Report,
    saved_groups: dict,
    voltage_range: str = "",
    full_load_current: str = "",
) -> dict:
    groups = []
    for group in report.picker_groups():
        saved = saved_groups.get(group.key)
        voltage_default = default_voltage(group)
        load_default = default_load(group)
        previous = None
        if isinstance(saved, dict):
            voltage = _stored_text(saved.get("input_voltage"), voltage_default)
            load = _stored_text(saved.get("load_current"), load_default)
            if isinstance(saved.get("variables"), dict):
                previous = saved["variables"]
        else:
            voltage = voltage_default
            load = load_default
        variables = []
        for variable in group.plottable_variables:
            if previous is not None and variable.uid in previous:
                checked = bool(previous[variable.uid])
            else:
                checked = default_checked(group.ti_name, variable.name)
            variables.append(
                {"uid": variable.uid, "name": variable.name, "checked": checked}
            )
        groups.append(
            {
                "key": group.key,
                "seq": group.seq,
                "tiName": group.ti_name,
                "extName": group.ext_name,
                "inputVoltage": voltage,
                "inputDefault": voltage_default,
                "loadCurrent": load,
                "loadDefault": load_default,
                "variables": variables,
            }
        )
    return {
        "fileName": report.source_name,
        "voltageRange": voltage_range,
        "fullLoadCurrent": full_load_current,
        "groups": groups,
    }


def choices_from_payload(report: Report, payload: object) -> list[GroupChoice]:
    incoming = _payload_groups(payload)
    choices: list[GroupChoice] = []
    for group in report.picker_groups():
        item = incoming.get(group.key)
        known = {variable.uid for variable in group.plottable_variables}
        if not isinstance(item, dict):
            choices.append(
                GroupChoice(
                    key=group.key,
                    input_voltage=default_voltage(group),
                    load_current=default_load(group),
                    checked_uids=set(),
                )
            )
            continue
        raw_checked = item.get("checked")
        checked: set[str] = set()
        if isinstance(raw_checked, list):
            checked = {uid for uid in raw_checked if isinstance(uid, str) and uid in known}
        choices.append(
            GroupChoice(
                key=group.key,
                input_voltage=blank_to_not_provided(_text_field(item.get("input_voltage"))),
                load_current=blank_to_not_provided(_text_field(item.get("load_current"))),
                checked_uids=checked,
            )
        )
    return choices


def save_choices(
    path: Path,
    report: Report,
    choices: list[GroupChoice],
    voltage_range: str = "",
    full_load_current: str = "",
) -> None:
    by_key = {choice.key: choice for choice in choices}
    groups = {}
    for group in report.picker_groups():
        choice = by_key.get(group.key)
        if choice is None:
            continue
        groups[group.key] = {
            "input_voltage": choice.input_voltage,
            "load_current": choice.load_current,
            "variables": {
                variable.uid: variable.uid in choice.checked_uids
                for variable in group.plottable_variables
            },
        }
    _write_json(
        path,
        {
            "version": 2,
            "voltage_range": voltage_range,
            "full_load_current": full_load_current,
            "groups": groups,
        },
    )


def _stored_text(value: object, default: str) -> str:
    if value is None:
        return default
    if isinstance(value, str) and value.strip() == "":
        return NOT_PROVIDED
    return str(value).strip()


def _text_field(value: object) -> str:
    if value is None:
        return ""
    return str(value)


def _payload_groups(payload: object) -> dict:
    if not isinstance(payload, dict):
        return {}
    groups = payload.get("groups")
    if not isinstance(groups, list):
        return {}
    found = {}
    for item in groups:
        if isinstance(item, dict) and isinstance(item.get("key"), str):
            found[item["key"]] = item
    return found


def _write_json(path: Path, payload: dict) -> None:
    partial = path.with_suffix(path.suffix + ".partial")
    try:
        partial.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        partial.replace(path)
    except Exception:
        partial.unlink(missing_ok=True)
        raise


"""Render the assembled deck to a 16:9 PowerPoint file."""




SLIDE_W = 13.3333333333
SLIDE_H = 7.5
INK = RGBColor(0x1A, 0x1A, 0x1A)
GRAY_BOX = RGBColor(0xBF, 0xBF, 0xBF)
GRAY_LINE = RGBColor(0x7F, 0x7F, 0x7F)
DARK = RGBColor(0x40, 0x40, 0x40)
POINT = RGBColor(0x1F, 0x4E, 0x79)
SPEC = RGBColor(0xFF, 0x00, 0x00)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
HEADER_FILL = RGBColor(0xF2, 0xF2, 0xF2)
STAT_HEADERS = (
    "VarName",
    "Upper Spec",
    "Lower Spec",
    "Mean",
    "Ca",
    "Cpk",
    "Cp",
)


def write_pptx(deck: Deck, dest: Path) -> None:
    presentation = build_presentation(deck)
    partial = dest.with_suffix(dest.suffix + ".partial")
    try:
        presentation.save(partial)
        partial.replace(dest)
    except Exception:
        partial.unlink(missing_ok=True)
        raise


def build_presentation(deck: Deck) -> Presentation:
    presentation = Presentation()
    presentation.slide_width = Emu(12192000)
    presentation.slide_height = Emu(6858000)
    blank = presentation.slide_layouts[6]
    _render_summary(presentation.slides.add_slide(blank), deck)
    for content in deck.slides:
        _render_content(presentation, blank, content)
    return presentation


def _render_summary(slide, deck: Deck) -> None:
    _white_background(slide)
    margin = 0.4
    top = 0.3
    width = SLIDE_W - 2 * margin
    date_text = ", ".join(deck.dates)
    meta_height = 1.72
    label_width = 2.45
    _add_table(
        slide,
        margin,
        top,
        width,
        meta_height,
        [
            [Path(deck.source_name).stem, ""],
            ["Input voltage range", deck.voltage_range],
            ["Full load current", deck.full_load_current],
            ["Date", date_text],
            ["N", str(deck.unit_count)],
        ],
        font_pt=13,
        widths=[label_width, width - label_width],
        header=False,
        span_top=True,
    )
    lists_top = top + meta_height + 0.16
    lists_height = SLIDE_H - 0.3 - lists_top
    gutter = 0.22
    left_width = 4.2
    right_width = width - left_width - gutter
    serial_rows = [["S/N"], *([serial] for serial in deck.serials)]
    test_rows = [["SEQ", "TIName", "Input voltage"]]
    test_rows.extend([row.seq, row.ti_name, row.input_voltage] for row in deck.rows)
    row_height = lists_height / max(len(serial_rows), len(test_rows), 1)
    font_pt = max(6, min(14, row_height * 72 * 0.62))
    _add_table(
        slide,
        margin,
        lists_top,
        left_width,
        row_height * len(serial_rows),
        serial_rows,
        font_pt=font_pt,
        widths=[left_width],
    )
    _add_table(
        slide,
        margin + left_width + gutter,
        lists_top,
        right_width,
        row_height * len(test_rows),
        test_rows,
        font_pt=min(font_pt, 12),
        widths=[right_width * 0.14, right_width * 0.46, right_width * 0.40],
    )


CHART_COLUMNS = 3
CHART_ROWS = 2
CHARTS_PER_SLIDE = CHART_COLUMNS * CHART_ROWS


def _content_frame() -> dict[str, float]:
    """Header, table, and a fixed 2-by-3 chart band. Chart cells never change size."""
    margin_x = 0.32
    margin_y = 0.18
    width = SLIDE_W - 2 * margin_x
    header_h = 0.70
    table_top = margin_y + header_h + 0.04
    chart_h = 4.36
    chart_top = SLIDE_H - margin_y - chart_h
    table_h = chart_top - 0.08 - table_top
    return {
        "margin_x": margin_x,
        "margin_y": margin_y,
        "width": width,
        "header_h": header_h,
        "table_top": table_top,
        "table_h": table_h,
        "chart_top": chart_top,
        "chart_h": chart_h,
    }


def _render_content(presentation: Presentation, blank, content: ContentSlide) -> None:
    variables = content.variables
    pages = [variables[start : start + CHARTS_PER_SLIDE] for start in range(0, len(variables), CHARTS_PER_SLIDE)]
    if not pages:
        pages = [[]]
    for page, chunk in enumerate(pages):
        slide = presentation.slides.add_slide(blank)
        _render_content_slide(slide, content, chunk, include_table=(page == 0))


def _render_content_slide(slide, content: ContentSlide, variables: list[SlideVariable], include_table: bool) -> None:
    _white_background(slide)
    frame = _content_frame()
    _write_lines(
        slide,
        frame["margin_x"],
        frame["margin_y"],
        frame["width"],
        frame["header_h"],
        [
            (content.title, 16, True),
            (f"Input voltage: {content.input_voltage}", 12, False),
            (f"Load current: {content.load_current}", 12, False),
        ],
    )
    if include_table:
        rows = [list(STAT_HEADERS)]
        for variable in content.variables:
            rows.append(
                [
                    variable.name,
                    _spec_text(variable.upper),
                    _spec_text(variable.lower),
                    variable.mean,
                    variable.ca,
                    variable.cpk,
                    variable.cp,
                ]
            )
        table_height, font_pt = _compact_table_size(len(rows), frame["table_h"])
        _add_table(
            slide,
            frame["margin_x"],
            frame["table_top"],
            frame["width"],
            table_height,
            rows,
            font_pt=font_pt,
            widths=[frame["width"] * part for part in (0.20, 0.14, 0.14, 0.13, 0.13, 0.13, 0.13)],
        )
    _draw_chart_grid(
        slide,
        frame["margin_x"],
        frame["chart_top"],
        frame["width"],
        frame["chart_h"],
        variables,
    )


def _compact_table_size(row_count: int, max_height: float, max_pt: float = 11) -> tuple[float, float]:
    """Row height stays close to the text unless the rows must shrink to fit."""
    font_pt = float(max_pt)
    row_in = font_pt / 72 + 0.10
    if row_count * row_in <= max_height:
        return row_count * row_in, font_pt
    font_pt = _fit_font(row_count, max_height, max_pt, 6)
    return max_height, font_pt


def _draw_chart_grid(slide, left, top, width, height, variables: list[SlideVariable]) -> None:
    if height <= 0:
        return
    cell_w = width / CHART_COLUMNS
    cell_h = height / CHART_ROWS
    for index, variable in enumerate(variables[:CHARTS_PER_SLIDE]):
        column = index % CHART_COLUMNS
        row = index // CHART_COLUMNS
        _draw_boxplot(
            slide,
            left + column * cell_w,
            top + row * cell_h,
            cell_w,
            cell_h,
            variable.name,
            variable.values,
            variable.lower,
            variable.upper,
        )


def _draw_boxplot(slide, left, top, width, height, title, values, lower, upper) -> None:
    pad = 0.03
    gutter = min(0.52, max(0.34, width * 0.24))
    title_h = min(0.22, height * 0.18)
    _write_lines(
        slide,
        left + gutter,
        top,
        max(0.2, width - gutter - pad),
        title_h,
        [(title, 8 if height >= 0.9 else 7, True)],
    )
    plot_left = left + gutter
    plot_top = top + title_h
    plot_w = max(0.2, width - gutter - pad)
    # Room under the plot so a lower-limit label is not covered by the next chart.
    footer = 0.24
    plot_h = max(0.2, height - title_h - pad - footer)
    lo, hi = _axis_limits(values)
    q1, median, q3 = quartiles(values)

    def y_of(value: float) -> float:
        fraction = (value - lo) / (hi - lo)
        return plot_top + (1 - fraction) * plot_h

    _vline(slide, plot_left, plot_top, plot_top + plot_h, GRAY_LINE, 0.75)
    _draw_axis_labels(slide, left, plot_left, plot_top, plot_h, lo, hi, y_of)
    center = plot_left + plot_w * 0.5
    box_w = max(0.08, plot_w * 0.28)
    y_min = y_of(min(values))
    y_max = y_of(max(values))
    _hline(slide, center - box_w * 0.25, y_min, center + box_w * 0.25, DARK, 1)
    _hline(slide, center - box_w * 0.25, y_max, center + box_w * 0.25, DARK, 1)
    _vline(slide, center, y_min, y_max, DARK, 1)
    y_q1 = y_of(q1)
    y_q3 = y_of(q3)
    box_top = min(y_q1, y_q3)
    box_h = abs(y_q1 - y_q3)
    if box_h < 0.015:
        box_top -= 0.007
        box_h = 0.015
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(center - box_w / 2),
        Inches(box_top),
        Inches(box_w),
        Inches(box_h),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = GRAY_BOX
    shape.line.color.rgb = GRAY_LINE
    shape.line.width = Pt(0.75)
    shape.shadow.inherit = False
    _hline(slide, center - box_w / 2, y_of(median), center + box_w / 2, DARK, 1.25)
    radius = min(0.035, plot_w * 0.035)
    for value in values:
        point_y = y_of(value)
        dot = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            Inches(center - radius),
            Inches(point_y - radius),
            Inches(radius * 2),
            Inches(radius * 2),
        )
        dot.fill.solid()
        dot.fill.fore_color.rgb = POINT
        dot.line.fill.background()
        dot.shadow.inherit = False
    _mark_limit(slide, plot_left, plot_top, plot_w, plot_h, y_of, lo, hi, upper, "U", above=True)
    _mark_limit(slide, plot_left, plot_top, plot_w, plot_h, y_of, lo, hi, lower, "L", above=False, footer=footer)


def _mark_limit(
    slide,
    plot_left,
    plot_top,
    plot_w,
    plot_h,
    y_of,
    lo,
    hi,
    limit,
    tag: str,
    above: bool,
    footer: float = 0.24,
) -> None:
    if limit is None:
        return
    label = f"{tag}: {round_half_up(limit)}"
    label_h = 0.16
    label_w = min(1.2, plot_w * 0.46)
    below_plot = plot_top + plot_h + 0.04
    if lo <= limit <= hi:
        y_pos = y_of(limit)
        _hline(slide, plot_left, y_pos, plot_left + plot_w, SPEC, 1.5)
        raw_top = y_pos - label_h - 0.01 if above else y_pos + 0.04
        if raw_top + label_h > plot_top + plot_h:
            raw_top = below_plot
    elif limit > hi:
        raw_top = plot_top
        _hline(slide, plot_left, plot_top + 0.02, plot_left + min(0.7, plot_w), SPEC, 1.5)
    else:
        edge = plot_top + plot_h
        _hline(slide, plot_left, edge, plot_left + min(0.7, plot_w), SPEC, 1.5)
        raw_top = below_plot
    floor = plot_top if raw_top < plot_top + plot_h else plot_top + plot_h
    ceiling = plot_top + plot_h + footer - label_h
    top = min(max(raw_top, floor), ceiling)
    _write_lines(
        slide,
        plot_left + 0.05,
        top,
        label_w,
        label_h,
        [(label, 8, True)],
        color=SPEC,
        wrap=False,
    )


def _axis_limits(values: list[float]) -> tuple[float, float]:
    """Fit the axis to the readings so the box and points use the plot height."""
    lo = min(values)
    hi = max(values)
    span = hi - lo
    if span == 0:
        pad = max(abs(lo) * 0.05, 0.05)
        return lo - pad, hi + pad
    pad = span * 0.12
    return lo - pad, hi + pad


def _nice_ticks(lo: float, hi: float, maximum: int = 4) -> list[float]:
    span = hi - lo
    if span <= 0:
        return [lo]
    raw = span / max(maximum, 1)
    exponent = math.floor(math.log10(raw))
    magnitude = 10.0 ** exponent
    step = 10 * magnitude
    for factor in (1, 2, 2.5, 5, 10):
        candidate = factor * magnitude
        if raw <= candidate * 1.001:
            step = candidate
            break
    start = math.ceil((lo - step * 1e-9) / step) * step
    ticks: list[float] = []
    value = start
    for _ in range(12):
        if value > hi + step * 1e-9:
            break
        if value >= lo - step * 1e-9:
            ticks.append(value)
        value += step
    return ticks or [lo, hi]


def _draw_axis_labels(slide, left, plot_left, plot_top, plot_h, lo, hi, y_of) -> None:
    span = hi - lo
    kept: list[float] = []
    last_y: float | None = None
    for value in _nice_ticks(lo, hi):
        y_pos = y_of(value)
        if last_y is not None and abs(y_pos - last_y) < 0.16:
            continue
        kept.append(value)
        last_y = y_pos
    label_w = max(0.2, plot_left - left - 0.03)
    label_h = 0.16
    seen: set[str] = set()
    for value in kept:
        text = round_half_up(value)
        if text in seen:
            continue
        seen.add(text)
        y_pos = y_of(value)
        top = min(max(y_pos - label_h / 2, plot_top), plot_top + plot_h - label_h)
        _write_lines(
            slide,
            left,
            top,
            label_w,
            label_h,
            [(text, 7, False)],
            alignment=PP_ALIGN.RIGHT,
            wrap=False,
        )
        _hline(slide, plot_left, y_pos, plot_left + 0.045, GRAY_LINE, 0.75)


def _hline(slide, x1, y1, x2, color, width_pt) -> None:
    _line(slide, x1, y1, x2, y1, color, width_pt)


def _vline(slide, x_pos, y1, y2, color, width_pt) -> None:
    _line(slide, x_pos, y1, x_pos, y2, color, width_pt)


def _line(slide, x1, y1, x2, y2, color, width_pt) -> None:
    if abs(x1 - x2) < 1e-6 and abs(y1 - y2) < 1e-6:
        return
    line = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2)
    )
    line.line.color.rgb = color
    line.line.width = Pt(width_pt)


def _spec_text(value: float | None) -> str:
    if value is None:
        return ""
    return round_half_up(value)


def _fit_font(row_count: int, height_in: float, max_pt: float, min_pt: float = 6) -> float:
    if row_count <= 0:
        return max_pt
    raw = (height_in / row_count) * 72 * 0.62
    return max(min_pt, min(max_pt, raw))


def _write_lines(
    slide,
    left,
    top,
    width,
    height,
    lines: list[tuple[str, float, bool]],
    alignment=PP_ALIGN.LEFT,
    wrap: bool = True,
    color: RGBColor = INK,
):
    shape = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    frame = shape.text_frame
    frame.word_wrap = wrap
    frame.auto_size = None
    for index, (text, size, bold) in enumerate(lines):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.alignment = alignment
        paragraph.space_before = Pt(0)
        paragraph.space_after = Pt(0)
        run = paragraph.add_run()
        run.text = text
        _style_run(run, size, bold, color)
    return shape


def _add_table(slide, left, top, width, height, rows, font_pt, widths, header=True, span_top=False):
    shape = slide.shapes.add_table(
        len(rows), len(rows[0]), Inches(left), Inches(top), Inches(width), Inches(height)
    )
    _remove_table_style(shape)
    table = shape.table
    row_height = int(Inches(height) / len(rows))
    for row in table.rows:
        row.height = row_height
    for index, column_width in enumerate(widths):
        table.columns[index].width = Inches(column_width)
    for row_index, row_values in enumerate(rows):
        for column_index, text in enumerate(row_values):
            is_header = header and row_index == 0
            _fill_cell(
                table.cell(row_index, column_index),
                text,
                font_pt,
                bold=is_header or (span_top and row_index == 0),
                fill=HEADER_FILL if is_header else WHITE,
            )
    if span_top and len(rows[0]) > 1:
        table.cell(0, 0).merge(table.cell(0, len(rows[0]) - 1))
    return shape


def _fill_cell(cell, text: str, font_pt: float, bold: bool, fill: RGBColor) -> None:
    cell.text = text
    frame = cell.text_frame
    frame.word_wrap = True
    paragraph = frame.paragraphs[0]
    paragraph.alignment = PP_ALIGN.LEFT
    paragraph.space_before = Pt(0)
    paragraph.space_after = Pt(0)
    for run in paragraph.runs:
        _style_run(run, font_pt, bold)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.margin_left = Inches(0.04)
    cell.margin_right = Inches(0.03)
    cell.margin_top = Inches(0.01)
    cell.margin_bottom = Inches(0.01)
    cell.fill.solid()
    cell.fill.fore_color.rgb = fill
    _set_cell_border(cell)


def _style_run(run, size_pt: float, bold: bool, color: RGBColor = INK) -> None:
    run.font.name = "Calibri"
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = False
    run.font.color.rgb = color
    properties = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        element = properties.find(qn(tag))
        if element is None:
            element = etree.SubElement(properties, qn(tag))
        element.set("typeface", "Calibri")


def _set_cell_border(cell, color: str = "A6A6A6") -> None:
    properties = cell._tc.get_or_add_tcPr()
    for edge in ("lnL", "lnR", "lnT", "lnB"):
        tag = qn(f"a:{edge}")
        for old in properties.findall(tag):
            properties.remove(old)
        line = etree.SubElement(properties, tag)
        line.set("w", "6350")
        line.set("cap", "flat")
        line.set("cmpd", "sng")
        line.set("algn", "ctr")
        fill = etree.SubElement(line, qn("a:solidFill"))
        color_el = etree.SubElement(fill, qn("a:srgbClr"))
        color_el.set("val", color)
        etree.SubElement(line, qn("a:prstDash")).set("val", "solid")


def _remove_table_style(shape) -> None:
    table_properties = shape._element.graphic.graphicData.tbl.tblPr
    if table_properties is None:
        return
    style = table_properties.find(qn("a:tableStyleId"))
    if style is not None:
        table_properties.remove(style)


def _white_background(slide) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = WHITE


"""Local page that collects choices and writes the PowerPoint file."""





def serve(report: Report, workbook: Path, open_browser: bool = True) -> int:
    server = make_server(report, workbook)
    url = f"http://127.0.0.1:{server.server_address[1]}/"
    print(f"已讀取：{report.source_name}")
    print(f"勾選頁：{url}")
    print("按 Ctrl+C 結束。")
    if open_browser:
        try:
            webbrowser.open(url)
        except OSError:
            pass
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n已結束。")
    finally:
        server.server_close()
    return 0


def make_server(report: Report, workbook: Path) -> ThreadingHTTPServer:
    handler = _handler_class(report, workbook_location(workbook))
    return ThreadingHTTPServer(("127.0.0.1", 0), handler)


def _handler_class(report: Report, workbook: Path):
    class Handler(BaseHTTPRequestHandler):
        def do_GET(self):
            if self.path.split("?", 1)[0] not in {"/", "/index.html"}:
                self.send_error(404)
                return
            selection = selection_path(workbook)
            voltage_range, full_load = load_saved_cover(selection)
            model = page_model(report, load_saved_groups(selection), voltage_range, full_load)
            body = PAGE.replace("__REPORT_JSON__", _json_for_html(model)).encode("utf-8")
            self._send(200, "text/html; charset=utf-8", body)

        def do_POST(self):
            if self.path.split("?", 1)[0] != "/api/generate":
                self.send_error(404)
                return
            length = int(self.headers.get("Content-Length", "0"))
            if length > 2_000_000:
                self._send_json(413, {"ok": False, "error": "勾選內容過大"})
                return
            raw = self.rfile.read(length)
            try:
                payload = json.loads(raw.decode("utf-8"))
                choices = choices_from_payload(report, payload)
                voltage_range = _optional_text(payload.get("voltage_range")) if isinstance(payload, dict) else ""
                full_load = _optional_text(payload.get("full_load_current")) if isinstance(payload, dict) else ""
                deck = assemble(report, choices, voltage_range, full_load)
                pptx_path = workbook.with_suffix(".pptx")
                saved_path = selection_path(workbook)
                write_pptx(deck, pptx_path)
                save_choices(saved_path, report, choices, voltage_range, full_load)
            except Exception as exc:
                traceback.print_exc()
                self._send_json(500, {"ok": False, "error": str(exc)})
                return
            self._send_json(
                200,
                {"ok": True, "path": str(pptx_path), "selectionPath": str(saved_path)},
            )

        def log_message(self, fmt, *args):
            print("[%s] %s" % (self.log_date_time_string(), fmt % args))

        def _send(self, status: int, content_type: str, body: bytes) -> None:
            self.send_response(status)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def _send_json(self, status: int, payload: dict) -> None:
            self._send(status, "application/json; charset=utf-8", json.dumps(payload).encode("utf-8"))

    return Handler


def _json_for_html(payload: dict) -> str:
    return json.dumps(payload, ensure_ascii=False).replace("<", "\\u003c")


"""Command line entry for the workbook-to-PowerPoint tool."""


def main(argv: list[str] | None = None, choose_workbook: Callable[[], Path | None] | None = None) -> int:
    parser = argparse.ArgumentParser(description="將電源測試 Excel 轉成 PPT。先開啟勾選頁，再產生報告。")
    parser.add_argument("workbook", nargs="?", type=Path, help="電源測試 Excel。省略時會跳出選檔視窗。")
    args = parser.parse_args(argv)
    if args.workbook is None:
        chooser = choose_workbook or pick_workbook
        print("請在視窗中選擇電源測試 Excel。")
        picked = chooser()
        if picked is None:
            print("未選擇檔案。", file=sys.stderr)
            return 1
        workbook = workbook_location(picked)
    else:
        workbook = workbook_location(args.workbook)
    try:
        report = load_report(workbook)
    except TemplateError as exc:
        print(exc, file=sys.stderr)
        return 1
    except InvalidFileException:
        print(f"無法讀取 Excel：{workbook.name}", file=sys.stderr)
        return 1
    return serve(report, workbook)


def pick_workbook() -> Path | None:
    if tk is None or filedialog is None:
        print("無法開啟選檔視窗。請改為：python run_report.py 你的檔案.xlsx", file=sys.stderr)
        return None
    window = tk.Tk()
    window.withdraw()
    try:
        window.attributes("-topmost", True)
    except tk.TclError:
        pass
    chosen = filedialog.askopenfilename(
        title="選擇電源測試 Excel",
        filetypes=[("Excel", "*.xlsx"), ("所有檔案", "*.*")],
    )
    window.destroy()
    if not chosen:
        return None
    return Path(chosen)


def launch() -> int:
    try:
        code = main()
    except Exception as exc:
        print(f"發生錯誤：{exc}")
        code = 1
    if code or len(sys.argv) <= 1:
        try:
            input("\n按 Enter 關閉...")
        except EOFError:
            pass
    return code


if __name__ == "__main__":
    raise SystemExit(launch())
