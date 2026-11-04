"""Persist checkbox and text-field choices beside the workbook."""

import json
import sys
from pathlib import Path

from power_report.assemble import GroupChoice
from power_report.defaults import (
    NOT_PROVIDED,
    blank_to_not_provided,
    default_load,
    default_voltage,
)
from power_report.model import Report


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
    groups = payload.get("groups") if isinstance(payload, dict) else None
    if not isinstance(groups, dict):
        return {}
    return groups


def page_model(report: Report, saved_groups: dict) -> dict:
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
            if previous is None or variable.uid not in previous:
                checked = True
            else:
                checked = bool(previous[variable.uid])
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
    return {"fileName": report.source_name, "groups": groups}


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


def save_choices(path: Path, report: Report, choices: list[GroupChoice]) -> None:
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
    _write_json(path, {"version": 1, "groups": groups})


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

