import json
import tempfile
import threading
import unittest
import urllib.request
from pathlib import Path

import run_report

from openpyxl import Workbook
from pptx import Presentation

from run_report import (
    TemplateError,
    assemble,
    build_presentation,
    choices_from_payload,
    _axis_limits,
    default_checked,
    default_load,
    default_voltage,
    load_report,
    main,
    make_server,
    page_model,
    quartiles,
    round_half_up_2,
    save_choices,
    selection_path,
)

SAMPLE = Path(__file__).resolve().parents[1] / "TN-5516-8PoE-150W HV.xlsx"
LABELS = {
    "Date": 4,
    "SEQNo": 5,
    "TIName": 6,
    "Ext. Name": 7,
    "VarName": 8,
    "Upper Spec": 12,
    "Lower Spec": 13,
    "Samples": 14,
    "Mean": 15,
    "Ca": 16,
    "Cpk": 17,
    "Cp": 18,
    "S/N": 30,
}


def write_fixture(path: Path, columns: list[dict], serials: list[str], omit: str | None = None) -> None:
    workbook = Workbook()
    sheet = workbook.active
    for label, row in LABELS.items():
        if label == omit:
            continue
        sheet.cell(row, 1, label)
    sheet.cell(31, 1, serials[0])
    sheet.cell(33, 1, serials[1] if len(serials) > 1 else None)
    if len(serials) > 2:
        raise AssertionError("fixture helper only places two serial rows")
    for offset, column in enumerate(columns, start=2):
        sheet.cell(LABELS["SEQNo"], offset, column.get("seq"))
        sheet.cell(LABELS["TIName"], offset, column.get("ti"))
        sheet.cell(LABELS["Ext. Name"], offset, column.get("ext"))
        sheet.cell(LABELS["VarName"], offset, column.get("var"))
        sheet.cell(LABELS["Date"], offset, column.get("date", "2026/1/1~2026/1/2"))
        sheet.cell(LABELS["Upper Spec"], offset, column.get("upper"))
        sheet.cell(LABELS["Lower Spec"], offset, column.get("lower"))
        sheet.cell(LABELS["Samples"], offset, column.get("samples", 99))
        sheet.cell(LABELS["Mean"], offset, column.get("mean", 1))
        sheet.cell(LABELS["Ca"], offset, column.get("ca", 0.1))
        sheet.cell(LABELS["Cpk"], offset, column.get("cpk", 1.2))
        sheet.cell(LABELS["Cp"], offset, column.get("cp", 1.3))
        values = column.get("values", [1, 2])
        if len(serials) > 0:
            sheet.cell(31, offset, values[0] if values else None)
        if len(serials) > 1:
            sheet.cell(33, offset, values[1] if len(values) > 1 else None)
    workbook.save(path)
    workbook.close()


def _assert_shapes_inside(presentation, slide) -> None:
    tolerance = 20000
    for shape in slide.shapes:
        right = shape.left + shape.width
        bottom = shape.top + shape.height
        if shape.left < -tolerance or shape.top < -tolerance:
            raise AssertionError(f"shape starts off the slide: {shape.left}, {shape.top}")
        if right > presentation.slide_width + tolerance or bottom > presentation.slide_height + tolerance:
            raise AssertionError(f"shape ends off the slide: {right}, {bottom}")


def slide_text(slide) -> str:
    chunks = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            chunks.append(shape.text_frame.text)
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    chunks.append(cell.text)
    return "\n".join(chunks)


class NumberTests(unittest.TestCase):
    def test_round_half_up_always_has_two_decimals(self):
        self.assertEqual(round_half_up_2(1.253), "1.25")
        self.assertEqual(round_half_up_2(1.255), "1.26")
        self.assertEqual(round_half_up_2(0.004248), "0.00")
        self.assertEqual(round_half_up_2(3.0029), "3.00")
        self.assertEqual(round_half_up_2(2.5), "2.50")

    def test_axis_follows_the_readings_rather_than_a_distant_spec(self):
        lo, hi = _axis_limits([4.9753, 5.0058])
        self.assertLess(hi - lo, 0.08)
        self.assertGreater(lo, 4.9)
        self.assertLess(hi, 5.1)

    def test_inclusive_quartiles(self):
        self.assertEqual(quartiles([1, 2, 3, 4]), (1.75, 2.5, 3.25))
        self.assertEqual(quartiles([4]), (4, 4, 4))


class ParseTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.path = Path(self.directory.name) / "fixture.xlsx"

    def tearDown(self):
        self.directory.cleanup()

    def test_labels_serial_gap_and_seq_runs(self):
        write_fixture(
            self.path,
            [
                {"seq": 1, "ti": "Static Load Test", "ext": "90VAC", "var": "Vout_1", "upper": 5.25, "lower": 4.75, "mean": 4.5, "values": [5, 5.1]},
                {"seq": 1, "ti": "Static Load Test", "ext": "90VAC", "var": "Iout_1", "upper": 3.0029, "lower": None, "values": [3, 3]},
                {"seq": None, "ti": "ignored", "ext": "1VAC", "var": "Vout_1", "upper": 9, "values": [1, 1]},
                {"seq": 1, "ti": "Static Load Test", "ext": "90VAC", "var": "Vout_1", "upper": 5.25, "lower": 4.75, "values": [4, 4]},
                {"seq": 2, "ti": "Hold Up", "ext": None, "var": "iout_a", "upper": 1.253, "values": [None, None]},
                {"seq": 2, "ti": "Hold Up", "ext": None, "var": "Thd_1", "upper": "      ------", "lower": "------", "values": [8, 9]},
            ],
            ["SN-A", "SN-B"],
        )
        report = load_report(self.path)
        self.assertEqual(report.serials, ["SN-A", "SN-B"])
        self.assertEqual([group.seq for group in report.groups], ["1", "1", "2"])
        self.assertEqual(default_load(report.groups[0]), "3 A")
        self.assertEqual(default_voltage(report.groups[2]), "Not provided")
        self.assertEqual(default_load(report.groups[2]), "1.25 A")
        self.assertEqual([variable.name for variable in report.groups[2].plottable_variables], ["Thd_1"])
        thd = report.groups[2].plottable_variables[0]
        self.assertIsNone(thd.upper)
        self.assertIsNone(thd.lower)

    def test_permission_denied_explains_the_next_step(self):
        def deny(path: Path) -> bytes:
            raise PermissionError(13, "Permission denied", str(path))

        original = run_report.read_workbook_bytes
        run_report.read_workbook_bytes = deny
        try:
            with self.assertRaises(TemplateError) as caught:
                load_report(Path("locked.xlsx"))
        finally:
            run_report.read_workbook_bytes = original
        self.assertIn("一律保留在此裝置上", str(caught.exception))

    def test_missing_label_names_the_label(self):
        write_fixture(self.path, [{"seq": 1, "ti": "T", "ext": "1VAC", "var": "Vout_1", "values": [1, 2]}], ["SN-A", "SN-B"], omit="Cp")
        with self.assertRaises(TemplateError) as caught:
            load_report(self.path)
        self.assertIn("Cp", str(caught.exception))

    def test_zero_upper_spec_is_a_limit_and_zero_readings_count(self):
        write_fixture(
            self.path,
            [{"seq": 1, "ti": "T", "ext": "", "var": "Vnpp_1", "upper": 0, "lower": None, "values": [0, 0]}],
            ["SN-A", "SN-B"],
        )
        report = load_report(self.path)
        variable = report.groups[0].plottable_variables[0]
        self.assertEqual(variable.upper, 0)
        self.assertEqual(variable.values, [0, 0])
        self.assertEqual(default_voltage(report.groups[0]), "Not provided")


class DeckTests(unittest.TestCase):
    def test_n_is_serial_count_and_blank_field_becomes_not_provided(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        path = Path(directory.name) / "deck.xlsx"
        write_fixture(
            path,
            [
                {"seq": 1, "ti": "Static", "ext": "90VAC", "var": "Vout_1", "upper": 5.25, "lower": 4.75, "mean": 4.6, "date": "2026/1/1~2026/1/2", "values": [1, None]},
                {"seq": 1, "ti": "Static", "ext": "90VAC", "var": "Iout_1", "upper": 1.2, "date": "2026/2/2~2026/2/3", "values": [1, 1]},
            ],
            ["SN-A", "SN-B"],
        )
        # One serial cell is blank in the helper when values[1] is None, but the serial row still exists.
        report = load_report(path)
        self.assertEqual(report.serials, ["SN-A", "SN-B"])
        model = page_model(report, {})
        payload = {
            "groups": [
                {
                    "key": model["groups"][0]["key"],
                    "input_voltage": "   ",
                    "load_current": "1.20",
                    "checked": ["Vout_1"],
                }
            ]
        }
        deck = assemble(report, choices_from_payload(report, payload))
        self.assertEqual(deck.unit_count, 2)
        self.assertEqual(deck.slides[0].variables[0].values, [1])
        self.assertEqual(deck.slides[0].input_voltage, "Not provided")
        self.assertEqual(deck.dates, ["2026/1/1~2026/1/2"])
        self.assertEqual(deck.rows[0].input_voltage, "Not provided")
        text = slide_text(build_presentation(deck).slides[1])
        self.assertIn("Vout_1", text)
        self.assertNotIn("Iout_1", text)
        self.assertNotIn("Samples", text)
        self.assertNotIn("2026/1/1~2026/1/2", text)
        summary_slide = build_presentation(deck).slides[0]
        summary = slide_text(summary_slide)
        self.assertIn("Date", summary)
        self.assertIn("2026/1/1~2026/1/2", summary)
        meta = next(shape.table for shape in summary_slide.shapes if shape.has_table)
        self.assertEqual(meta.rows[0].cells[0].text, "deck")
        self.assertEqual([cell.text for cell in meta.rows[2].cells], ["N", "2"])
        self.assertIn("SN-A", summary)
        self.assertNotIn("Samples", summary)


class SelectionTests(unittest.TestCase):
    def test_saved_choice_wins_and_a_new_variable_uses_the_preset(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        path = Path(directory.name) / "choice.xlsx"
        write_fixture(
            path,
            [
                {"seq": 1, "ti": "Static Load Test", "ext": "90VAC", "var": "Vout_1", "values": [1, 2]},
                {"seq": 1, "ti": "Static Load Test", "ext": "90VAC", "var": "Vout_2", "values": [1, 2]},
                {"seq": 1, "ti": "Static Load Test", "ext": "90VAC", "var": "Iout_1", "upper": 3, "values": [1, 2]},
            ],
            ["SN-A", "SN-B"],
        )
        report = load_report(path)
        choices = choices_from_payload(
            report,
            {"groups": [{"key": report.groups[0].key, "input_voltage": "100VAC", "load_current": "3 A", "checked": ["Vout_1"]}]},
        )
        save_choices(selection_path(path), report, choices)
        saved = json.loads(selection_path(path).read_text(encoding="utf-8"))
        self.assertEqual(saved["version"], 2)
        del saved["groups"][report.groups[0].key]["variables"]["Vout_2"]
        del saved["groups"][report.groups[0].key]["variables"]["Iout_1"]
        model = page_model(report, saved["groups"])
        flags = {item["name"]: item["checked"] for item in model["groups"][0]["variables"]}
        self.assertEqual(flags["Vout_1"], True)
        self.assertEqual(flags["Vout_2"], True)
        self.assertEqual(flags["Iout_1"], False)
        self.assertEqual(model["groups"][0]["inputVoltage"], "100VAC")


class ServerTests(unittest.TestCase):
    def test_generate_writes_pptx_and_selection(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        path = Path(directory.name) / "web.xlsx"
        write_fixture(
            path,
            [{"seq": 1, "ti": "Static", "ext": "90VAC", "var": "Vout_1", "upper": 5.08, "lower": 5.02, "values": [5, 5.1]}],
            ["SN-A", "SN-B"],
        )
        report = load_report(path)
        server = make_server(report, path)
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        self.addCleanup(server.shutdown)
        port = server.server_address[1]
        page = urllib.request.urlopen(f"http://127.0.0.1:{port}/").read().decode("utf-8")
        self.assertIn("產生 PPT", page)
        self.assertIn("Vout_1", page)
        self.assertIn('payload.path + "\\n已儲存勾選："', page)
        script = page.split("<script>\n", 1)[1].split("</script>", 1)[0]
        self.assertNotIn("\n已儲存勾選", script)
        payload = json.dumps(
            {
                "groups": [
                    {
                        "key": report.groups[0].key,
                        "input_voltage": "100VAC",
                        "load_current": "1.50",
                        "checked": ["Vout_1"],
                    }
                ]
            }
        ).encode("utf-8")
        request = urllib.request.Request(
            f"http://127.0.0.1:{port}/api/generate",
            data=payload,
            headers={"Content-Type": "application/json"},
        )
        body = json.loads(urllib.request.urlopen(request).read().decode("utf-8"))
        self.assertTrue(body["ok"])
        presentation = Presentation(body["path"])
        self.assertEqual(len(presentation.slides), 2)
        content = slide_text(presentation.slides[1])
        self.assertIn("100VAC", content)
        self.assertIn("1.50", content)
        self.assertIn("Calibri", presentation.slides[1]._element.xml)
        self.assertIn("FF0000", presentation.slides[1]._element.xml)
        saved = json.loads(Path(body["selectionPath"]).read_text(encoding="utf-8"))
        self.assertEqual(saved["groups"][report.groups[0].key]["input_voltage"], "100VAC")
        server.server_close()


@unittest.skipUnless(SAMPLE.is_file(), "sample workbook is not in the workspace")
class SampleWorkbookTests(unittest.TestCase):
    def test_known_report_shape_and_slide_contents(self):
        report = load_report(SAMPLE)
        self.assertEqual(len(report.serials), 10)
        self.assertEqual(report.serials[0], "TBFED1022848")
        self.assertEqual(len(report.picker_groups()), 25)
        by_seq = {group.seq: group for group in report.groups}
        self.assertEqual(default_voltage(by_seq["1"]), "90VAC")
        self.assertEqual(default_load(by_seq["1"]), "3 / 1.25 A")
        self.assertEqual(default_load(by_seq["6"]), "6 / 2.50 / 6 / 2.50 / 6 / 2.50 A")
        self.assertEqual(default_load(by_seq["18"]), "0.10 / 0.10 A")
        self.assertEqual(default_load(by_seq["25"]), "0 / 0 A")
        self.assertEqual(default_voltage(by_seq["2"]), "Not provided")
        self.assertEqual(default_load(by_seq["2"]), "Not provided")
        vout = next(variable for variable in by_seq["1"].variables if variable.name == "Vout_1")
        self.assertNotEqual(round(sum(vout.values) / len(vout.values), 3), round(float(vout.mean), 3))
        model = page_model(report, {})
        by_model = {group["seq"]: group for group in model["groups"]}
        static_flags = {item["name"]: item["checked"] for item in by_model["1"]["variables"]}
        self.assertEqual(static_flags["Vout_1"], True)
        self.assertEqual(static_flags["Vout_2"], True)
        self.assertEqual(static_flags["Vnpp_1"], True)
        self.assertEqual(static_flags["Vnpp_2"], True)
        self.assertEqual(static_flags["Iout_1"], False)
        turn_on = {item["name"]: item["checked"] for item in by_model["2"]["variables"]}
        self.assertEqual(turn_on["Ton_1"], True)
        self.assertEqual(turn_on["Ton_2"], True)
        self.assertEqual(turn_on["Trise_1"], False)
        self.assertTrue(all(not item["checked"] for item in by_model["21"]["variables"]))
        efficiency = {item["name"]: item["checked"] for item in by_model["9"]["variables"]}
        self.assertEqual(efficiency["PF_1"], True)
        self.assertEqual(efficiency["Eff_1"], True)
        self.assertEqual(efficiency["Vout_1"], False)
        self.assertTrue(default_checked("Short Circuit Protection Test_2(M)", "Rec_Vout_1"))
        self.assertFalse(default_checked("Dynamic Load Setup(M)", "Dynamic_Load_1_1"))
        choices = choices_from_payload(report, {"groups": [
            {
                "key": group["key"],
                "input_voltage": group["inputVoltage"],
                "load_current": group["loadCurrent"],
                "checked": [variable["uid"] for variable in group["variables"]],
            }
            for group in model["groups"]
        ]})
        deck = assemble(report, choices)
        self.assertEqual(deck.unit_count, 10)
        self.assertEqual(deck.dates, ["2026/7/27~2026/9/30"])
        presentation = build_presentation(deck)
        self.assertEqual(len(presentation.slides), 26)
        _assert_shapes_inside(presentation, presentation.slides[0])
        summary_tables = [shape for shape in presentation.slides[0].shapes if shape.has_table]
        self.assertLess(summary_tables[1].height, summary_tables[2].height)
        summary = slide_text(presentation.slides[0])
        self.assertIn("Date", summary)
        self.assertIn("2026/7/27~2026/9/30", summary)
        self.assertIn("TBFHD1048401", summary)
        self.assertIn("Static Load Test", summary)
        self.assertIn(SAMPLE.stem, summary)
        self.assertNotIn(".xlsx", summary)
        self.assertNotIn("Samples", summary)
        first = presentation.slides[1]
        tables = [shape.table for shape in first.shapes if shape.has_table]
        self.assertEqual(
            [cell.text for cell in tables[0].rows[0].cells],
            ["VarName", "Upper Spec", "Lower Spec", "Mean", "Ca", "Cpk", "Cp"],
        )
        text = slide_text(first)
        self.assertNotIn(SAMPLE.name, text)
        self.assertNotIn(SAMPLE.stem, text)
        self.assertIn("SEQ 1 Static Load Test 90VAC", text)
        self.assertIn("3 / 1.25 A", text)
        self.assertIn("4.637", text)
        self.assertNotIn("4.637153", text)
        self.assertIn("50.4", text)
        self.assertNotIn("50.400002", text)
        self.assertNotIn("2026/7/27~2026/9/30", text)
        self.assertNotIn("TBFED1022848", text)
        self.assertIn("1F4E79", first._element.xml)
        self.assertIn("BFBFBF", first._element.xml)
        for index in range(1, len(presentation.slides)):
            slide = presentation.slides[index]
            self.assertNotIn("Samples", slide_text(slide))
            _assert_shapes_inside(presentation, slide)


class CliTests(unittest.TestCase):
    def test_missing_file_returns_1(self):
        self.assertEqual(main(["/tmp/does-not-exist-power-report.xlsx"]), 1)

    def test_no_path_uses_the_file_dialog_hook(self):
        calls = []

        def choose():
            calls.append("choose")
            return None

        self.assertEqual(main([], choose_workbook=choose), 1)
        self.assertEqual(calls, ["choose"])


if __name__ == "__main__":
    unittest.main()
