"""Render the assembled deck to a 16:9 PowerPoint file."""

import math
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

from power_report.assemble import ContentSlide, Deck
from power_report.numbers import format_number, quartiles

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
        _render_content(presentation.slides.add_slide(blank), deck, content)
    return presentation


def _render_summary(slide, deck: Deck) -> None:
    _white_background(slide)
    margin = 0.4
    top = 0.3
    width = SLIDE_W - 2 * margin
    date_text = ", ".join(deck.dates)
    meta_height = 0.78
    _add_table(
        slide,
        margin,
        top,
        width,
        meta_height,
        [["Date", date_text], ["N", str(deck.unit_count)]],
        font_pt=14,
        widths=[1.5, width - 1.5],
        header=False,
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


def _render_content(slide, deck: Deck, content: ContentSlide) -> None:
    _white_background(slide)
    margin_x = 0.4
    margin_y = 0.26
    width = SLIDE_W - 2 * margin_x
    header_h = 1.25
    _write_lines(
        slide,
        margin_x,
        margin_y,
        width,
        header_h,
        [
            (deck.source_name, 12, True),
            (content.title, 16, True),
            (f"Input voltage: {content.input_voltage}", 12, False),
            (f"Load current: {content.load_current}", 12, False),
        ],
    )
    count = len(content.variables)
    content_top = margin_y + header_h + 0.06
    content_bottom = SLIDE_H - margin_y
    available = content_bottom - content_top
    table_h = min(available * 0.55, 0.28 + count * 0.20)
    table_h = max(table_h, 0.7)
    if available - table_h < 1.6:
        table_h = max(0.6, available - 1.6)
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
    _add_table(
        slide,
        margin_x,
        content_top,
        width,
        table_h,
        rows,
        font_pt=_fit_font(len(rows), table_h, 11, 6),
        widths=[width * part for part in (0.20, 0.14, 0.14, 0.13, 0.13, 0.13, 0.13)],
    )
    chart_top = content_top + table_h + 0.08
    chart_h = content_bottom - chart_top
    _draw_chart_grid(slide, margin_x, chart_top, width, chart_h, content)


def _draw_chart_grid(slide, left, top, width, height, content: ContentSlide) -> None:
    count = len(content.variables)
    if count == 0 or height <= 0:
        return
    columns, rows = _grid_shape(count)
    cell_w = width / columns
    cell_h = height / rows
    for index, variable in enumerate(content.variables):
        column = index % columns
        row = index // columns
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


def _grid_shape(count: int) -> tuple[int, int]:
    if count <= 1:
        return 1, 1
    if count <= 2:
        return 2, 1
    if count <= 4:
        return 2, 2
    if count <= 6:
        return 3, 2
    if count <= 8:
        return 4, 2
    if count <= 12:
        return 4, 3
    if count <= 15:
        return 5, 3
    return 6, math.ceil(count / 6)


def _draw_boxplot(slide, left, top, width, height, title, values, lower, upper) -> None:
    pad = 0.04
    title_h = min(0.24, height * 0.22)
    _write_lines(
        slide,
        left + pad,
        top,
        max(0.2, width - 2 * pad),
        title_h,
        [(title, 8 if height >= 0.9 else 7, True)],
    )
    plot_left = left + pad
    plot_top = top + title_h
    plot_w = max(0.2, width - 2 * pad)
    plot_h = max(0.2, height - title_h - pad)
    lo, hi = _axis_limits(values, lower, upper)
    q1, median, q3 = quartiles(values)

    def y_of(value: float) -> float:
        fraction = (value - lo) / (hi - lo)
        return plot_top + (1 - fraction) * plot_h

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
    for limit in (lower, upper):
        if limit is None:
            continue
        _hline(slide, plot_left, y_of(limit), plot_left + plot_w, SPEC, 1.25)


def _axis_limits(values: list[float], lower: float | None, upper: float | None) -> tuple[float, float]:
    lo = min(values)
    hi = max(values)
    for limit in (lower, upper):
        if limit is None:
            continue
        lo = min(lo, limit)
        hi = max(hi, limit)
    if hi == lo:
        pad = abs(lo) * 0.05 if lo else 1.0
        return lo - pad, hi + pad
    pad = (hi - lo) * 0.08
    return lo - pad, hi + pad


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
    return format_number(value)


def _fit_font(row_count: int, height_in: float, max_pt: float, min_pt: float = 6) -> float:
    if row_count <= 0:
        return max_pt
    raw = (height_in / row_count) * 72 * 0.62
    return max(min_pt, min(max_pt, raw))


def _write_lines(slide, left, top, width, height, lines: list[tuple[str, float, bool]]):
    shape = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    frame = shape.text_frame
    frame.word_wrap = True
    frame.auto_size = None
    for index, (text, size, bold) in enumerate(lines):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.alignment = PP_ALIGN.LEFT
        paragraph.space_before = Pt(0)
        paragraph.space_after = Pt(0)
        run = paragraph.add_run()
        run.text = text
        _style_run(run, size, bold)
    return shape


def _add_table(slide, left, top, width, height, rows, font_pt, widths, header=True):
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
                bold=is_header,
                fill=HEADER_FILL if is_header else WHITE,
            )
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


def _style_run(run, size_pt: float, bold: bool) -> None:
    run.font.name = "Calibri"
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = False
    run.font.color.rgb = INK
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
