"""Command line entry for the workbook-to-PowerPoint tool."""

import argparse
import sys
from pathlib import Path

from openpyxl.utils.exceptions import InvalidFileException

from power_report.parse import TemplateError, load_report
from power_report.server import serve


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="將電源測試 Excel 轉成 PPT。先開啟勾選頁，再產生報告。")
    parser.add_argument("workbook", type=Path, help="電源測試 Excel 檔")
    args = parser.parse_args(argv)
    workbook = args.workbook.expanduser().resolve()
    if not workbook.is_file():
        print(f"找不到檔案：{workbook}", file=sys.stderr)
        return 1
    try:
        report = load_report(workbook)
    except TemplateError as exc:
        print(exc, file=sys.stderr)
        return 1
    except InvalidFileException:
        print(f"無法讀取 Excel：{workbook.name}", file=sys.stderr)
        return 1
    return serve(report, workbook)
