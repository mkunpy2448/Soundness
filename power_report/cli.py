"""Command line entry for the workbook-to-PowerPoint tool."""

import argparse
import sys
from collections.abc import Callable
from pathlib import Path

from openpyxl.utils.exceptions import InvalidFileException

try:
    import tkinter as tk
    from tkinter import filedialog
except ImportError:  # A Python install without Tcl can still run when a path is given.
    tk = None
    filedialog = None

from power_report.parse import TemplateError, load_report
from power_report.server import serve


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
        workbook = picked.expanduser().resolve()
    else:
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
