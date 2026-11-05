"""Double-click this file, or run_report.bat on Windows, to build a report."""

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent


def ensure_packages() -> None:
    # Imported here, not at module load, so a missing package can be installed first.
    try:
        import openpyxl  # noqa: F401
        import pptx  # noqa: F401
    except ImportError:
        requirements = ROOT / "requirements.txt"
        print("正在安裝需要的套件，請稍候...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", str(requirements)])


def main() -> int:
    if str(ROOT) not in sys.path:
        sys.path.insert(0, str(ROOT))
    try:
        ensure_packages()
        # Imported after the repo root is on the path and packages are installed.
        from power_report.cli import main as run

        code = run()
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
    raise SystemExit(main())
