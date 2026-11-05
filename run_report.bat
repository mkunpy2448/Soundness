@echo off
cd /d "%~dp0"
where py >nul 2>&1
if %errorlevel%==0 (
  py -3 "%~dp0run_report.py" %*
) else (
  python "%~dp0run_report.py" %*
)
if errorlevel 1 pause
