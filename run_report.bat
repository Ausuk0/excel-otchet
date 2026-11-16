@echo off
chcp 65001 >nul
cd /d "%~dp0"
python build_report.py %*
echo.
pause
