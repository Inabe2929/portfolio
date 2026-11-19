@echo off
chcp 65001 >nul
cd /d "%~dp0"
py csv_report.py sample_input
pause
