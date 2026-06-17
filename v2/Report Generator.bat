@echo off
cd /d "%~dp0"
if exist uv.exe (
    uv.exe run --with pandas --with openpyxl report_gui.py
) else (
    uv run --with pandas --with openpyxl report_gui.py
)
