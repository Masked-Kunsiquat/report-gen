@echo off
cd /d "%~dp0"
uv run --with pandas --with openpyxl report_gui.py
