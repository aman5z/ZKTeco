@echo off
title MDB Employee Exporter
echo.
echo ================================================
echo   Exporting employees from .mdb database...
echo ================================================
echo.

pip install pyodbc openpyxl pandas --quiet 2>nul

python "%~dp0EXPORT_EMPLOYEES.py" %1

if %ERRORLEVEL% EQU 0 (
    echo.
    echo Opening exported files...
    start "" "%~dp0employees_export.xlsx"
)
