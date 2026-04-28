@echo off
title ZKTeco Attendance Dashboard  v2.2
color 0A
setlocal enabledelayedexpansion

pushd "%~dp0" 2>nul
if errorlevel 1 (echo  [ERROR] Cannot access: %~dp0 & pause & exit /b 1)

python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python not found.
    echo  Download from https://www.python.org/downloads/
    echo  IMPORTANT: check "Add Python to PATH" during install.
    popd & pause & exit /b 1
)

:MENU
cls
echo.
echo  ============================================================
echo   ZKTeco Attendance Dashboard  v2.2
echo  ============================================================
echo.
echo    [1]  Open Dashboard   (start server + browser)
echo    [2]  Sync Database    (full: employees + MDB + devices)
echo    [3]  Sync Devices Only
echo    [4]  Today's Absent Report  (Excel)
echo    [5]  History Absent Report  (Excel, from MDB)
echo    [6]  Export Employees from MDB
echo    [7]  Scan MDB         (diagnostics)
echo    [8]  Backup Database
echo    [9]  First-Time Setup (install packages)
echo    [0]  Exit
echo.
set /p CHOICE="  Choice: "

if "%CHOICE%"=="1" goto DASHBOARD
if "%CHOICE%"=="2" goto SYNC_FULL
if "%CHOICE%"=="3" goto SYNC_DEVICES
if "%CHOICE%"=="4" goto TODAY
if "%CHOICE%"=="5" goto HISTORY
if "%CHOICE%"=="6" goto EXPORT_EMP
if "%CHOICE%"=="7" goto SCAN_MDB
if "%CHOICE%"=="8" goto BACKUP
if "%CHOICE%"=="9" goto SETUP
if "%CHOICE%"=="0" goto END
echo  Invalid choice. & timeout /t 1 /nobreak >nul & goto MENU

:: ─────────────────────────────────────────────────────────────────────────────
:DASHBOARD
cls
echo.
echo  Starting server...
echo  URL : http://127.0.0.1:5000/d
echo  Keep this window open. Ctrl+C to stop.
echo  ============================================================
start /b cmd /c "timeout /t 6 /nobreak >nul 2>&1 & start http://127.0.0.1:5000/d"
python server.py
popd & pause & goto MENU

:: ─────────────────────────────────────────────────────────────────────────────
:SYNC_FULL
cls & echo. & echo  Full sync (employees + MDB history + devices)...
echo.
python sync_db.py
goto AFTER_ACTION

:SYNC_DEVICES
cls & echo. & echo  Device-only sync...
echo.
python sync_db.py devices-only
goto AFTER_ACTION

:: ─────────────────────────────────────────────────────────────────────────────
:TODAY
cls & echo. & echo  Today's Absent Report...
echo.
python mdb_tools.py today
goto AFTER_ACTION

:HISTORY
cls & echo.
set MDB_FOUND=
for %%f in ("*.mdb" "*.accdb") do if not defined MDB_FOUND set MDB_FOUND=%%f
if defined MDB_FOUND (
    echo  MDB detected: %MDB_FOUND%
) else (
    set /p MDB_FOUND="  Path to .mdb file: "
)
set /p DATE_FROM="  From (DD/MM/YYYY): "
set /p DATE_TO="    To (DD/MM/YYYY): "
echo.
python mdb_tools.py history "%MDB_FOUND%" %DATE_FROM% %DATE_TO%
goto AFTER_ACTION

:: ─────────────────────────────────────────────────────────────────────────────
:EXPORT_EMP
cls & echo. & echo  Exporting employees from MDB...
echo.
python mdb_tools.py export
goto AFTER_ACTION

:SCAN_MDB
cls & echo.
set MDB_FOUND=
for %%f in ("*.mdb" "*.accdb") do if not defined MDB_FOUND set MDB_FOUND=%%f
if defined MDB_FOUND (
    echo  Scanning: %MDB_FOUND%
    python mdb_tools.py scan "%MDB_FOUND%"
) else (
    set /p MDB_PATH="  Path to .mdb file: "
    python mdb_tools.py scan "%MDB_PATH%"
)
goto AFTER_ACTION

:BACKUP
cls & echo. & echo  Backing up attendance.db...
python sync_db.py backup
goto AFTER_ACTION

:: ─────────────────────────────────────────────────────────────────────────────
:SETUP
cls
echo.
echo  [1/4] Checking Python...
for /f "tokens=*" %%v in ('python --version 2^>^&1') do echo  OK: %%v
echo.
echo  [2/4] Installing packages...
python -m pip install --upgrade pip --quiet
if exist "requirements.txt" (
    python -m pip install -r requirements.txt --upgrade --quiet
) else (
    python -m pip install flask pandas openpyxl pyodbc pyzk Werkzeug --upgrade --quiet
)
echo  OK: packages installed.
echo.
echo  [3/4] Exporting employees...
set MDB_FOUND=
for %%f in ("*.mdb" "*.accdb") do if not defined MDB_FOUND set MDB_FOUND=%%f
if defined MDB_FOUND (
    python mdb_tools.py export
    echo  OK: employees exported.
) else (
    echo  SKIP: no .mdb found. Place your database here and re-run Setup.
)
echo.
echo  [4/4] Device sync...
python sync_db.py devices-only
echo.
echo  Setup complete! Run option [1] to open the dashboard.
goto AFTER_ACTION

:: ─────────────────────────────────────────────────────────────────────────────
:AFTER_ACTION
echo.
echo  ────────────────────────────────────────
pause
goto MENU

:END
popd
endlocal
