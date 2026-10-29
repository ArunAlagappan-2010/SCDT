@echo off
cd /d "%~dp0"
title School Discipline Tracker - server

echo ============================================================
echo   School Discipline and Grooming Tracker
echo ============================================================
echo.

rem ---- 1. Find Python -----------------------------------------
set "PY="
where py >nul 2>nul
if not errorlevel 1 set "PY=py"
if defined PY goto :havepy
where python >nul 2>nul
if not errorlevel 1 set "PY=python"
:havepy
if not defined PY goto :nopython

echo [1/3] Using Python command: %PY%
%PY% --version
echo.

rem ---- 2. Install the required packages (first run only) ------
%PY% -c "import flask, openpyxl, PIL, pytesseract, requests, dotenv" >nul 2>nul
if not errorlevel 1 goto :haspkgs
echo [2/3] Installing required packages, this takes a minute...
%PY% -m pip install --upgrade pip
%PY% -m pip install -r requirements.txt
if errorlevel 1 goto :pipfailed
goto :pkgsdone
:haspkgs
echo [2/3] Required packages already installed.
:pkgsdone
echo.

rem ---- 3. Create the Excel database on first run only ---------
if exist "data\SchoolDisciplineSystem.xlsx" goto :hasdb
echo [3/3] Creating the Excel database for the first time...
%PY% build_excel_template.py
echo.
echo     NOTE: the database starts with 5 example students.
echo     Remove them on the "Master student list" page before real use.
goto :dbdone
:hasdb
echo [3/3] Database found - leaving your existing data untouched.
:dbdone
echo.

echo ============================================================
echo   Starting the server. KEEP THIS WINDOW OPEN.
echo.
echo   On this PC open:        http://localhost:5000
echo   On other PCs on the     http://THIS-PC-IP:5000
echo   same network open:      (run "ipconfig" to find the IP)
echo.
echo   If Windows asks about the firewall, click "Allow access".
echo   Press Ctrl+C in this window to stop the server.
echo ============================================================
echo.

%PY% app.py
echo.
echo The server has stopped.
pause
exit /b 0

:nopython
echo.
echo   PYTHON WAS NOT FOUND ON THIS PC.
echo.
echo   Install it from https://www.python.org/downloads/
echo   and tick "Add python.exe to PATH" in the installer,
echo   then run this file again.
echo.
pause
exit /b 1

:pipfailed
echo.
echo   Installing the packages failed - see the messages above.
echo   A common fix is to run this file again, or to run:
echo       %PY% -m pip install --user -r requirements.txt
echo.
pause
exit /b 1
