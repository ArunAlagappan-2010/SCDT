@echo off
cd /d "%~dp0"
title SCDT - set up the Gemini API key

if exist ".env" goto :open
copy ".env.example" ".env" >nul
if errorlevel 1 goto :copyfailed
echo Created a new .env file from the template.
echo.

:open
echo ============================================================
echo   Opening .env in Notepad.
echo.
echo   1. Find the line that says:   GEMINI_API_KEY=
echo   2. Paste your key straight after the "=" (no quotes,
echo      no spaces), so it looks like:  GEMINI_API_KEY=AIza...
echo   3. Save (Ctrl+S) and close Notepad.
echo.
echo   Get a free key at https://aistudio.google.com/apikey
echo.
echo   Then close the server window and run START_SCDT.bat again.
echo ============================================================
echo.
notepad ".env"
exit /b 0

:copyfailed
echo.
echo   Could not create .env - is .env.example missing from this folder?
echo.
pause
exit /b 1
