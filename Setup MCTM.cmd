@echo off
rem MCTM Tracker - one-time setup. Double-click once on the server PC.
rem Installs Python (if missing), the Python packages, the Tesseract reading
rem program, an optional Gemini key, the database and the Desktop/Start menu
rem shortcuts. Safe to run again: anything already installed is skipped.
title MCTM Tracker setup
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0installer\setup.ps1" %*
