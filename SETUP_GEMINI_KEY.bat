@echo off
rem Adds the Gemini API key so Gemini reads register photos first and
rem Tesseract is only the backup. Asks for the key, checks it with Google,
rem saves it in .env and restarts the tracker.
rem   SETUP_GEMINI_KEY.bat -Remove    goes back to Tesseract only.
title MCTM Tracker - Gemini key
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0installer\set_gemini_key.ps1" %*
