@echo off
rem Installs MCTM Tracker for this Windows user: shortcuts on the Desktop and Start menu,
rem and an entry in Settings ^> Apps. No admin rights needed.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0installer\install.ps1" %*
