@echo off
rem Removes MCTM Tracker. Your data is backed up to Documents first unless you say no,
rem and anything removed goes to the Recycle Bin.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0installer\uninstall.ps1" %*
