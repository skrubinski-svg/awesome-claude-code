@echo off
rem Double-click launcher: runs run_fsc_report.ps1 and keeps the window open.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0run_fsc_report.ps1" %*
pause
