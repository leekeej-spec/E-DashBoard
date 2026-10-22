@echo off
rem E-Dashboard (read-only). Keep this window open while the dashboard runs.
rem Options can be added, e.g.  EDashboard.cmd -RefreshSeconds 30
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0EDashboard.ps1" %*
if errorlevel 1 pause
