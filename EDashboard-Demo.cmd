@echo off
rem E-Dashboard with built-in sample data (no SharePoint connection needed).
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0EDashboard.ps1" -Demo %*
if errorlevel 1 pause
