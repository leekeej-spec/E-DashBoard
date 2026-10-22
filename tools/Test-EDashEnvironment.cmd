@echo off
rem E-Dashboard environment check (read-only). Writes a report to the Desktop.
rem Extra arguments are passed through, e.g.  Test-EDashEnvironment.cmd -ListNames EasDashboard
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Test-EDashEnvironment.ps1" %*
echo.
pause
