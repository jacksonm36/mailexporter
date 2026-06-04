@echo off
cd /d "%~dp0"
echo Mail Exporter - Defender exclusion + rebuild
echo.
echo This runs PowerShell with ExecutionPolicy Bypass for this script only.
echo For Defender exclusions, right-click this file and choose "Run as administrator".
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0allow_and_build.ps1"
pause
