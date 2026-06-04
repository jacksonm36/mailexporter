@echo off
REM Launches Mail Exporter from AppData build (outside quarantined Documents\dist)
set "APP=%LOCALAPPDATA%\MailExporter\MailExporter_x32\MailExporter_x32.exe"
if not exist "%APP%" (
  echo Mail Exporter not built yet.
  echo Run: python build_exe.py --onedir --dist-dir %%LOCALAPPDATA%%\MailExporter
  echo Or double-click RUN_FROM_SOURCE.bat to use Python directly.
  pause
  exit /b 1
)
start "" "%APP%"
