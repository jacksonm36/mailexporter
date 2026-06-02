# Add Windows Defender exclusion for Mail Exporter, then rebuild.
#
# Option A (easiest): double-click RUN_ALLOW_AND_BUILD.bat
#   - Bypasses PowerShell execution policy for this script only
#   - Right-click -> Run as administrator for Defender exclusions
#
# Option B (manual PowerShell):
#   Set-ExecutionPolicy -Scope Process Bypass
#   .\allow_and_build.ps1

$ErrorActionPreference = 'Stop'
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$DistPath = Join-Path $ProjectRoot 'dist'

Write-Host "Mail Exporter - Defender exclusion + rebuild" -ForegroundColor Cyan
Write-Host "Project: $ProjectRoot"

try {
    Add-MpPreference -ExclusionPath $ProjectRoot
    Add-MpPreference -ExclusionPath $DistPath
    Write-Host "Added Defender exclusions for project and dist folders." -ForegroundColor Green
} catch {
    Write-Host "Could not add Defender exclusion (need Administrator?):" -ForegroundColor Yellow
    Write-Host $_.Exception.Message
    Write-Host ""
    Write-Host "Do this manually:" -ForegroundColor Yellow
    Write-Host "  Windows Security -> Virus and threat protection -> Manage settings"
    Write-Host "  -> Exclusions -> Add an exclusion -> Folder -> $ProjectRoot"
    Read-Host "Press Enter after adding the exclusion (or to continue anyway)"
}

Set-Location $ProjectRoot
python -m pip install -r requirements.txt -q
python build_exe.py --onedir
python build_exe.py

Write-Host ""
Write-Host "Done. Try:" -ForegroundColor Green
Write-Host "  $DistPath\MailExporter_x32\MailExporter_x32.exe"
Write-Host "  $DistPath\MailExporter_x32.exe"
Write-Host ""
Write-Host "If still blocked: Windows Security -> Protection history -> Restore/Allow"
