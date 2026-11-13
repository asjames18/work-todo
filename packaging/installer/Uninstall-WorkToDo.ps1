<#  Work To-Do uninstaller. Removes the app and its shortcuts. Your data (tasks, notes, settings, backups) is kept
    unless you answer Yes or pass -RemoveData.  #>
param(
  [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA "WorkToDo"),
  [string]$ShortcutName = "Work To-Do",
  [string]$DesktopDir = [Environment]::GetFolderPath("Desktop"),
  [string]$StartMenuDir = (Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs"),
  [string]$StartupDir = [Environment]::GetFolderPath("Startup"),
  [switch]$RemoveData,
  [switch]$Quiet
)
$ErrorActionPreference = "Continue"
$AppDir = Join-Path $InstallRoot "app"
$DataDir = Join-Path $InstallRoot "data"
Write-Host ""; Write-Host "  Uninstall Work To-Do" -ForegroundColor Cyan
if (-not (Test-Path (Join-Path $AppDir "todo.pyw")) -and -not (Test-Path $DataDir)) {
  Write-Host "  Work To-Do isn't installed at $InstallRoot." -ForegroundColor Yellow
  if (-not $Quiet) { Read-Host "Press Enter to close" | Out-Null }; exit 1
}
if (-not $Quiet -and -not $RemoveData) {
  $a = Read-Host "  Also delete your tasks, notes and settings in $DataDir ? (make a backup first if unsure) [y/N]"
  $RemoveData = $a -match '^(y|yes)$'
}
$procs = Get-CimInstance Win32_Process -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "python*" -and $_.CommandLine -and $_.CommandLine -like "*$AppDir*" }
foreach ($p in $procs) { try { Stop-Process -Id $p.ProcessId -Force; Write-Host "  Stopped pid $($p.ProcessId)" } catch { } }
Start-Sleep -Milliseconds 600
foreach ($l in @((Join-Path $DesktopDir "$ShortcutName.lnk"), (Join-Path $StartMenuDir "$ShortcutName.lnk"),
                 (Join-Path $StartMenuDir "Uninstall $ShortcutName.lnk"), (Join-Path $StartupDir "$ShortcutName Reminders.lnk"))) {
  if (Test-Path $l) { Remove-Item $l -Force; Write-Host "  Removed $l" }
}
if (Test-Path $AppDir) { Remove-Item $AppDir -Recurse -Force; Write-Host "  Removed the app" }
if ($RemoveData) { if (Test-Path $DataDir) { Remove-Item $DataDir -Recurse -Force; Write-Host "  Removed your data" } }
else { Write-Host "  Kept your data in $DataDir" -ForegroundColor Green }
$self = Join-Path $InstallRoot "Uninstall-WorkToDo.ps1"
if (Test-Path $self) { Remove-Item $self -Force -ErrorAction SilentlyContinue }
if ((Test-Path $InstallRoot) -and -not (Get-ChildItem $InstallRoot -Force -ErrorAction SilentlyContinue)) { Remove-Item $InstallRoot -Force }
Write-Host "  Work To-Do was uninstalled." -ForegroundColor Green
if (-not $Quiet) { Read-Host "Press Enter to close" | Out-Null }
exit 0
