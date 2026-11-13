<#  Work To-Do uninstaller. Removes the app, its voice-input files and its shortcuts. Your data (tasks, notes, settings,
    backups) is kept unless you answer Yes or pass -RemoveData. Python itself is left installed (other apps may use it).  #>
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
if (-not (Test-Path -LiteralPath (Join-Path $AppDir "todo.pyw")) -and -not (Test-Path -LiteralPath $DataDir)) {
  Write-Host "  Work To-Do isn't installed at $InstallRoot." -ForegroundColor Yellow
  if (-not $Quiet) { Read-Host "Press Enter to close" | Out-Null }; exit 1
}
if (-not $Quiet -and -not $RemoveData) {
  $a = Read-Host "  Also delete your tasks, notes and settings in $DataDir ? (make a backup first if unsure) [y/N]"
  $RemoveData = $a -match '^(y|yes)$'
}
$procs = Get-CimInstance Win32_Process -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "python*" -and $_.CommandLine -and $_.CommandLine -like "*$InstallRoot\*" }
foreach ($p in $procs) { try { Stop-Process -Id $p.ProcessId -Force; Write-Host "  Stopped pid $($p.ProcessId)" } catch { } }
Start-Sleep -Milliseconds 600
$sh = New-Object -ComObject WScript.Shell
foreach ($l in @((Join-Path $DesktopDir "$ShortcutName.lnk"), (Join-Path $DesktopDir "$ShortcutName (installed).lnk"), (Join-Path $StartMenuDir "$ShortcutName.lnk"),
                 (Join-Path $StartMenuDir "$ShortcutName (installed).lnk"), (Join-Path $StartMenuDir "Uninstall $ShortcutName.lnk"), (Join-Path $StartupDir "$ShortcutName Reminders.lnk"))) {
  if (-not (Test-Path -LiteralPath $l)) { continue }
  # only remove shortcuts that point into this install
  $mine = $true; try { $s = $sh.CreateShortcut($l); $mine = (($s.Arguments + $s.TargetPath) -like "*$InstallRoot*") } catch { }
  if ($mine) { Remove-Item -LiteralPath $l -Force; Write-Host "  Removed $l" }
}
foreach ($d in @($AppDir, (Join-Path $InstallRoot "voice"))) { if (Test-Path -LiteralPath $d) { Remove-Item -LiteralPath $d -Recurse -Force; Write-Host "  Removed $d" } }
if ($RemoveData) { if (Test-Path -LiteralPath $DataDir) { Remove-Item -LiteralPath $DataDir -Recurse -Force; Write-Host "  Removed your data" } }
else { Write-Host "  Kept your data in $DataDir" -ForegroundColor Green }
foreach ($f in @("SetupCommon.ps1", "Uninstall-WorkToDo.ps1")) { $p = Join-Path $InstallRoot $f; if (Test-Path -LiteralPath $p) { Remove-Item -LiteralPath $p -Force -ErrorAction SilentlyContinue } }
if ((Test-Path -LiteralPath $InstallRoot) -and -not (Get-ChildItem -LiteralPath $InstallRoot -Force -ErrorAction SilentlyContinue)) { Remove-Item -LiteralPath $InstallRoot -Force }
Write-Host "  Work To-Do was uninstalled." -ForegroundColor Green
if (-not $Quiet) { Read-Host "Press Enter to close" | Out-Null }
exit 0
