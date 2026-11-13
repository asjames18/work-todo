<#
  Work To-Do installer (no admin needed). Run it with Install-WorkToDo.cmd (double-click).
  Installs to %LOCALAPPDATA%\WorkToDo:  app\ = code (updates replace only this), data\ = your tasks and settings.
  Test / sandbox example (touches nothing real):
    Install-WorkToDo.cmd -InstallRoot "%TEMP%\WTDTest" -ShortcutName "Work To-Do TEST" -DesktopDir "%TEMP%\WTDTest\desk" -StartMenuDir "%TEMP%\WTDTest\start" -NoHotkey -NoStartup -NoLaunch -Quiet
#>
param(
  [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA "WorkToDo"),
  [string]$ShortcutName = "Work To-Do",
  [string]$DesktopDir = [Environment]::GetFolderPath("Desktop"),
  [string]$StartMenuDir = (Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs"),
  [string]$StartupDir = [Environment]::GetFolderPath("Startup"),
  [string]$Hotkey = "Ctrl+Alt+T",
  [switch]$NoHotkey,
  [switch]$NoStartup,
  [switch]$StartAtLogin,
  [switch]$NoLaunch,
  [switch]$Quiet
)
$ErrorActionPreference = "Stop"
$Here = Split-Path -Parent $MyInvocation.MyCommand.Path
$PkgRoot = Split-Path -Parent $Here
$Src = Join-Path $PkgRoot "app"
$AppDir = Join-Path $InstallRoot "app"
$DataDir = Join-Path $InstallRoot "data"
$PyUrl = "https://www.python.org/downloads/windows/"

function Say($msg, $color) { if ($color) { Write-Host $msg -ForegroundColor $color } else { Write-Host $msg } }
function Finish($code) { if (-not $Quiet) { Write-Host ""; Read-Host "Press Enter to close" | Out-Null }; exit $code }

Say ""
Say "  Work To-Do setup" Cyan
Say "  ----------------"
if (-not (Test-Path (Join-Path $Src "todo.pyw"))) {
  Say "  Can't find the app files next to this installer." Red
  Say "  Right-click the downloaded zip, choose 'Extract All...', then run Install-WorkToDo.cmd from the extracted folder."
  Finish 1
}
$Version = "?"
$m = Select-String -Path (Join-Path $Src "wtd_paths.py") -Pattern 'VERSION = "([^"]+)"' | Select-Object -First 1
if ($m) { $Version = $m.Matches[0].Groups[1].Value }
Say "  Version $Version"

# ---- 1. Python 3.10+ ----
function Test-Py($exe, [string[]]$pre) {
  try {
    $args2 = @($pre) + @("-c", "import sys;print(sys.executable);print('%d.%d' % sys.version_info[:2])")
    $out = & $exe @args2 2>$null
    if ($LASTEXITCODE -ne 0 -or -not $out) { return $null }
    $lines = @($out)
    if ($lines.Count -lt 2) { return $null }
    $path = $lines[0].Trim(); $ver = [version]($lines[1].Trim())
    if ($path -like "*WindowsApps*") { return $null }
    if ($ver -ge [version]"3.10") { return @{ Path = $path; Version = $ver } }
  } catch { }
  return $null
}
$py = $null
foreach ($c in @(@("py", @("-3")), @("python", @()), @("python3", @()))) {
  if (-not $py -and (Get-Command $c[0] -ErrorAction SilentlyContinue)) { $py = Test-Py $c[0] $c[1] }
}
if (-not $py) {
  $globs = @("$env:LOCALAPPDATA\Programs\Python\Python3*\python.exe", "$env:ProgramFiles\Python3*\python.exe", "C:\Python3*\python.exe")
  foreach ($g in $globs) {
    foreach ($f in (Get-ChildItem $g -ErrorAction SilentlyContinue | Sort-Object FullName -Descending)) {
      if (-not $py) { $py = Test-Py $f.FullName @() }
    }
  }
}
if (-not $py) {
  Say ""
  Say "  Python 3.10 or newer is needed and wasn't found." Yellow
  Say "  Opening the python.org download page. Install Python (tick 'Add python.exe to PATH'),"
  Say "  then run Install-WorkToDo.cmd again. Nothing was installed."
  if (-not $Quiet) { Start-Process $PyUrl }
  Finish 2
}
$PyExe = $py.Path
$PyW = Join-Path (Split-Path $PyExe) "pythonw.exe"
if (-not (Test-Path $PyW)) { $PyW = $PyExe }
Say "  Python $($py.Version) found: $PyExe" Green
& $PyExe -c "import tkinter" 2>$null
if ($LASTEXITCODE -ne 0) { Say "  Note: tkinter is missing, so the Ctrl+Alt+N quick-add box won't open. Everything else works." Yellow }

# ---- 2. stop a running copy (upgrade) ----
$running = Get-CimInstance Win32_Process -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "python*" -and $_.CommandLine -and $_.CommandLine -like "*$AppDir*" }
foreach ($p in $running) { try { Stop-Process -Id $p.ProcessId -Force -ErrorAction Stop; Say "  Stopped the running copy (pid $($p.ProcessId))" } catch { } }

# ---- 3. copy the app (code only; data\ is never touched) ----
New-Item -ItemType Directory -Force -Path $AppDir, $DataDir | Out-Null
Copy-Item -Path (Join-Path $Src "*") -Destination $AppDir -Recurse -Force
Copy-Item -Path (Join-Path $Here "Uninstall-WorkToDo.ps1") -Destination (Join-Path $InstallRoot "Uninstall-WorkToDo.ps1") -Force
Get-ChildItem $InstallRoot -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.FullName -notlike "$DataDir*" } | Unblock-File -ErrorAction SilentlyContinue
$info = [ordered]@{ data_dir = "..\data"; version = $Version; installed = (Get-Date -Format s); python = $PyW; shortcut = $ShortcutName }
($info | ConvertTo-Json) | Set-Content -Path (Join-Path $AppDir "installed.json") -Encoding ASCII
Say "  Installed to $InstallRoot" Green
Say "  Your data lives in $DataDir (updates never touch it)"

# ---- 4. shortcuts ----
$sh = New-Object -ComObject WScript.Shell
$Ico = Join-Path $AppDir "todo.ico"
function New-Lnk($path, $target, $arguments, $hot, $desc) {
  $l = $sh.CreateShortcut($path)
  $l.TargetPath = $target; $l.Arguments = $arguments; $l.WorkingDirectory = $AppDir
  $l.IconLocation = "$Ico,0"; $l.Description = $desc
  if ($hot) { $l.Hotkey = $hot }
  $l.Save()
}
$appArgs = '"' + (Join-Path $AppDir "todo.pyw") + '"'
$hot = $null
if (-not $NoHotkey) {
  $hot = $Hotkey
  $taken = Get-ChildItem $DesktopDir -Filter *.lnk -ErrorAction SilentlyContinue | Where-Object { $_.BaseName -ne $ShortcutName } |
    Where-Object { try { $sh.CreateShortcut($_.FullName).Hotkey -eq $Hotkey } catch { $false } } | Select-Object -First 1
  if ($taken) { Say "  Note: '$($taken.Name)' already uses $Hotkey, so the new shortcut has no hotkey." Yellow; $hot = $null }
}
New-Item -ItemType Directory -Force -Path $DesktopDir, $StartMenuDir | Out-Null
$desk = Join-Path $DesktopDir "$ShortcutName.lnk"
New-Lnk $desk $PyW $appArgs $hot "Work To-Do: projects, tasks and reminders"
New-Lnk (Join-Path $StartMenuDir "$ShortcutName.lnk") $PyW $appArgs $null "Work To-Do: projects, tasks and reminders"
$un = $sh.CreateShortcut((Join-Path $StartMenuDir "Uninstall $ShortcutName.lnk"))
$un.TargetPath = "powershell.exe"
$un.Arguments = '-NoProfile -ExecutionPolicy Bypass -File "' + (Join-Path $InstallRoot "Uninstall-WorkToDo.ps1") + '" -InstallRoot "' + $InstallRoot + '" -ShortcutName "' + $ShortcutName + '" -DesktopDir "' + $DesktopDir + '" -StartMenuDir "' + $StartMenuDir + '" -StartupDir "' + $StartupDir + '"'
$un.IconLocation = "$Ico,0"; $un.Save()
$hotMsg = ""; if ($hot) { $hotMsg = " (hotkey $hot)" }
Say ("  Shortcuts: Desktop + Start menu" + $hotMsg) Green

# ---- 5. reminder helper at login (also powers Ctrl+Alt+N quick add) ----
$useStartup = $false
if ($StartAtLogin) { $useStartup = $true }
elseif (-not $NoStartup) {
  if ($Quiet) { $useStartup = $true }
  else {
    $a = Read-Host "  Start the reminder helper when you sign in? It shows reminders and powers Ctrl+Alt+N quick add. [Y/n]"
    $useStartup = -not ($a -match '^(n|no)$')
  }
}
$remArgs = '"' + (Join-Path $AppDir "reminder.pyw") + '"'
if ($useStartup) {
  New-Item -ItemType Directory -Force -Path $StartupDir | Out-Null
  New-Lnk (Join-Path $StartupDir "$ShortcutName Reminders.lnk") $PyW $remArgs $null "Work To-Do reminders and quick add (Ctrl+Alt+N)"
  if (-not $NoLaunch) { Start-Process -FilePath $PyW -ArgumentList $remArgs -WorkingDirectory $AppDir }
  Say "  Reminder helper starts at sign-in (and is running now)" Green
} else { Say "  Reminder helper not set to start at sign-in (turn it on later in Settings)" }

# ---- 6. launch into the setup wizard ----
if (-not $NoLaunch) {
  Start-Process -FilePath $PyW -ArgumentList $appArgs -WorkingDirectory $AppDir
  Say "  Opening Work To-Do. The setup wizard walks you through the rest." Cyan
}
"installed $Version $(Get-Date -Format s) python=$PyExe" | Add-Content -Path (Join-Path $DataDir "install.log")
Say ""
Say "  Done!" Green
Finish 0
