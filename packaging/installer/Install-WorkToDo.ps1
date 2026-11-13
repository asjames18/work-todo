<#
  Work To-Do installer (no admin needed). Run it with Install-WorkToDo.cmd (double-click).
  Installs to %LOCALAPPDATA%\WorkToDo:
    app\    the code (updates replace only this)
    data\   your tasks, settings and backups (never touched by updates or a reinstall)
    voice\  optional local Whisper for voice input (private Python environment + speech model)
  If Python 3.10+ isn't on the PC it offers to install the official python.org Python (per-user) first.
  Running it again is safe: it repairs or upgrades the app and skips anything already in place.

  Test / sandbox example (touches nothing real; -PythonDir unpacks a throwaway Python into that folder):
    Install-WorkToDo.cmd -InstallRoot "%TEMP%\WTDTest" -ShortcutName "Work To-Do TEST" -DesktopDir "%TEMP%\WTDTest\desk"
      -StartMenuDir "%TEMP%\WTDTest\start" -StartupDir "%TEMP%\WTDTest\startup" -NoHotkey -NoStartup -NoLaunch -Quiet
      [-WithVoice] [-ForcePythonInstall -PythonDir "%TEMP%\WTDTest\py"]
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
  [switch]$Quiet,
  [switch]$WithVoice,
  [switch]$NoVoice,
  [string]$VoiceModel = "small.en",
  [switch]$InstallPython,
  [switch]$ForcePythonInstall,
  [string]$PythonDir = ""
)
$ErrorActionPreference = "Stop"
$Here = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $Here "SetupCommon.ps1")
Initialize-Net
$PkgRoot = Split-Path -Parent $Here
$Src = Join-Path $PkgRoot "app"
$InstallRoot = [IO.Path]::GetFullPath($InstallRoot)
$AppDir = Join-Path $InstallRoot "app"
$DataDir = Join-Path $InstallRoot "data"
$VoiceDir = Join-Path $InstallRoot "voice"
$T0 = Get-Date

function Finish($code) {
  $secs = [int]((Get-Date) - $T0).TotalSeconds
  Say ("  (took {0} min {1} s)" -f [math]::Floor($secs / 60), ($secs % 60))
  if (-not $Quiet) { Write-Host ""; Read-Host "Press Enter to close" | Out-Null }
  exit $code
}

Say ""
Say "  Work To-Do setup" Cyan
Say "  ----------------"
if (-not (Test-Path -LiteralPath (Join-Path $Src "todo.pyw"))) {
  Say "  Can't find the app files next to this installer." Red
  Say "  Right-click the downloaded zip, choose 'Extract All...', then run Install-WorkToDo.cmd from the extracted folder."
  Finish 1
}
$Version = "?"
$m = Select-String -LiteralPath (Join-Path $Src "wtd_paths.py") -Pattern 'VERSION = "([^"]+)"' | Select-Object -First 1
if ($m) { $Version = $m.Matches[0].Groups[1].Value }
Say "  Version $Version  ->  $InstallRoot"
Test-PathLength $InstallRoot
if (-not (Assert-FreeSpace $InstallRoot 60 "Work To-Do")) { Finish 1 }

# ---- 1. voice input: decide up front so all the questions come first ----
$voiceReuse = $null
$lt = Join-Path $env:LOCALAPPDATA "LiveTranscriber"
if ((Test-Path -LiteralPath (Join-Path $lt "venv\Scripts\python.exe")) -and (Get-ChildItem -LiteralPath (Join-Path $lt "models") -Filter "models--*" -Directory -ErrorAction SilentlyContinue)) { $voiceReuse = $lt }
$ownVoice = Test-Path -LiteralPath (Join-Path $VoiceDir "venv\Scripts\python.exe")
$wantVoice = $false
if ($NoVoice) { $wantVoice = $false }
elseif ($WithVoice -or $ownVoice) { $wantVoice = $true }
elseif ($voiceReuse) { Say "  Voice input: will use Live Transcriber's local Whisper ($voiceReuse)" Green }
else {
  Say ""
  Say "  Optional: voice input. Speak a task and it's typed for you, on this PC (nothing is sent online)."
  Say "  It downloads a private speech engine (Whisper): about 550 MB, roughly 1 GB on disk, 3-10 minutes."
  $wantVoice = Ask-YesNo "Set up voice input now? (you can also run this installer again later)" $true -Quiet:$Quiet
  if ($Quiet -and -not $WithVoice) { $wantVoice = $false }
}

# ---- 2. Python ----
$minPy = [version]"3.10"; if ($wantVoice) { $minPy = [version]"3.12" }
$py = Resolve-Python -Min $minPy -NeedTk -NeedVenv:$wantVoice -Force:$ForcePythonInstall -SandboxDir $PythonDir -AutoYes:($InstallPython -or $ForcePythonInstall) -Quiet:$Quiet -AppName "Work To-Do"
if (-not $py) {
  Say "  Work To-Do needs Python. Nothing was installed." Yellow
  Finish 2
}
$PyExe = $py.Path
$PyW = Join-Path (Split-Path $PyExe) "pythonw.exe"
if (-not (Test-Path -LiteralPath $PyW)) { $PyW = $PyExe }

# ---- 3. stop a running copy (upgrade) and copy the app (code only; data\ is never touched) ----
Stop-AppProcesses $AppDir
New-Item -ItemType Directory -Force -Path $AppDir, $DataDir | Out-Null
Copy-Item -Path (Join-Path $Src "*") -Destination $AppDir -Recurse -Force
foreach ($f in @("Uninstall-WorkToDo.ps1", "SetupCommon.ps1")) { Copy-Item -LiteralPath (Join-Path $Here $f) -Destination (Join-Path $InstallRoot $f) -Force }
Get-ChildItem -LiteralPath $InstallRoot -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.FullName -notlike "$DataDir*" -and $_.FullName -notlike "$VoiceDir*" } | Unblock-File -ErrorAction SilentlyContinue
Say "  App copied to $AppDir" Green

# ---- 4. voice input (private venv + Whisper model) ----
$voicePy = $null; $voiceModels = $null
if ($wantVoice) {
  Say ""
  Say "  Setting up voice input..." Cyan
  if ((Assert-FreeSpace $InstallRoot 1500 "voice input") -and (Ensure-VCRuntime -AutoYes:$InstallPython -Quiet:$Quiet)) {
    $vpy = Ensure-Venv $PyExe (Join-Path $VoiceDir "venv") (Join-Path $AppDir "requirements-voice.txt") "faster_whisper, ctranslate2, av, numpy" "voice input"
    if ($vpy -and (Ensure-WhisperModel $vpy (Join-Path $VoiceDir "models") $VoiceModel)) {
      $voicePy = "..\voice\venv\Scripts\python.exe"; $voiceModels = "..\voice\models"
      Say "  Voice input is ready" Green
    }
  }
  if (-not $voicePy) { Say "  Voice input isn't set up (see above). Work To-Do works without it; run the installer again to retry." Yellow }
}

$info = [ordered]@{ data_dir = "..\data"; version = $Version; installed = (Get-Date -Format s); python = $PyW; shortcut = $ShortcutName }
if ($voicePy) { $info.voice_python = $voicePy; $info.voice_models = $voiceModels }
Write-Utf8NoBom (Join-Path $AppDir "installed.json") ($info | ConvertTo-Json)
Say "  Your data lives in $DataDir (updates never touch it)"
if (-not $py.Tk) { Say "  Note: this Python has no tkinter, so the Ctrl+Alt+N quick-add box won't open. Everything else works." Yellow }

# ---- 5. shortcuts ----
$Ico = Join-Path $AppDir "todo.ico"
$appArgs = '"' + (Join-Path $AppDir "todo.pyw") + '"'
New-Item -ItemType Directory -Force -Path $DesktopDir, $StartMenuDir | Out-Null
function Get-LinkPath($dir, $name) {
  # Never overwrite someone else's shortcut with the same name (e.g. a hand-made one that runs another copy).
  $p = Join-Path $dir "$name.lnk"
  if (Test-Path -LiteralPath $p) {
    try { $l = (New-Object -ComObject WScript.Shell).CreateShortcut($p); if (($l.Arguments + $l.TargetPath) -notlike "*$InstallRoot*") { $p = Join-Path $dir "$name (installed).lnk" } } catch { }
  }
  return $p
}
$desk = Get-LinkPath $DesktopDir $ShortcutName
$start = Get-LinkPath $StartMenuDir $ShortcutName
$hot = $null
if (-not $NoHotkey) {
  $taken = Find-HotkeyClash $Hotkey (Get-ShortcutDirs $DesktopDir $StartMenuDir) @($desk, $start)
  if ($taken) { Say "  Note: '$(Split-Path $taken -Leaf)' already uses $Hotkey, so the new shortcut has no hotkey." Yellow } else { $hot = $Hotkey }
}
New-AppShortcut $desk $PyW $appArgs $AppDir $Ico $hot "Work To-Do: projects, tasks and reminders"
New-AppShortcut $start $PyW $appArgs $AppDir $Ico $null "Work To-Do: projects, tasks and reminders"
$unArgs = '-NoProfile -ExecutionPolicy Bypass -File "' + (Join-Path $InstallRoot "Uninstall-WorkToDo.ps1") + '" -InstallRoot "' + $InstallRoot + '" -ShortcutName "' + $ShortcutName + '" -DesktopDir "' + $DesktopDir + '" -StartMenuDir "' + $StartMenuDir + '" -StartupDir "' + $StartupDir + '"'
New-AppShortcut (Join-Path $StartMenuDir "Uninstall $ShortcutName.lnk") "powershell.exe" $unArgs $InstallRoot $Ico $null "Uninstall Work To-Do"
$hotMsg = ""; if ($hot) { $hotMsg = " (hotkey $hot)" }
Say ("  Shortcuts: Desktop + Start menu" + $hotMsg) Green

# ---- 6. reminder helper at login (also powers Ctrl+Alt+N quick add) ----
$useStartup = $false
if ($StartAtLogin) { $useStartup = $true }
elseif (-not $NoStartup) {
  if ($Quiet) { $useStartup = $true }
  else { $useStartup = Ask-YesNo "Start the reminder helper when you sign in? It shows reminders and powers Ctrl+Alt+N quick add." $true }
}
$remArgs = '"' + (Join-Path $AppDir "reminder.pyw") + '"'
if ($useStartup) {
  New-Item -ItemType Directory -Force -Path $StartupDir | Out-Null
  New-AppShortcut (Join-Path $StartupDir "$ShortcutName Reminders.lnk") $PyW $remArgs $AppDir $Ico $null "Work To-Do reminders and quick add (Ctrl+Alt+N)"
  if (-not $NoLaunch) { Start-Process -FilePath $PyW -ArgumentList $remArgs -WorkingDirectory $AppDir }
  Say "  Reminder helper starts at sign-in" Green
} else { Say "  Reminder helper not set to start at sign-in (turn it on later in Settings)" }

# ---- 7. launch into the setup wizard ----
if (-not $NoLaunch) {
  Start-Process -FilePath $PyW -ArgumentList $appArgs -WorkingDirectory $AppDir
  Say "  Opening Work To-Do. The setup wizard walks you through the rest." Cyan
}
$vtxt = "off"; if ($voicePy) { $vtxt = "own" } elseif ($voiceReuse) { $vtxt = "live-transcriber" }
"installed $Version $(Get-Date -Format s) python=$PyExe voice=$vtxt" | Add-Content -LiteralPath (Join-Path $DataDir "install.log")
Say ""
Say "  Done!" Green
Finish 0
