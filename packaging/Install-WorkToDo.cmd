@echo off
setlocal
rem Work To-Do installer: double-click this file.
rem If it was opened straight from inside the zip (Windows only pulls out this one file),
rem it finds the downloaded zip, extracts it for you and carries on.
set "WTD_PS=%~dp0installer\Install-WorkToDo.ps1"
if exist "%WTD_PS%" goto run
set "WTD_HERE=%~dp0"
set "WTD_SELF=%~f0"
set "WTD_OUT=%TEMP%\wtd-setup-path.txt"
if exist "%WTD_OUT%" del "%WTD_OUT%" >nul 2>&1
echo.
echo   Work To-Do setup: unpacking the download first...
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$s=[IO.File]::ReadAllText($env:WTD_SELF); $i=$s.LastIndexOf('#'+'PS-BOOT'); Invoke-Expression $s.Substring($i)"
if not exist "%WTD_OUT%" exit /b 1
set /p WTD_PS=<"%WTD_OUT%"
del "%WTD_OUT%" >nul 2>&1
if not exist "%WTD_PS%" exit /b 1
:run
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%WTD_PS%" %*
exit /b %errorlevel%

#PS-BOOT
$ErrorActionPreference = 'Stop'
function Show-Msg($text) {
  try { Add-Type -AssemblyName System.Windows.Forms; [void][System.Windows.Forms.MessageBox]::Show($text, 'Work To-Do setup', 'OK', 'Information') }
  catch { Write-Host $text; Read-Host 'Press Enter to close' | Out-Null }
}
try {
  $here = $env:WTD_HERE.TrimEnd('\')
  $name = Split-Path $here -Leaf
  if ($name -notlike 'WorkToDo-Setup-*') { $name = 'WorkToDo-Setup-*' }
  $dirs = @()
  if ($env:WTD_SEARCH) { $dirs += $env:WTD_SEARCH }
  try { $dirs += (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path } catch { }
  $dirs += (Join-Path $env:USERPROFILE 'Downloads'), [Environment]::GetFolderPath('Desktop'), [Environment]::GetFolderPath('MyDocuments')
  $dirs = $dirs | Where-Object { $_ -and (Test-Path $_) } | Select-Object -Unique
  $zip = Get-ChildItem -Path $dirs -Filter ($name + '.zip') -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
  if (-not $zip) {
    Show-Msg ("Windows opened this installer from inside the zip, so the rest of the app files aren't here.`r`n`r`nRight-click the downloaded WorkToDo-Setup zip, choose 'Extract All...', then double-click Install-WorkToDo.cmd in the folder that opens.")
    exit 1
  }
  $dest = Join-Path $env:TEMP ($zip.BaseName + '-unpacked')
  if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
  Expand-Archive -LiteralPath $zip.FullName -DestinationPath $dest -Force
  Get-ChildItem $dest -Recurse -File | Unblock-File -ErrorAction SilentlyContinue
  $ps = Get-ChildItem $dest -Recurse -Filter 'Install-WorkToDo.ps1' -File | Where-Object { $_.Directory.Name -eq 'installer' } | Select-Object -First 1
  if (-not $ps) { Show-Msg ("The download looks incomplete. Please download WorkToDo-Setup again and run Install-WorkToDo.cmd."); exit 1 }
  Write-Host ('  Found ' + $zip.Name + ' and unpacked it.')
  [IO.File]::WriteAllText($env:WTD_OUT, $ps.FullName)
  exit 0
} catch {
  Show-Msg ("Couldn't unpack the download automatically:`r`n" + $_.Exception.Message + "`r`n`r`nRight-click the zip, choose 'Extract All...', then run Install-WorkToDo.cmd from the extracted folder.")
  exit 1
}
