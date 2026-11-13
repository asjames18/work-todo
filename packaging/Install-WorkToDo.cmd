@echo off
setlocal
rem Work To-Do installer: double-click this file (after extracting the zip).
if not exist "%~dp0installer\Install-WorkToDo.ps1" (
  echo Please extract the zip first: right-click it, choose "Extract All...", then run Install-WorkToDo.cmd from the extracted folder.
  pause
  exit /b 1
)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0installer\Install-WorkToDo.ps1" %*
exit /b %errorlevel%
