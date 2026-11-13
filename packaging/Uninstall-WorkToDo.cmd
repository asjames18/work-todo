@echo off
setlocal
rem Work To-Do uninstaller. Your data is kept unless you choose to delete it.
if exist "%LOCALAPPDATA%\WorkToDo\Uninstall-WorkToDo.ps1" if "%~1"=="" (
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%LOCALAPPDATA%\WorkToDo\Uninstall-WorkToDo.ps1"
  exit /b %errorlevel%
)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0installer\Uninstall-WorkToDo.ps1" %*
exit /b %errorlevel%
