@echo off
rem SmartGrid launcher for Windows: double-click this file.
rem First run: creates the .venv, installs the dependencies and a "SmartGrid"
rem shortcut on the Desktop, then starts SmartGrid. Later runs start it directly.
setlocal EnableExtensions
chcp 65001 >nul
title SmartGrid

rem Unlike cd, pushd also accepts a network or shared folder (\\server\...).
pushd "%~dp0" 2>nul || (
    echo SmartGrid: cannot open the folder "%~dp0".
    goto fail
)

set "VENV_PY=.venv\Scripts\python.exe"
set "VENV_PYW=.venv\Scripts\pythonw.exe"
set "CHECK=import sys; sys.exit(0 if sys.platform == 'win32' and sys.version_info >= (3, 11) else 1)"

if not exist "smartgrid.py" (
    echo SmartGrid: smartgrid.py was not found next to this file.
    echo Keep smartgrid.bat in the SmartGrid folder and create a shortcut to it instead.
    goto fail
)
if not exist "requirements.txt" (
    echo SmartGrid: requirements.txt was not found.
    goto fail
)

if exist "%VENV_PY%" goto check_venv
if exist ".venv" (
    echo SmartGrid: the .venv folder exists but contains no Windows Python.
    echo Delete the .venv folder, then run smartgrid.bat again.
    goto fail
)

rem Look for Windows Python 3.11 or later.
set "BOOT="
py -3 -c "%CHECK%" >nul 2>&1 && set "BOOT=py -3"
if not defined BOOT python -c "%CHECK%" >nul 2>&1 && set "BOOT=python"
if not defined BOOT goto no_python

echo First run: preparing SmartGrid, this can take a few minutes...
%BOOT% -m venv .venv || (
    echo SmartGrid: creating the virtual environment failed.
    goto fail
)

:check_venv
"%VENV_PY%" -c "%CHECK%" >nul 2>&1 || (
    echo SmartGrid: the .venv does not run Windows Python 3.11 or later.
    echo Delete the .venv folder, then run smartgrid.bat again.
    goto fail
)
"%VENV_PY%" -m smartgrid.deps >nul 2>&1 && goto launch

echo Installing dependencies (PySide6, Pillow)...
"%VENV_PY%" -m pip --version >nul 2>&1 || "%VENV_PY%" -m ensurepip --upgrade >nul || (
    echo SmartGrid: cannot initialize pip in the virtual environment.
    goto fail
)
"%VENV_PY%" -m pip install --upgrade pip --disable-pip-version-check --retries 1 --timeout 10 >nul 2>&1
"%VENV_PY%" -m pip install --disable-pip-version-check -r requirements.txt || (
    echo SmartGrid: installation failed. Check the Internet connection; the next run will try again.
    goto fail
)
"%VENV_PY%" -m smartgrid.deps || (
    echo SmartGrid: dependencies are still unavailable after installation.
    goto fail
)

:launch
call :desktop_shortcut
rem Full paths and the original folder: the temporary drive letter created by
rem pushd disappears at popd, so SmartGrid must not depend on it.
set "SG_DIR=%~dp0"
if "%SG_DIR:~-1%"=="\" set "SG_DIR=%SG_DIR:~0,-1%"
rem pythonw starts SmartGrid without a console; its icon appears near the clock.
if exist "%VENV_PYW%" (
    start "" /D "%SG_DIR%" "%SG_DIR%\%VENV_PYW%" "%SG_DIR%\smartgrid.py" %*
) else (
    start "" /min /D "%SG_DIR%" "%SG_DIR%\%VENV_PY%" "%SG_DIR%\smartgrid.py" %*
)
popd
endlocal
exit /b 0

:desktop_shortcut
rem Once per installation: create the "SmartGrid" Desktop shortcut, or repair
rem existing ones. The shortcut starts pythonw directly: no console, hence no
rem window and no cmd warning about UNC paths. smartgrid.bat remains the
rem installer and repair tool. A deleted shortcut is not recreated.
if exist ".venv\.smartgrid-shortcut-v3" exit /b 0
if not exist "%VENV_PYW%" exit /b 0
set "SG_FOLDER=%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$shell = New-Object -ComObject WScript.Shell;" ^
  "$desktop = [Environment]::GetFolderPath('Desktop');" ^
  "$folder = $env:SG_FOLDER.TrimEnd('\');" ^
  "$pythonw = Join-Path $folder '.venv\Scripts\pythonw.exe';" ^
  "function Set-SmartGrid($shortcut) {" ^
  "  $shortcut.TargetPath = $pythonw;" ^
  "  $shortcut.Arguments = [char]34 + (Join-Path $folder 'smartgrid.py') + [char]34;" ^
  "  $shortcut.WorkingDirectory = $folder;" ^
  "  $shortcut.WindowStyle = 1;" ^
  "  $shortcut.Description = 'SmartGrid';" ^
  "  $shortcut.IconLocation = $pythonw + ',0';" ^
  "  $shortcut.Save() }" ^
  "$found = $false;" ^
  "foreach ($file in Get-ChildItem -LiteralPath $desktop -Filter *.lnk) {" ^
  "  $shortcut = $shell.CreateShortcut($file.FullName);" ^
  "  if ($shortcut.TargetPath -like '*\smartgrid.bat' -or ($shortcut.TargetPath -like '*\pythonw.exe' -and $shortcut.Arguments -like '*smartgrid.py*')) {" ^
  "    $found = $true; Set-SmartGrid $shortcut } }" ^
  "if (-not $found -and -not (Test-Path -LiteralPath (Join-Path $desktop 'SmartGrid.lnk'))) {" ^
  "  Set-SmartGrid $shell.CreateShortcut((Join-Path $desktop 'SmartGrid.lnk')) }" >nul 2>&1 && (
    echo.> ".venv\.smartgrid-shortcut-v3"
    echo SmartGrid shortcut added to the Desktop.
)
exit /b 0

:no_python
echo.
echo SmartGrid needs Python 3.11 or later for Windows, which was not found on this PC.
echo Install Python, then run smartgrid.bat again.
goto fail

:fail
echo.
pause
popd 2>nul
endlocal
exit /b 1
