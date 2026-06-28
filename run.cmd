@echo off
setlocal enableextensions
cd /d "%~dp0"
chcp 65001 >nul 2>&1

echo ============================================
echo    HR Letter Generator
echo ============================================
echo.

call :find_python
if not defined PY (
  echo Python was not found. Attempting to install it via winget...
  where winget >nul 2>&1
  if errorlevel 1 goto :no_python_no_winget
  winget install --id Python.Python.3.12 -e --source winget --accept-package-agreements --accept-source-agreements
  call :find_python
  if not defined PY goto :installed_reopen
)

echo Using Python:
%PY% --version
echo.

if not exist ".venv\Scripts\python.exe" (
  echo Creating virtual environment...
  %PY% -m venv .venv
  if errorlevel 1 goto :venv_fail
)
set "VENVPY=.venv\Scripts\python.exe"

if not exist ".venv\.deps_ok" (
  echo Installing required packages, please wait...
  "%VENVPY%" -m pip install --upgrade pip
  "%VENVPY%" -m pip install -r requirements.txt
  if errorlevel 1 goto :deps_fail
  echo ok>".venv\.deps_ok"
)

echo.
"%VENVPY%" generate_letter.py
echo.
pause
exit /b 0

:find_python
set "PY="
where py >nul 2>&1
if not errorlevel 1 (
  set "PY=py"
  goto :eof
)
where python >nul 2>&1
if not errorlevel 1 set "PY=python"
goto :eof

:no_python_no_winget
echo.
echo [!] winget is not available, so Python cannot be installed automatically.
echo     Please install Python 3 from https://www.python.org/downloads/
echo     (tick "Add python.exe to PATH" during setup), then run this file again.
echo.
pause
exit /b 1

:installed_reopen
echo.
echo Python has been installed, but this window must reopen to detect it.
echo Please CLOSE this window and double-click run.cmd again.
echo.
pause
exit /b 0

:venv_fail
echo [!] Failed to create the virtual environment.
pause
exit /b 1

:deps_fail
echo [!] Failed to install required packages. Check your internet connection.
pause
exit /b 1
