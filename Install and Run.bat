@echo off
setlocal
cd /d "%~dp0"
if errorlevel 1 goto failed

echo BCIT Word Project Grader - BSYS 1000 Fall 2026 - Windows setup
echo Keep this window open while using the app.
echo.
if not exist "desktop.py" goto missing_files
if not exist "requirements-desktop.txt" goto missing_files

rem Prefer supported Python versions with wheels for the pinned dependencies.
call :find_python
if defined PYTHON_CMD goto python_ready

echo Python 3.11-3.13 was not found. Installing Python 3.12 for your user...
where winget >nul 2>&1
if errorlevel 1 goto no_winget
winget install --id Python.Python.3.12 --exact --source winget --scope user --silent --accept-package-agreements --accept-source-agreements --disable-interactivity
if errorlevel 1 goto failed
rem A new installation may not be on PATH in this existing command window.
call :find_python
if not defined PYTHON_CMD goto no_python

:python_ready
echo Using %PYTHON_CMD%
if exist ".venv-windows\Scripts\python.exe" goto check_venv
%PYTHON_CMD% -m venv ".venv-windows"
if errorlevel 1 goto failed

:check_venv
".venv-windows\Scripts\python.exe" -c "import sys; sys.exit(0 if (3,11) <= sys.version_info[:2] < (3,14) else 1)"
if errorlevel 1 goto bad_venv
".venv-windows\Scripts\python.exe" -m pip install --disable-pip-version-check -r "requirements-desktop.txt"
if errorlevel 1 goto failed

echo.
echo Starting the desktop app.
echo Close the app window when you are finished.
".venv-windows\Scripts\python.exe" "desktop.py"
if errorlevel 1 goto failed
exit /b 0

:find_python
set "PYTHON_CMD="
call :try_python py -3.12
if defined PYTHON_CMD exit /b 0
call :try_python py -3.11
if defined PYTHON_CMD exit /b 0
call :try_python py -3.13
if defined PYTHON_CMD exit /b 0
call :try_python "%LocalAppData%\Programs\Python\Python312\python.exe"
if defined PYTHON_CMD exit /b 0
call :try_python python
exit /b 0

:try_python
%* -c "import sys; sys.exit(0 if (3,11) <= sys.version_info[:2] < (3,14) else 1)" >nul 2>&1
if not errorlevel 1 set PYTHON_CMD=%*
exit /b 0

:missing_files
echo ERROR: Extract the entire GitHub ZIP first, then run this file inside the extracted folder.
goto failed
:no_winget
echo ERROR: Windows Package Manager ^(winget^) is unavailable.
echo Install Python 3.12 from https://www.python.org/downloads/windows/
echo Choose "Add python.exe to PATH", then run this file again.
goto failed
:no_python
echo ERROR: Python was installed but could not be found yet.
echo Close this window and run this file again. If needed, restart Windows.
goto failed
:bad_venv
echo ERROR: The .venv-windows environment is incompatible or damaged.
echo Rename or delete only the .venv-windows folder, then run this file again.
goto failed
:failed
echo.
echo Setup or startup did not finish. Read the error above.
echo Internet access is required for the first installation.
pause
exit /b 1
