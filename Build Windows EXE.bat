@echo off
setlocal
cd /d "%~dp0"
echo Building BearsGrader.exe. Python 3.12 must be installed.
py -3.12 -m venv .venv-build
if errorlevel 1 goto failed
".venv-build\Scripts\python.exe" -m pip install -r requirements-build.txt
if errorlevel 1 goto failed
".venv-build\Scripts\python.exe" -m pytest -q
if errorlevel 1 goto failed
".venv-build\Scripts\python.exe" -m PyInstaller --noconfirm --clean --onefile --windowed --name BearsGrader --add-data "assets;assets" desktop.py
if errorlevel 1 goto failed
".venv-build\Scripts\python.exe" desktop.py --self-test --output desktop-smoke.json
if errorlevel 1 goto failed
echo Finished. Open dist\BearsGrader.exe to use the app.
pause
exit /b 0
:failed
echo Build failed. Read the error above. Python 3.12 and internet access are required.
pause
exit /b 1
