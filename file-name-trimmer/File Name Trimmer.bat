@echo off
setlocal
 title File Name Trimmer
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0FileNameTrimmer.ps1"
if errorlevel 1 (
    echo.
    echo File Name Trimmer could not start. See the message above.
    pause
)
endlocal
