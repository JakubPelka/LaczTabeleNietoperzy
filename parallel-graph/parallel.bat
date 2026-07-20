@echo off
setlocal

set "PROJECT_DIR=%~dp0"

where py >nul 2>nul
if %errorlevel%==0 (
    py -3 "%PROJECT_DIR%parallel.py"
    exit /b %errorlevel%
)

where python >nul 2>nul
if %errorlevel%==0 (
    python "%PROJECT_DIR%parallel.py"
    exit /b %errorlevel%
)

echo Nie znaleziono Pythona.
echo Uruchom parallel.py za pomoca kompletnego Pythona z USB.
exit /b 1
