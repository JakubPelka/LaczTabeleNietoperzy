@echo off
setlocal

set "PROJECT_DIR=%~dp0"
set "VENDOR_DIR=%PROJECT_DIR%vendor"

where py >nul 2>nul
if not errorlevel 1 goto prepare_with_py

where python >nul 2>nul
if not errorlevel 1 goto prepare_with_python

echo Nie znaleziono Pythona.
echo Zainstaluj Python 3.10 lub nowszy i uruchom skrypt ponownie.
exit /b 1

:prepare_with_py
py -3 -m pip --version >nul 2>nul
if errorlevel 1 goto pip_missing
echo Pobieranie bibliotek do: %VENDOR_DIR%
py -3 -m pip install --disable-pip-version-check --upgrade --target "%VENDOR_DIR%" --requirement "%PROJECT_DIR%requirements.txt"
if errorlevel 1 exit /b %errorlevel%
py -3 "%PROJECT_DIR%parallel.py"
exit /b %errorlevel%

:prepare_with_python
python -m pip --version >nul 2>nul
if errorlevel 1 goto pip_missing
echo Pobieranie bibliotek do: %VENDOR_DIR%
python -m pip install --disable-pip-version-check --upgrade --target "%VENDOR_DIR%" --requirement "%PROJECT_DIR%requirements.txt"
if errorlevel 1 exit /b %errorlevel%
python "%PROJECT_DIR%parallel.py"
exit /b %errorlevel%

:pip_missing
echo Brakuje pip. Zainstaluj pip dla uzywanego Pythona i uruchom skrypt ponownie.
exit /b 1
