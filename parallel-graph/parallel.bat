@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if %errorlevel%==0 (
    set "PYTHON=py -3"
) else (
    where python >nul 2>nul
    if errorlevel 1 (
        echo Nie znaleziono Pythona. Zainstaluj Python 3.10 lub nowszy z python.org.
        pause
        exit /b 1
    )
    set "PYTHON=python"
)

if not exist ".venv\Scripts\python.exe" (
    %PYTHON% -m venv .venv
    if errorlevel 1 goto :error
)

.venv\Scripts\python.exe -c "import openpyxl, plotly" >nul 2>nul
if errorlevel 1 (
    .venv\Scripts\python.exe -m pip install --upgrade pip
    if errorlevel 1 goto :error
    .venv\Scripts\python.exe -m pip install -e .
    if errorlevel 1 goto :error
)

.venv\Scripts\python.exe -m parallel_graph
if errorlevel 1 goto :error
exit /b 0

:error
echo.
echo Uruchomienie nie powiodlo sie. Szczegoly bledu znajduja sie powyzej.
pause
exit /b 1

