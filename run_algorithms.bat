@echo off
setlocal

set "PROJECT_DIR=%~dp0"
set "VENV_DIR=%PROJECT_DIR%.venv"
set "PYTHON_BIN=%VENV_DIR%\Scripts\python.exe"

if not exist "%PYTHON_BIN%" (
    py -3 -m venv "%VENV_DIR%" || goto :error
    "%PYTHON_BIN%" -m pip install --upgrade pip || goto :error
    "%PYTHON_BIN%" -m pip install -r "%PROJECT_DIR%requirements.txt" || goto :error
)

echo Wybierz algorytm:
echo   1^) Tabele i wykresy zbiorcze
echo   2^) Tabele oraz wykresy zbiorcze i/lub noc-po-nocy
set /p "SELECTION=Wybor [1-2]: "

if "%SELECTION%"=="1" (
    "%PYTHON_BIN%" "%PROJECT_DIR%algorithms\merge_summary_charts.py"
) else if "%SELECTION%"=="2" (
    "%PYTHON_BIN%" "%PROJECT_DIR%algorithms\merge_summary_and_nightly_charts.py"
) else (
    echo Nieprawidlowy wybor.
    exit /b 2
)

exit /b %ERRORLEVEL%

:error
echo Nie udalo sie przygotowac srodowiska Python.
exit /b 1
