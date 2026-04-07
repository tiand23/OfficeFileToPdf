@echo off
setlocal

chcp 65001 >nul
set "PYTHONUTF8=1"

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

set "LOG_FILE=%SCRIPT_DIR%ToPdf_run.log"
set "PYTHON_LOG=%SCRIPT_DIR%ToPdf_python.log"

set "SOURCE_PATH=%~1"
if "%SOURCE_PATH%"=="" set "SOURCE_PATH=%SCRIPT_DIR%"
for %%I in ("%SOURCE_PATH%") do set "SOURCE_PATH=%%~fI"

echo ===== ToPdf start ===== > "%LOG_FILE%"
echo ScriptDir: %SCRIPT_DIR% >> "%LOG_FILE%"
echo SourcePath: %SOURCE_PATH% >> "%LOG_FILE%"
echo Requirement: Microsoft Office must be installed for Office document conversion. >> "%LOG_FILE%"

call :resolve_python
if errorlevel 1 goto :fail

if not exist "%SCRIPT_DIR%.venv\Scripts\python.exe" (
    echo [1/5] Creating virtual environment...
    echo [1/5] create venv >> "%LOG_FILE%"
    "%BOOTSTRAP_PYTHON%" %BOOTSTRAP_ARGS% -m venv "%SCRIPT_DIR%.venv"
    if errorlevel 1 (
        echo Failed to create virtual environment.
        echo create venv failed >> "%LOG_FILE%"
        goto :fail
    )
) else (
    echo [1/5] Virtual environment already exists.
    echo [1/5] venv already exists >> "%LOG_FILE%"
)

call "%SCRIPT_DIR%.venv\Scripts\activate.bat"
if errorlevel 1 (
    echo Failed to activate virtual environment.
    echo activate venv failed >> "%LOG_FILE%"
    goto :fail
)

echo [2/5] Upgrading pip...
echo [2/5] upgrade pip >> "%LOG_FILE%"
python -m pip install --upgrade pip
if errorlevel 1 (
    echo Failed to upgrade pip.
    echo upgrade pip failed >> "%LOG_FILE%"
    goto :fail
)

echo [3/5] Installing Python packages...
echo [3/5] install requirements >> "%LOG_FILE%"
python -m pip install -r "%SCRIPT_DIR%requirements.txt"
if errorlevel 1 (
    echo Failed to install Python packages.
    echo install requirements failed >> "%LOG_FILE%"
    goto :fail
)

echo [4/5] Microsoft Office COM mode only.
echo [4/5] microsoft office com only >> "%LOG_FILE%"

echo [5/5] Processing: "%SOURCE_PATH%"
echo [5/5] run python script >> "%LOG_FILE%"
python "%SCRIPT_DIR%ToPdf.py" "%SOURCE_PATH%" --output-dir res > "%PYTHON_LOG%" 2>&1
set "EXIT_CODE=%ERRORLEVEL%"

if not "%EXIT_CODE%"=="0" (
    echo Script failed with exit code: %EXIT_CODE%
    echo run failed, exit code: %EXIT_CODE% >> "%LOG_FILE%"
    goto :fail_with_code
)

echo success >> "%LOG_FILE%"
echo Done.
echo Run log: "%LOG_FILE%"
echo Python log: "%PYTHON_LOG%"
pause
exit /b 0

:resolve_python
py -3 -c "import sys" >nul 2>&1
if not errorlevel 1 (
    set "BOOTSTRAP_PYTHON=py"
    set "BOOTSTRAP_ARGS=-3"
    exit /b 0
)

python -c "import sys" >nul 2>&1
if not errorlevel 1 (
    set "BOOTSTRAP_PYTHON=python"
    set "BOOTSTRAP_ARGS="
    exit /b 0
)

where winget >nul 2>&1
if errorlevel 1 (
    echo Python was not found and winget is unavailable.
    echo python missing and winget unavailable >> "%LOG_FILE%"
    exit /b 1
)

echo Python was not found. Trying to install Python 3.12 with winget...
echo install python via winget >> "%LOG_FILE%"
winget install -e --id Python.Python.3.12 --accept-package-agreements --accept-source-agreements
if errorlevel 1 (
    echo Python install failed.
    echo install python failed >> "%LOG_FILE%"
    exit /b 1
)

py -3 -c "import sys" >nul 2>&1
if not errorlevel 1 (
    set "BOOTSTRAP_PYTHON=py"
    set "BOOTSTRAP_ARGS=-3"
    exit /b 0
)

python -c "import sys" >nul 2>&1
if not errorlevel 1 (
    set "BOOTSTRAP_PYTHON=python"
    set "BOOTSTRAP_ARGS="
    exit /b 0
)

echo Python is still unavailable after installation. Reopen the terminal and try again.
echo python installed but still unavailable >> "%LOG_FILE%"
exit /b 1

:fail_with_code
echo.
echo Failed.
echo Run log: "%LOG_FILE%"
echo Python log: "%PYTHON_LOG%"
pause
exit /b %EXIT_CODE%

:fail
set "EXIT_CODE=1"
echo.
echo Failed.
echo Run log: "%LOG_FILE%"
if exist "%PYTHON_LOG%" echo Python log: "%PYTHON_LOG%"
pause
exit /b %EXIT_CODE%
