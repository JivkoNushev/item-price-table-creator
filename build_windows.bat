@echo off
setlocal enabledelayedexpansion

echo ============================================
echo   Table Generator - Windows Build Script
echo ============================================
echo.

cd /d "%~dp0"

:: -----------------------------------------------
:: Step 1: Check for Python
:: -----------------------------------------------
echo [1/5] Checking for Python...

set "PYTHON_CMD="

:: Try 'py' launcher first (works even if not in PATH)
where py >nul 2>&1
if %errorlevel% equ 0 (
    py -3 --version >nul 2>&1
    if !errorlevel! equ 0 (
        set "PYTHON_CMD=py -3"
        goto :found_python
    )
)

:: Try 'python' directly
where python >nul 2>&1
if %errorlevel% equ 0 (
    python --version 2>&1 | findstr /R "^Python 3\." >nul
    if !errorlevel! equ 0 (
        set "PYTHON_CMD=python"
        goto :found_python
    )
)

:: Try 'python3'
where python3 >nul 2>&1
if %errorlevel% equ 0 (
    set "PYTHON_CMD=python3"
    goto :found_python
)

:: -----------------------------------------------
:: Step 2: Install Python if not found
:: -----------------------------------------------
echo    Python not found. Installing Python 3.12...
echo    This requires internet and may take ~30 seconds.
echo.

set "PY_VER=3.12.7"
set "PY_INSTALLER=python-%PY_VER%-amd64.exe"
set "PY_URL=https://www.python.org/ftp/python/%PY_VER%/%PY_INSTALLER%"

:: Download Python installer
echo    Downloading Python %PY_VER%...
powershell -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '%PY_URL%' -OutFile '%PY_INSTALLER%'" 2>nul
if %errorlevel% neq 0 (
    echo    ERROR: Failed to download Python.
    echo    Please install Python 3.12+ manually from https://www.python.org
    goto :build_error
)

:: Run silent install (no admin required)
echo    Installing Python (silent, no admin needed)...
"%PY_INSTALLER%" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0 Include_launcher=1
if %errorlevel% neq 0 (
    echo    ERROR: Python installation failed.
    del "%PY_INSTALLER%" 2>nul
    goto :build_error
)

:: Clean up installer
del "%PY_INSTALLER%" 2>nul

:: Refresh PATH
set "PATH=%LOCALAPPDATA%\Programs\Python\Python312;%LOCALAPPDATA%\Programs\Python\Python312\Scripts;%PATH%"

:: Verify installation
set "PYTHON_CMD=python"
%PYTHON_CMD% --version >nul 2>&1
if %errorlevel% neq 0 (
    set "PYTHON_CMD=py -3"
    %PYTHON_CMD% --version >nul 2>&1
    if !errorlevel! neq 0 (
        echo    ERROR: Python installed but not reachable.
        echo    Try closing and re-opening this window, or install manually.
        goto :build_error
    )
)

echo    Python installed successfully.
echo.

:found_python
%PYTHON_CMD% --version
echo.

:: -----------------------------------------------
:: Step 3: Create virtual environment
:: -----------------------------------------------
echo [2/5] Creating virtual environment...

if exist ".venv" rmdir /s /q ".venv"
%PYTHON_CMD% -m venv .venv
if %errorlevel% neq 0 (
    echo    ERROR: Failed to create virtual environment.
    goto :build_error
)

:: Use venv python and pip
set "VENV_PY=.venv\Scripts\python.exe"
set "VENV_PIP=.venv\Scripts\pip.exe"

echo    Done.
echo.

:: -----------------------------------------------
:: Step 4: Install dependencies
:: -----------------------------------------------
echo [3/5] Installing build dependencies...
echo    (pyinstaller + openpyxl)

%VENV_PIP% install --quiet --upgrade pip
%VENV_PIP% install --quiet pyinstaller openpyxl
if %errorlevel% neq 0 (
    echo    ERROR: Failed to install dependencies.
    goto :build_error
)

echo    Done.
echo.

:: -----------------------------------------------
:: Step 5: Build with PyInstaller
:: -----------------------------------------------
echo [4/5] Building TableGenerator.exe...
echo    This may take 1-2 minutes.
echo.

%VENV_PY% -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name "TableGenerator" ^
    --hidden-import openpyxl.cell._writer ^
    --collect-submodules openpyxl ^
    --distpath dist ^
    --workpath build ^
    --specpath . ^
    main.py

if %errorlevel% neq 0 (
    echo    ERROR: Build failed.
    goto :build_error
)

echo    Done.
echo.

:: -----------------------------------------------
:: Step 6: Clean up
:: -----------------------------------------------
echo [5/5] Cleaning up build artifacts...

if exist "build" rmdir /s /q "build"
if exist "TableGenerator.spec" del "TableGenerator.spec"
if exist ".venv" rmdir /s /q ".venv"

echo    Done.
echo.

:: -----------------------------------------------
:: Done
:: --------------------------------===============
echo ============================================
echo   BUILD SUCCESSFUL
echo ============================================
echo.
echo   Output: dist\TableGenerator.exe
echo.
echo   This is a standalone executable.
echo   Copy it to any Windows machine and run it.
echo   No Python installation required.
echo.

if exist "dist\TableGenerator.exe" (
    echo   Would you like to open the output folder? (Y/N)
    choice /c YN /n /m "> "
    if !errorlevel! equ 1 (
        explorer dist
    )
)

goto :eof

:build_error
echo.
echo ============================================
echo   BUILD FAILED
echo ============================================
echo.
echo   Check the errors above and try again.
echo.
pause
exit /b 1
