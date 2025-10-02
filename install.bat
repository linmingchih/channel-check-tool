@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "VENV_DIR=%SCRIPT_DIR%.venv"
set "PYTHON_EXE=%VENV_DIR%\Scripts\python.exe"
set "UV_EXE=%VENV_DIR%\Scripts\uv.exe"
set "REQUIREMENTS_FILE=%SCRIPT_DIR%requirements.txt"

set "BOOTSTRAP_PY="
where py >nul 2>nul && set "BOOTSTRAP_PY=py"
if not defined BOOTSTRAP_PY (
    where python >nul 2>nul && set "BOOTSTRAP_PY=python"
)

if not defined BOOTSTRAP_PY (
    echo Could not locate a Python interpreter to create the virtual environment.
    exit /b 1
)

if not exist "%VENV_DIR%\NUL" (
    %BOOTSTRAP_PY% -m venv "%VENV_DIR%"
    if errorlevel 1 (
        echo Failed to create the virtual environment.
        exit /b 1
    )
) else (
    echo Virtual environment already exists at "%VENV_DIR%".
)

if not exist "%PYTHON_EXE%" (
    echo Virtual environment Python interpreter not found after creation.
    exit /b 1
)

"%PYTHON_EXE%" -m pip install --upgrade pip setuptools wheel uv
if errorlevel 1 (
    echo Failed to upgrade packaging tools or install uv in the virtual environment.
    exit /b 1
)

if not exist "%UV_EXE%" (
    echo uv was not found after installation.
    exit /b 1
)

if exist "%REQUIREMENTS_FILE%" (
    "%UV_EXE%" pip install -r "%REQUIREMENTS_FILE%"
) else (
    "%UV_EXE%" pip install "pyedb>=0.6.0" "PySide6>=6.5" "pyaedt" "numpy>=1.24" "scikit-rf"
)
if errorlevel 1 (
    echo Failed to install required Python packages with uv.
    exit /b 1
)

echo Virtual environment setup complete.
exit /b 0
