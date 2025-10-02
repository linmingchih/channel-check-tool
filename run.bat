@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "VENV_DIR=%SCRIPT_DIR%.venv"
set "PYTHON_EXE=%VENV_DIR%\Scripts\python.exe"
set "QT_PLUGIN_BASE=%VENV_DIR%\Lib\site-packages\PySide6\plugins"

if not exist "%PYTHON_EXE%" (
    echo Virtual environment not found. Please run install.bat first.
    exit /b 1
)

if exist "%QT_PLUGIN_BASE%" (
    set "QT_PLUGIN_PATH=%QT_PLUGIN_BASE%"
    set "QT_QPA_PLATFORM_PLUGIN_PATH=%QT_PLUGIN_BASE%\platforms"
)

"%PYTHON_EXE%" "%SCRIPT_DIR%src\aedb_gui.py"
set "EXIT_CODE=%errorlevel%"
exit /b %EXIT_CODE%
