@echo off
title APP_REVISA_EXCEL_V2 - Panel de Control

echo ==========================================================
echo   APP_REVISA_EXCEL_V2 - PANEL DE CONTROL (MODO HIBRIDO)
echo ==========================================================
echo.

cd /d "%~dp0"

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python no esta instalado o no esta en el PATH de Windows.
    echo         Instalalo desde https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

set VENV_DIR=.venv
if not exist "%VENV_DIR%\Scripts\python.exe" (
    echo [INFO] Creando entorno virtual en %VENV_DIR%...
    python -m venv %VENV_DIR%
)

call %VENV_DIR%\Scripts\activate.bat

if exist "requirements.txt" (
    echo [INFO] Verificando dependencias...
    python -m pip install -q -r requirements.txt
)

if not exist "gui_server.py" (
    echo [ERROR] No se encontro gui_server.py
    echo.
    pause
    exit /b 1
)

echo [INFO] Iniciando servidor de interfaz grafica local...
echo.

python gui_server.py

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Ocurrio un problema al ejecutar la interfaz grafica.
    echo.
    pause
)