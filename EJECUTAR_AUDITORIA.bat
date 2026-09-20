@echo off
title APP_REVISA_EXCEL_V2 - Auditoria de Trabajos

echo ==========================================================
echo   APP_REVISA_EXCEL_V2 - AUDITORIA Y EVALUACION POR RUBRICA
echo ==========================================================
echo.

cd /d "%~dp0"

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python no esta instalado o no esta en el PATH.
    echo.
    pause
    exit /b 1
)

set VENV_DIR=.venv
if not exist "%VENV_DIR%\Scripts\python.exe" (
    echo [INFO] Creando entorno virtual aislado en %VENV_DIR%...
    python -m venv %VENV_DIR%
)

call %VENV_DIR%\Scripts\activate.bat

if exist "requirements.txt" (
    echo [INFO] Verificando dependencias...
    python -m pip install -q -r requirements.txt
)

if not exist "main.py" (
    echo [ERROR] No se encontro el archivo main.py
    echo.
    pause
    exit /b 1
)

echo [INFO] Iniciando auditoria...
echo.

python main.py

echo.
echo ==========================================================
if %errorlevel% equ 0 (
    echo   [EXITO] La auditoria finalizo correctamente sin errores.
) else (
    echo   [ERROR] La auditoria finalizo con codigo de salida: %errorlevel%.
)
echo ==========================================================
echo.
pause