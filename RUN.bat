@echo off
setlocal
rem ==================================================================
rem  Arranca la aplicacion desde el codigo fuente (modo desarrollo).
rem  La primera vez crea el entorno virtual e instala dependencias.
rem  Para el usuario final se distribuye el instalador (ver README).
rem ==================================================================

cd /d "%~dp0"
title Iberostar Gestor de Pedidos (desarrollo)

if not exist ".venv\Scripts\python.exe" (
    where python >nul 2>&1 || (
        echo [ERROR] Python 3.11 o superior no esta instalado o no esta en el PATH.
        pause
        exit /b 1
    )
    echo Preparando el entorno virtual por primera vez...
    python -m venv .venv || goto :error
    ".venv\Scripts\python.exe" -m pip install --quiet --upgrade pip || goto :error
    ".venv\Scripts\python.exe" -m pip install --quiet -r requirements-dev.txt || goto :error
)

".venv\Scripts\python.exe" app\backend\desktop.py
exit /b %errorlevel%

:error
echo [ERROR] No se pudo preparar el entorno. Revisa los mensajes anteriores.
pause
exit /b 1
