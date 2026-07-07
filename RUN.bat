@echo off
setlocal

title Iberostar Inventory Synchronizer

cd /d "%~dp0"

echo.
echo ==================================================
echo         IBEROSTAR INVENTORY SYNCHRONIZER
echo ==================================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python no esta instalado o no esta en el PATH.
    echo.
    pause
    exit /b 1
)

python app\backend\main.py

echo.
echo ==================================================
echo              EXECUTION FINISHED
echo ==================================================
echo.

pause
endlocal