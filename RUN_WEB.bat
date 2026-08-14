@echo off
setlocal

title Iberostar Inventory Synchronizer - Web

cd /d "%~dp0"

echo.
echo ==================================================
echo    IBEROSTAR INVENTORY SYNCHRONIZER - WEB
echo ==================================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python no esta instalado o no esta en el PATH.
    echo.
    pause
    exit /b 1
)

start "" cmd /c "timeout /t 2 /nobreak >nul & start http://127.0.0.1:5000/"

python app\backend\api.py

echo.
echo ==================================================
echo              SERVIDOR DETENIDO
echo ==================================================
echo.

pause
endlocal
