@echo off
echo.
echo ============================================
echo   Transformador de Bases Comprometido
echo ============================================
echo.

cd /d "%~dp0"

echo [1/3] Verificando Node.js...
node --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Node.js no esta instalado. Instalo desde https://nodejs.org
    pause
    exit /b 1
)

echo [2/3] Ejecutando transformador...
node transformar.js

echo.
echo [3/3] Proceso completado.
echo.
pause