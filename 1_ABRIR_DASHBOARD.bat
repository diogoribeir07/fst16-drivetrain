@echo off
chcp 65001 >nul
title FST 16 - Drivetrain Web Dashboard
cd /d "%~dp0"
echo ========================================================
echo          FST 16 - Drivetrain Web Dashboard
echo ========================================================
echo.
echo A iniciar a aplicacao Web...
echo.
echo 1. Web Online (GitHub Pages 24/7):
echo    https://diogoribeir07.github.io/fst16-drivetrain/
echo.
echo 2. Servidor Local (Offline):
echo    http://localhost:8000/
echo.
echo ========================================================
start "" "http://localhost:8000/"
echo A escutar na porta 8000 (Ctrl+C para encerrar)...
set "PY_CMD=python"
if exist "%LOCALAPPDATA%\Programs\Python\Python314\python.exe" (
    set "PY_CMD=%LOCALAPPDATA%\Programs\Python\Python314\python.exe"
)
"%PY_CMD%" -m http.server 8000
pause
