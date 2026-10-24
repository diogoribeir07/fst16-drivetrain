@echo off
chcp 65001 >nul
title FST 16 - Sincronizar Excel com a Web
cd /d "%~dp0"
echo ========================================================
echo       FST 16 - Sincronizar Excel com a Web
echo ========================================================
echo.
echo [1/2] A exportar dados dos ficheiros Excel para a Cloud...
set "PY_CMD=python"
if exist "%LOCALAPPDATA%\Programs\Python\Python314\python.exe" (
    set "PY_CMD=%LOCALAPPDATA%\Programs\Python\Python314\python.exe"
)
"%PY_CMD%" sync_cloud.py export
echo.
echo [2/2] Sincronizacao local concluida!
echo Para enviar as alteracoes para o GitHub Pages (24/7), corre:
echo   2_ATUALIZAR_ONLINE.bat
echo ========================================================
echo.
pause
