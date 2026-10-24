@echo off
chcp 65001 >nul
title FST 16 - Descarregar da Cloud
cd /d "%~dp0"
echo ========================================================
echo       FST 16 - Descarregar Alteracoes da Cloud
echo ========================================================
echo.
echo [1/2] A puxar alteracoes do repositorio GitHub (git pull)...
git pull origin main
echo.
echo [2/2] A importar dados da Web (data/*.json) para Excel...
set "PY_CMD=python"
if exist "%LOCALAPPDATA%\Programs\Python\Python314\python.exe" (
    set "PY_CMD=%LOCALAPPDATA%\Programs\Python\Python314\python.exe"
)
"%PY_CMD%" sync_cloud.py import
echo.
echo ========================================================
echo Sincronizacao concluida com sucesso.
echo Ficheiros Excel locais atualizados com a Cloud.
echo ========================================================
pause
