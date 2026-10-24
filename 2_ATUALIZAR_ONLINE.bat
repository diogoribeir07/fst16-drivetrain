@echo off
chcp 65001 >nul
title FST 16 - Publicar / Atualizar no GitHub (Web + Streamlit)
echo ========================================================
echo       FST 16 - Publicar / Atualizar no GitHub
echo ========================================================
echo.
echo Repositorio: https://github.com/diogoribeir07/fst16-drivetrain.git
echo Web 24/7:   https://diogoribeir07.github.io/fst16-drivetrain/
echo Streamlit:  https://fst16-drivetrain.streamlit.app
echo.

set REPO_URL=https://github.com/diogoribeir07/fst16-drivetrain.git

cd /d "%~dp0"

echo [1/3] A sincronizar dados dos ficheiros Excel para a Web...
set "PY_CMD=python"
if exist "%LOCALAPPDATA%\Programs\Python\Python314\python.exe" (
    set "PY_CMD=%LOCALAPPDATA%\Programs\Python\Python314\python.exe"
)
"%PY_CMD%" sync_cloud.py export >nul 2>&1

git remote get-url origin >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    git remote add origin %REPO_URL%
) else (
    git remote set-url origin %REPO_URL%
)
git branch -M main >nul 2>&1

echo [2/3] A preparar ficheiros e alteracoes...
git add .
git commit -m "Update drivetrain tracker data and web portal" >nul 2>&1

echo [3/3] A enviar para o GitHub (main)...
git push -u origin main

echo.
echo ========================================================
echo Envio concluido!
echo.
echo 1. Web Portal (GitHub Pages, 0 s arranque, NUNCA hiberna):
echo    https://diogoribeir07.github.io/fst16-drivetrain/
echo.
echo 2. Streamlit Cloud:
echo    https://fst16-drivetrain.streamlit.app
echo ========================================================
pause
