@echo off
title FST 16 - Publicar / Atualizar no GitHub
echo ========================================================
echo       FST 16 - Publicar / Atualizar no GitHub
echo ========================================================
echo.
echo Repositorio: https://github.com/diogoribeir07/fst16-drivetrain.git
echo Dashboard:   https://fst16-drivetrain.streamlit.app
echo.

set REPO_URL=https://github.com/diogoribeir07/fst16-drivetrain.git

git remote get-url origin >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    git remote add origin %REPO_URL%
) else (
    git remote set-url origin %REPO_URL%
)
git branch -M main >nul 2>&1

echo A preparar ficheiros e alteracoes...
git add .
git commit -m "Update drivetrain tracker" >nul 2>&1

echo A enviar para o GitHub (main)...
git push -u origin main

echo.
echo ========================================================
echo Envio concluido!
echo A dashboard no Streamlit Cloud atualiza-se automaticamente:
echo https://fst16-drivetrain.streamlit.app
echo ========================================================
pause
