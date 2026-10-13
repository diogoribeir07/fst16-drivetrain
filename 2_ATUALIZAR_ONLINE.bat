@echo off
title FST 16 - Publicar / Atualizar no GitHub
echo ========================================================
echo       FST 16 - Publicar / Atualizar no GitHub
echo ========================================================
echo.

git remote get-url origin >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo Ainda nao associaste o teu repositorio do GitHub.
    echo.
    echo 1. Cria um repositorio PRIVADO em: https://github.com/new
    echo    (exemplo de nome: fst16-drivetrain)
    echo.
    set /p REPO_URL="2. Cola aqui o URL HTTPS do repositorio: "
    git remote add origin !REPO_URL!
    git branch -M main
)

echo.
echo A preparar ficheiros e alteracoes...
git add .
git commit -m "Update drivetrain tracker" >nul 2>&1

echo A enviar para o GitHub (main)...
git push -u origin main

echo.
echo ========================================================
echo Envio concluido!
echo A dashboard no Streamlit Cloud atualiza-se automaticamente.
echo ========================================================
pause
