@echo off
title FST 16 - Drivetrain Tracker
echo ========================================================
echo          FST 16 - Drivetrain Tracker
echo ========================================================
echo.
echo A iniciar a aplicacao...
echo Vai abrir automaticamente no teu browser (http://localhost:8501).
echo.
echo Para fechar a aplicacao, basta fechar esta janela.
echo ========================================================
echo.
streamlit run app.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Ocorreu um problema a iniciar o Streamlit. A verificar dependencias...
    pip install -r requirements.txt
    streamlit run app.py
)
pause
