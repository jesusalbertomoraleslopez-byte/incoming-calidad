@echo off
title Iniciar Sistema de Calidad - SIGRAMA
echo =======================================================================
echo     INICIANDO SISTEMA DE CONTROL DE CALIDAD EN RECEPCION - SIGRAMA
echo =======================================================================
echo.
cd /d "%~dp0"
py -m streamlit run app.py --server.port 8505
if %ERRORLEVEL% neq 0 (
    echo.
    echo [ADVERTENCIA] No se pudo iniciar usando 'py -m'. Intentando con 'python -m'...
    python -m streamlit run app.py --server.port 8505
)
if %ERRORLEVEL% neq 0 (
    echo.
    echo [ADVERTENCIA] Tampoco se pudo iniciar usando 'python -m'. Intentando con 'streamlit run'...
    streamlit run app.py --server.port 8505
)
if %ERRORLEVEL% neq 0 (
    echo.
    echo [ERROR] No se pudo iniciar la aplicacion. Verifique que Streamlit este instalado.
    echo Intente ejecutar: pip install streamlit pandas openpyxl reportlab matplotlib pypdf pillow
    pause
)
