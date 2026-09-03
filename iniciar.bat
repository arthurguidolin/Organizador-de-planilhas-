@echo off
chcp 65001 > nul
title Organizador Inteligente de Planilhas

echo ========================================================
echo    ORGANIZADOR INTELIGENTE DE PLANILHAS (DESKTOP)
echo ========================================================
echo.

:: Verifica se o Python esta instalado
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERRO] Python nao encontrado no sistema!
    echo Por favor, instale o Python 3.10 ou superior e marque a opcao "Add Python to PATH".
    echo Download: https://www.python.org/downloads/
    echo.
    pause
    exit /b
)

:: Verifica/instala dependencias
echo Verificando dependencias necessarias...
python -m pip install -r requirements.txt --quiet --disable-pip-version-check
if %errorlevel% neq 0 (
    echo [AVISO] Falha ao verificar dependencias automaticamente. Tentando iniciar o app...
)

echo.
echo Iniciando o aplicativo...
echo.

start "" pythonw main.py
if %errorlevel% neq 0 (
    python main.py
)

exit
