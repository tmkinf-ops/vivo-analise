@echo off
chcp 65001 >nul
title Auditoria Telecom

echo.
echo  =====================================================
echo   AuditoriaTel — Sistema de Auditoria de Contas
echo  =====================================================
echo.

:: Verificar Python
python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERRO] Python nao encontrado. Instale o Python 3.9 ou superior.
    pause
    exit /b 1
)

:: Instalar dependencias se necessario
if not exist ".deps_ok" (
    echo  Instalando dependencias, aguarde...
    python -m pip install -r requirements.txt --quiet
    if errorlevel 1 (
        echo  [ERRO] Falha ao instalar dependencias.
        pause
        exit /b 1
    )
    echo OK > .deps_ok
    echo  Dependencias instaladas com sucesso!
    echo.
)

:: Iniciar servidor
echo  Iniciando servidor em http://localhost:5000
echo  Pressione Ctrl+C para encerrar.
echo.
start "" http://localhost:5000
python app.py

pause
