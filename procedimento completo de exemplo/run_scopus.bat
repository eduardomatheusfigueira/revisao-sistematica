@echo off
rem -*- coding: utf-8 -*-
chcp 65001 > nul
title Scopus Harvester - Inicializador Rápido
echo =================================================================
echo             INICIALIZADOR DO SCOPUS HARVESTER
echo =================================================================
echo.
echo Verificando se o Python está disponível...
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERRO] Python não encontrado no PATH do sistema.
    echo Por favor, instale o Python e adicione-o ao seu PATH.
    pause
    exit /b 1
)

echo [OK] Python detectado.
echo.
echo DICA: Para buscas exatas na Scopus, use a sintaxe TITLE-ABS-KEY("termo"):
echo Exemplo: TITLE-ABS-KEY("inferencia causal")
echo Exemplo: TITLE-ABS-KEY("causal inference")
echo.
set /p search_query="Digite o termo de busca para a Scopus (pressione ENTER para usar a busca padrão do config_scopus.json): "

if "%search_query%"=="" (
    echo Executando harvester com a configuração padrão do config_scopus.json...
    python scopus_harvester.py
) else (
    echo Executando harvester com a busca: "%search_query%"
    python scopus_harvester.py --query "%search_query%"
)

echo.
echo =================================================================
echo Processo concluído! Os resultados estarão na pasta configurada.
echo =================================================================
pause
