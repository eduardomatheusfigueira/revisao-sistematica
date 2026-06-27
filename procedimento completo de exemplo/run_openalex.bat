@echo off
rem -*- coding: utf-8 -*-
chcp 65001 > nul
title OpenAlex Harvester - Inicializador Rápido
echo =================================================================
echo             INICIALIZADOR DO OPENALEX HARVESTER
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
echo DICA: Para buscas exatas, use aspas duplas no termo:
echo Exemplo: "inferencia causal"
echo.
set /p search_query="Digite o termo de busca para o OpenAlex (pressione ENTER para usar a busca padrão do config_openalex.json): "

if "%search_query%"=="" (
    echo Executando harvester com a configuração padrão do config_openalex.json...
    python openalex_harvester.py
) else (
    echo Executando harvester com a busca: "%search_query%"
    python openalex_harvester.py --query "%search_query%"
)

echo.
echo =================================================================
echo Processo concluído! Os resultados estarão na pasta configurada.
echo =================================================================
pause
