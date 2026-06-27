@echo off
chcp 65001 > nul
title BDTD Harvester - Inicializador Rápido
echo =================================================================
echo             INICIALIZADOR DO BDTD HARVESTER
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
echo DICA: Para buscas com mais de uma palavra contendo aspas, use barras invertidas para escapar:
echo Exemplo: \"inferência causal\"
echo Exemplo: saneamento AND \"abastecimento de água\"
echo.
set /p search_query="Digite o termo de busca para a BDTD (pressione ENTER para usar a busca padrão do config.json): "

if "%search_query%"=="" (
    echo Executando harvester com a configuração padrão do config.json...
    python bdtd_harvester.py
) else (
    echo Executando harvester com a busca: "%search_query%"
    python bdtd_harvester.py --query "%search_query%"
)

echo.
echo =================================================================
echo Processo concluído! Os resultados estarão na pasta configurada.
echo =================================================================
pause
