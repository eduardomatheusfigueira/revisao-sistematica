# BDTD Scientific Data Harvester CLI

O **BDTD Scientific Data Harvester** é uma ferramenta de linha de comando (CLI) em Python de código aberto para extração automatizada, limpeza e exportação de metadados acadêmicos da **BDTD (Biblioteca Digital Brasileira de Teses e Dissertações)**, coordenada pelo IBICT.

Este software foi desenhado do zero baseando-se nas melhores práticas de engenharia de raspagem de dados (resiliência de conexão, *exponential backoff*, polidez de requisição) e realiza o mapeamento completo do **Padrão Brasileiro de Metadados de Teses e Dissertações (MTD3-BR - Dublin Core)**.

---

## Recursos Principais

1. **Acesso Otimizado (API-First)**: Consulta direta à API REST pública do VuFind da BDTD (`/vufind/api/v1/search`), garantindo rapidez, baixo tráfego e evitando bloqueios de WAF ou CAPTCHAs.
2. **Resiliência a Falhas**: Implementa re-tentativas com política de *Exponential Backoff* em caso de falha de conexão, erro no servidor (500) ou limite de taxa (429).
3. **Coleta de Metadados Completos (MTD3-BR)**: Extrai 23 campos contendo dados ricos como **Orientadores, Coorientadores, Membros da Banca (Referees), Programa de Pós-Graduação, Departamento, Palavras-chave, Agência de Fomento, Idiomas e Resumos (Abstracts) limpos**.
4. **Múltiplos Formatos de Saída**: Salva automaticamente os resultados em JSON (backup cru), CSV (dataset rico para análise de dados) e uma planilha Excel (`.xlsx`) com design de visualização profissional e tabela de sumário.

---

## Instalação

### Pré-requisitos
* Python 3.8 ou superior instalado.
* Acesso à internet.

### Passos de Instalação
1. Abra o terminal (Prompt de Comando ou PowerShell) na pasta do projeto.
2. Instale as dependências executando:
   ```bash
   pip install -r requirements.txt
   ```

---

## Como Executar

O software oferece três formas flexíveis de inicialização e configuração:

### 1. Execução Rápida no Windows (run.bat)
Basta dar um duplo-clique no arquivo `run.bat` (ou executá-lo via terminal). O inicializador batch perguntará interativamente qual termo deseja buscar na BDTD e disparará a coleta automaticamente.

### 2. Uso via CLI (Interface de Linha de Comando)
Você pode parametrizar toda a coleta diretamente por argumentos de terminal:
```bash
# Exemplo 1: Buscar um termo simples e limitar o recorte temporal
python bdtd_harvester.py --query "reuso de água" --start-year 2018 --end-year 2026

# Exemplo 2: Buscar com expressão booleana complexa (escape as aspas duplas internas)
python bdtd_harvester.py --query "saneamento AND \"inferência causal\"" --output-dir minhas_saidas

# Exemplo 3: Buscar com limite personalizado de página de resultados
python bdtd_harvester.py --query "saneamento" --limit 100 --output-dir saidas_rapidas
```

### 3. Configuração via Arquivo JSON (config.json)
Se executado sem argumentos de linha de comando (`python bdtd_harvester.py`), o software carregará as definições padrão de busca e caminhos a partir do arquivo `config.json` local.

---

## Argumentos de Linha de Comando Disponíveis

* `--query`: Expressão de busca (e.g. `saneamento AND "inferência causal"`).
* `--start-year`: Limite de início do ano de defesa (padrão `2000`).
* `--end-year`: Limite de fim do ano de defesa (padrão `2026`).
* `--output-dir`: Diretório onde as saídas serão gravadas (padrão `data_outputs`).
* `--limit`: Quantidade de registros por página de requisição (padrão `50`).
* `--config`: Caminho para um arquivo JSON alternativo de configurações (padrão `config.json`).
* `-h`, `--help`: Exibe a tela de ajuda com todos os comandos disponíveis.

---

## Arquivos Gerados (`data_outputs/`)

O diretório de saídas conterá os seguintes arquivos estruturados:

* **`bdtd_raw_backup.json`**: Cópia dos objetos brutos Solr retornados pela API (importante para integridade de dados e auditoria).
* **`bdtd_clean_data.csv`**: Tabela organizada com 23 colunas, ideal para manipulação em Pandas, R ou softwares estatísticos. Os resumos (abstracts) são limpos de quebras de linha e tabulações.
* **`BDTD_Data_Export.xlsx`**: Planilha Excel contendo:
  * Aba *BDTD Cleaned Data*: Tabela estilizada com Segoe UI, cabeçalho de destaque e linhas em padrão zebra para fácil leitura.
  * Aba *Extraction Summary*: Dados resumidos com a contagem de tipos de trabalho (Tese/Dissertação) e data da execução.
* **`bdtd_summary_report.md`**: Relatório em formato Markdown com estatísticas descritivas (Top Universidades, Palavras-Chave, Idiomas, Orientadores).
* **`bdtd_harvester.log`**: Histórico detalhado de depuração e execuções anteriores do software.
