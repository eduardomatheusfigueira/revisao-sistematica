# Procedimento Completo de Exemplo — Revisão Sistemática

## Tema de Pesquisa

**Pergunta de pesquisa (PICo):**
> "Quais são os principais métodos e aplicações de inferência causal e descoberta causal reportados na literatura científica recente (2020–2026)?"

- **P (População/Contexto):** Literatura científica recente (2020–2026)
- **I (Fenômeno de Interesse):** Métodos de *descoberta causal* (aprendizado de grafos causais a partir de dados observacionais) e *inferência causal* (estimativa de efeitos causais de intervenções)
- **Co (Contexto):** Avanços conceituais, algoritmos comuns e áreas de aplicação dominantes

---

## Estrutura do Diretório

```
procedimento completo de exemplo/
│
├── README.md                              ← Este arquivo
├── Formulario_Desenho_Pesquisa_RSL.xlsx   ← Formulário de desenho pré-configurado
│
├── gerar_formulario_excel.py              ← Gerador da planilha de desenho
├── exportar_configs_do_excel.py           ← Exportador de configs JSON a partir do Excel
│
├── config.json                            ← Configuração BDTD
├── config_openalex.json                   ← Configuração OpenAlex
├── config_scielo.json                     ← Configuração SciELO
├── config_scopus.json                     ← Configuração Scopus
│
├── bdtd_harvester.py                      ← Coletor da BDTD
├── openalex_harvester.py                  ← Coletor do OpenAlex
├── scielo_harvester.py                    ← Coletor do SciELO
├── scopus_harvester.py                    ← Coletor do Scopus
│
├── run.bat / run_openalex.bat / ...       ← Scripts de execução (Windows)
├── requirements.txt                       ← Dependências Python
│
├── openalex_outputs/                      ← Dados coletados do OpenAlex
│   ├── OpenAlex_Data_Export.xlsx           ← Planilha formatada
│   ├── openalex_clean_data.csv            ← CSV limpo
│   ├── openalex_raw_backup.json           ← Backup cru completo (JSON)
│   ├── openalex_summary_report.md         ← Relatório analítico
│   └── openalex_harvester.log             ← Log de execução
│
├── scielo_outputs/                        ← Dados coletados do SciELO
│   ├── SciELO_Data_Export.xlsx
│   ├── scielo_clean_data.csv
│   ├── scielo_raw_backup.json
│   ├── scielo_summary_report.md
│   └── scielo_harvester.log
│
└── data_outputs/                          ← Dados da BDTD (bloqueada nesta execução)
    ├── BDTD_Data_Export.xlsx
    ├── bdtd_clean_data.csv
    ├── bdtd_raw_backup.json
    ├── bdtd_summary_report.md
    └── bdtd_harvester.log
```

---

## Resultados da Coleta (27/06/2026)

### Expressão Booleana de Busca
```
("inferência causal" OR "causal inference") AND ("descoberta causal" OR "causal discovery")
```

### Recorte Temporal
2020–2026

### Bases Consultadas e Resultados

| Base        | Registros | Status |
|-------------|-----------|--------|
| **OpenAlex**| **656**   | ✅ Sucesso |
| **SciELO**  | **1**     | ✅ Sucesso |
| **BDTD**    | 0         | ❌ HTTP 403 (acesso bloqueado pela API) |
| **Scopus**  | —         | ⏸ Não executado (requer API Key institucional) |

### Destaques do OpenAlex (656 artigos)

**Distribuição temporal:**
- 2025–2026 concentram **57%** da produção (374 artigos)
- Crescimento exponencial a partir de 2022

**Artigos mais citados:**
1. *"Causal inference for time series"* — Nature Reviews (255 citações)
2. *"Causal machine learning for healthcare"* — Royal Society Open Science (195 citações)
3. *"Methods and tools for causal discovery and causal inference"* — WIREs (192 citações)

**Principais fontes:**
- arXiv (22.6%), DOAJ (9.6%), PubMed (9.0%), Zenodo (9.0%)

**Pesquisadores mais prolíficos:**
- Jakob Runge (14 artigos), Urmi Ninad (7), Amir M. Rahmani (7)

---

## Como Reproduzir

### 1. Instalar dependências
```bash
pip install -r requirements.txt
```

### 2. (Opcional) Ajustar o formulário de desenho
Abra `Formulario_Desenho_Pesquisa_RSL.xlsx` e modifique os parâmetros desejados.

### 3. Exportar configurações
```bash
python exportar_configs_do_excel.py
```

### 4. Executar coletores
```bash
python openalex_harvester.py --config config_openalex.json
python scielo_harvester.py --config config_scielo.json
python bdtd_harvester.py --config config.json
```
Ou use os scripts `.bat` (Windows):
```bash
run_openalex.bat
run_scielo.bat
run.bat
```

### 5. Próximos Passos (Triagem)
Os dados coletados devem ser consolidados na **Planilha de Triagem** do formulário de desenho (`Formulario_Desenho_Pesquisa_RSL.xlsx`, aba "Planilha Triagem") para aplicação dos critérios de elegibilidade.

---

## Critérios de Elegibilidade

### Inclusão
1. Artigos que abordem explicitamente métodos de descoberta causal **e** inferência causal
2. Publicações a partir de 2020
3. Idiomas: Inglês, Português ou Espanhol

### Exclusão
1. Artigos de opinião, cartas, revisões informais ou notas técnicas
2. Estudos que usem os termos em contextos puramente metafóricos ou sem fundamentação estatística/computacional

---

## Notas

- A **BDTD** retornou HTTP 403 durante a coleta. Isso é comum: a API pública da BDTD aplica restrições por User-Agent e por taxa de requisições. Uma alternativa é acessar o portal web diretamente.
- O **Scopus** requer uma API Key institucional (Elsevier Developer Portal). Configure-a no campo correspondente do formulário Excel antes de executar.
- Os dados de saída já estão nos formatos CSV, JSON e Excel (.xlsx) para fácil importação em qualquer ferramenta de análise.
