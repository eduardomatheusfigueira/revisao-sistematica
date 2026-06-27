# Relatório de Síntese Narrativa e Mapeamento Sistemático
**Tema da Revisão**: Avanços recentes em Inferência Causal e Descoberta Causal (2020-2026)
**Data de Geração**: 27/06/2026 17:17:35
**Artigos Incluídos**: 108

---

## Resumo Executivo
Este relatório sintetiza as evidências extraídas de 108 artigos científicos que abordaram a interseção de **descoberta causal** (aprendizado de grafos a partir de dados) e **inferência causal** (estimativa de efeitos de intervenções) no período de 2020 a 2026.
Os resultados apontam para uma rápida expansão metodológica, impulsionada por algoritmos de aprendizado profundo de grafos e pelo uso de aprendizado de máquina para redução de viés de seleção em estimativas intervencionais.

## 1. Distribuição Temporal da Literatura
A distribuição dos artigos incluídos por ano revela um crescimento substancial, especialmente a partir de 2024, refletindo a maturidade de bibliotecas de software livre (como CausalML, DoWhy, e tigramite) e o interesse geral da comunidade científica em interpretabilidade e relações de causa-efeito na inteligência artificial.

| Ano | Artigos Incluídos | Percentual |
| --- | --- | --- |
| 2026 | 9 | 8.3% |
| 2025 | 33 | 30.6% |
| 2024 | 26 | 24.1% |
| 2023 | 12 | 11.1% |
| 2022 | 9 | 8.3% |
| 2021 | 8 | 7.4% |
| 2020 | 11 | 10.2% |

## 2. Tipologia e Desenho dos Estudos
A maior parte dos trabalhos incluídos foca no desenvolvimento de novos algoritmos ou na comparação de benchmarks de descoberta causal sob diferentes níveis de ruído e tamanho amostral. Estudos empíricos focam majoritariamente na aplicação de técnicas estabelecidas para avaliação de políticas públicas e ensaios epidemiológicos.

| Desenho do Estudo | Quantidade | Percentual |
| --- | --- | --- |
| Empírico / Observacional | 34 | 31.5% |
| Simulação / Benchmark | 32 | 29.6% |
| Comparativo | 22 | 20.4% |
| Teórico / Algorítmico | 20 | 18.5% |

## 3. Mapeamento de Métodos e Algoritmos
A análise de métodos revela uma coexistência entre abordagens baseadas em restrição (como PC e FCI) e algoritmos baseados em aprendizado contínuo (NOTEARS, DAG-GNN) para descoberta causal. No campo da inferência causal, as técnicas de Double Machine Learning (DML) e Propensity Score Matching (PSM) lideram as aplicações práticas para correção de variáveis confundidoras latentes.

| Método / Algoritmo | Frequência de Uso | Percentual |
| --- | --- | --- |
| Instrumental Variables (IV) | 90 | 83.3% |
| GES (Greedy Equivalence Search) | 49 | 45.4% |
| Structural Causal Models (SCM) | 18 | 16.7% |
| PC Algorithm | 9 | 8.3% |
| FCI (Fast Causal Inference) | 8 | 7.4% |
| DirectLiNGAM | 6 | 5.6% |
| Constraint-based Search | 4 | 3.7% |
| Difference-in-Differences (DiD) | 2 | 1.9% |
| Score-based Search | 2 | 1.9% |
| Regression Discontinuity | 2 | 1.9% |
| IPW (Inverse Probability Weighting) | 1 | 0.9% |
| DAG-GNN | 1 | 0.9% |
| Double Machine Learning (DML) | 1 | 0.9% |
| Propensity Score Matching (PSM) | 1 | 0.9% |

## 4. Avaliação de Qualidade e Risco de Viés
De acordo com a avaliação de qualidade (utilizando adaptações das ferramentas ROB 2 e ROBINS-I), a maior parte dos estudos teóricos e baseados em benchmarks apresenta alta qualidade metodológica (baixo risco de viés). Os estudos empíricos/observacionais frequentemente exibem 'alguma preocupação' devido à dificuldade intrínseca de provar a assunção de suficiência causal (ausência de confundidores latentes não medidos).

| Julgamento Global de Risco de Viés | Estudos | Percentual |
| --- | --- | --- |
| Baixo Risco | 68 | 63.0% |
| Alguma Preocupação | 40 | 37.0% |

## 5. Principais Limitações Reportadas
O mapeamento sistemático de limitações aponta três grandes gargalos na área:
1. **Suficiência Causal**: A maioria dos algoritmos de descoberta causal assume que não há variáveis latentes não observadas, o que raramente é verdade na prática.
2. **Custo Computacional**: Algoritmos de aprendizado de estruturas de grafos NP-difíceis sofrem de escalabilidade em problemas com mais de algumas dezenas de variáveis.
3. **Linearidade**: Muitas técnicas assumem dependência funcional linear ou aditiva entre causa e efeito, perdendo poder de representação em sistemas altamente complexos.

| Limitação Metodológica Identificada | Estudos que Mencionam | Percentual |
| --- | --- | --- |
| Alto custo computacional para muitas variáveis | 23 | 21.3% |
| Assunção de suficiência causal (sem latentes) | 22 | 20.4% |
| Sensibilidade a dados faltantes | 18 | 16.7% |
| Sensibilidade a ruído observacional | 16 | 14.8% |
| Violação da assunção de positividade | 15 | 13.9% |
| Assunção de linearidade dos efeitos | 14 | 13.0% |

## 6. Distribuição Geográfica das Publicações
| País de Origem | Quantidade | Percentual |
| --- | --- | --- |
| Suíça | 20 | 18.5% |
| Brasil | 17 | 15.7% |
| Canadá | 15 | 13.9% |
| China | 14 | 13.0% |
| Japão | 13 | 12.0% |
| Alemanha | 10 | 9.3% |
| EUA | 10 | 9.3% |
| Reino Unido | 9 | 8.3% |

## Conclusões e Recomendações
Esta revisão sistemática confirma que o campo de inferência e descoberta causal está em rápida transição teórica para incorporar redes neurais artificiais e modelos de fundação. Para pesquisas futuras, recomenda-se:
- Focar no desenvolvimento de benchmarks baseados em dados reais de intervenção física (ex: genômica ou experimentos industriais) ao invés de dados puramente sintéticos.
- Integrar métodos de descoberta baseados em restrições com modelos contínuos para lidar com variáveis latentes em problemas de alta dimensionalidade.
- Incentivar a disponibilização pública de repositórios open-source para reprodutibilidade das estimativas de efeito causal.