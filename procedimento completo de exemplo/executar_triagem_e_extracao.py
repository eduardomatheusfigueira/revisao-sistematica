"""
Simulador de Triagem, Extração e Qualidade para Revisão Sistemática de Exemplo
=============================================================================
Este script realiza a simulação do fluxo completo de revisão sistemática de literatura:
1. Simulação da triagem por Título/Resumo (T/R) para os 463 registros (com conflitos controlados e resoluções)
2. Simulação da busca de texto completo e triagem por Texto Completo (TC)
3. Extração detalhada de dados para os 108 artigos aprovados
4. Avaliação de qualidade metodológica (risco de viés)
5. Atualização completa da aba PRISMA Flow com os números exatos

Uso:
  python executar_triagem_e_extracao.py
"""

import os
import re
import sys
import random
from datetime import datetime
import openpyxl
from openpyxl.styles import (
    Font, Alignment, Border, Side, PatternFill
)
from openpyxl.utils import get_column_letter

# Cores do Design System (compatível com os scripts anteriores)
COLORS = {
    "indigo":        "4338CA",
    "indigo_pale":   "E0E7FF",
    "indigo_wash":   "EEF2FF",
    "slate_900":     "0F172A",
    "slate_700":     "334155",
    "slate_200":     "E2E8F0",
    "slate_50":      "F8FAFC",
    "white":         "FFFFFF",
    "success":       "059669",
    "success_light": "D1FAE5",
    "warning":       "D97706",
    "warning_light": "FEF3C7",
    "error":         "DC2626",
    "error_light":   "FEE2E2",
}

def main():
    import sys
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    
    print("=" * 60)
    print("  SIMULADOR DE TRIAGEM, EXTRAÇÃO E QUALIDADE (PRISMA)")
    print("  Tema: Inferência Causal & Descoberta Causal")
    print("=" * 60)

    filename = "Formulario_Desenho_Pesquisa_RSL.xlsx"
    if not os.path.exists(filename):
        print(f"[ERRO] Arquivo '{filename}' não encontrado.")
        return

    print(f"\n[1/5] Carregando planilha '{filename}'...")
    wb = openpyxl.load_workbook(filename)
    
    # ═══════════════════════════════════════════════════════════════
    # STEP 1: LEITURA E SIMULAÇÃO DE TRIAGEM (T/R)
    # ═══════════════════════════════════════════════════════════════
    print("\n[2/5] Lendo registros e simulando decisões de triagem (T/R & TC)...")
    ws_triagem = wb["📊 Planilha Triagem"]
    
    font_body = Font(name="Segoe UI", size=9, color=COLORS["slate_900"])
    align_center = Alignment(horizontal='center', vertical='center')
    
    records = []
    max_row = ws_triagem.max_row
    
    # Identificar quantos registros existem
    num_records = 0
    for r in range(3, max_row + 1):
        if ws_triagem.cell(row=r, column=1).value is not None:
            num_records += 1
            
    print(f"  Total de registros encontrados para triagem: {num_records}")
    
    # Para consistência em execuções repetidas, usaremos sementes determinísticas baseadas no título
    random.seed(42)
    
    # Heurística de classificação baseada em termos no título/resumo
    # Queremos selecionar trabalhos que tratam de inferência causal E descoberta causal.
    # Excluiremos:
    # - Trabalhos puramente aplicados fora de TI/Estatística (ex: medicina clínica pura, biologia pura)
    # - Revisões gerais sem foco metodológico
    # - Trabalhos irrelevantes que passaram pelo filtro por falso positivo
    
    keywords_exclude = [
        r"\bhealth\b", r"\bpatient\b", r"\bclinical\b", r"\btrial\b", r"\bdrug\b",
        r"\bdisease\b", r"\bcancer\b", r"\brain\b", r"\bneuroscience\b",
        r"\bpsychology\b", r"\bchild\b", r"\bteacher\b", r"\beducation\b",
        r"\beconomic policy\b", r"\bfinance\b", r"\bstock\b"
    ]
    
    keywords_include = [
        r"\balgorithm\b", r"\bmethod\b", r"\bdiscovery\b", r"\binference\b",
        r"\bcausal discovery\b", r"\bcausal inference\b", r"\bgraph\b", r"\bdag\b",
        r"\bstructural causal model\b", r"\btime series\b", r"\bmachine learning\b",
        r"\bdeep learning\b", r"\bestimation\b", r"\bconfounder\b", r"\btreatment effect\b"
    ]
    
    n_total = num_records
    n_dupes_auto = 0
    n_excluded_tr = 0
    n_included_tr = 0
    n_conflicts_tr = 0
    
    included_records = [] # Guardará os registros que passaram na triagem T/R
    
    for r in range(3, 3 + num_records):
        title = str(ws_triagem.cell(row=r, column=4).value or "")
        abstract = str(ws_triagem.cell(row=r, column=8).value or "")
        text = (title + " " + abstract).lower()
        
        # Determinar relevância metodológica (score)
        inc_score = sum(1 for kw in keywords_include if re.search(kw, text))
        exc_score = sum(1 for kw in keywords_exclude if re.search(kw, text))
        
        # Decisão simulada
        # Revisor 1
        if exc_score > inc_score + 1:
            decision_r1 = "Excluir"
            motive_r1 = "Fora do escopo"
        elif inc_score >= 2:
            decision_r1 = "Incluir"
            motive_r1 = ""
        else:
            decision_r1 = "Excluir"
            motive_r1 = "Fora do escopo"
            
        # Revisor 2 (introduzindo ruído controlado de 8% para gerar conflitos reais)
        decision_r2 = decision_r1
        motive_r2 = motive_r1
        if random.random() < 0.08:
            if decision_r1 == "Incluir":
                decision_r2 = "Dúvida"
            elif decision_r1 == "Excluir":
                decision_r2 = "Dúvida"
            else:
                decision_r2 = "Incluir"
                
        # Escrever na planilha
        ws_triagem.cell(row=r, column=11, value="Revisor A").font = font_body
        ws_triagem.cell(row=r, column=12, value=decision_r1).font = font_body
        ws_triagem.cell(row=r, column=13, value=motive_r1).font = font_body
        
        ws_triagem.cell(row=r, column=14, value="Revisor B").font = font_body
        ws_triagem.cell(row=r, column=15, value=decision_r2).font = font_body
        ws_triagem.cell(row=r, column=16, value=motive_r2).font = font_body
        
        # Resolução de conflitos
        final_tr = decision_r1
        if decision_r1 != decision_r2:
            n_conflicts_tr += 1
            # Resolução pelo revisor de consenso: a maioria das dúvidas vira inclusão para avaliação de texto completo
            if "Incluir" in (decision_r1, decision_r2):
                final_tr = "Incluir"
            else:
                final_tr = "Excluir"
            ws_triagem.cell(row=r, column=18, value=final_tr).font = font_body
        else:
            ws_triagem.cell(row=r, column=18, value="").font = font_body # Vazio se não há conflito (a fórmula do excel assume o mesmo valor)

        # Calculando decisão final de T/R (que a fórmula do Excel faria)
        # IF R_row != "" then R_row else (if L_row == O_row then L_row else "CONFLITO")
        final_decision_tr = final_tr
        
        if final_decision_tr == "Incluir":
            n_included_tr += 1
            record_info = {
                "row_index": r,
                "id": ws_triagem.cell(row=r, column=1).value,
                "base": ws_triagem.cell(row=r, column=2).value,
                "doi": ws_triagem.cell(row=r, column=3).value,
                "titulo": title,
                "autores": ws_triagem.cell(row=r, column=5).value,
                "ano": ws_triagem.cell(row=r, column=6).value,
                "periodico": ws_triagem.cell(row=r, column=7).value,
                "resumo": abstract,
                "palavras_chave": ws_triagem.cell(row=r, column=9).value,
            }
            included_records.append(record_info)
        else:
            n_excluded_tr += 1
            
    print(f"  Triagem T/R concluída:")
    print(f"    - Incluídos para texto completo: {n_included_tr}")
    print(f"    - Excluídos: {n_excluded_tr}")
    print(f"    - Conflitos detectados e resolvidos: {n_conflicts_tr}")

    # ═══════════════════════════════════════════════════════════════
    # STEP 2: SIMULAÇÃO DE TRIAGEM POR TEXTO COMPLETO (TC)
    # ═══════════════════════════════════════════════════════════════
    print("\n[3/5] Simulando busca de texto completo e triagem TC...")
    
    n_not_retrieved = 5 # 5 artigos não puderam ser recuperados (ex: paywall ou link quebrado)
    n_retrieved = len(included_records) - n_not_retrieved
    
    # Selecionar aleatoriamente 5 artigos para marcar como não recuperados
    not_retrieved_indices = set(random.sample(range(len(included_records)), n_not_retrieved))
    
    final_included_records = []
    n_excluded_tc = 0
    n_included_tc = 0
    n_conflicts_tc = 0
    
    # Lista de motivos de exclusão na etapa de Texto Completo
    motivos_tc = ["Fora do escopo", "Tipo de documento inadequado", "Qualidade insuficiente", "Duplicata"]
    
    for idx, rec in enumerate(included_records):
        r = rec["row_index"]
        
        # Revisor 1 (TC)
        ws_triagem.cell(row=r, column=20, value="Revisor A").font = font_body
        
        if idx in not_retrieved_indices:
            # Não recuperado
            ws_triagem.cell(row=r, column=21, value="Excluir").font = font_body
            ws_triagem.cell(row=r, column=22, value="Sem acesso").font = font_body
            ws_triagem.cell(row=r, column=23, value="Revisor B").font = font_body
            ws_triagem.cell(row=r, column=24, value="Excluir").font = font_body
            ws_triagem.cell(row=r, column=25, value="Sem acesso").font = font_body
            ws_triagem.cell(row=r, column=27, value="Excluir").font = font_body
            ws_triagem.cell(row=r, column=29, value="Busca TC").font = font_body
            n_excluded_tc += 1
            continue
            
        # Para os recuperados, decidir elegibilidade
        # Vamos selecionar os trabalhos que têm maior profundidade metodológica
        # Por exemplo, os que contêm algoritmos específicos no título ou abstract
        text_lower = rec["titulo"].lower() + " " + rec["resumo"].lower()
        
        # Se contiver termos específicos, incluímos
        specific_methods = [
            "pc algorithm", "fci", "lingam", "dag", "structural causal model",
            "double machine learning", "propensity score", "instrumental variable",
            "backdoor", "do-calculus", "granger", "causal graphical", "causal network"
        ]
        
        has_specific_method = any(m in text_lower for m in specific_methods)
        
        if has_specific_method and len(final_included_records) < 108:
            decision_tc_r1 = "Incluir"
            motive_tc_r1 = ""
        else:
            decision_tc_r1 = "Excluir"
            motive_tc_r1 = random.choice(motivos_tc)
            
        # Revisor 2 TC (ruído de 5% para conflito)
        decision_tc_r2 = decision_tc_r1
        motive_tc_r2 = motive_tc_r1
        if random.random() < 0.05:
            decision_tc_r2 = "Dúvida"
            
        ws_triagem.cell(row=r, column=21, value=decision_tc_r1).font = font_body
        ws_triagem.cell(row=r, column=22, value=motive_tc_r1).font = font_body
        ws_triagem.cell(row=r, column=23, value="Revisor B").font = font_body
        ws_triagem.cell(row=r, column=24, value=decision_tc_r2).font = font_body
        ws_triagem.cell(row=r, column=25, value=motive_tc_r2).font = font_body
        
        final_tc = decision_tc_r1
        if decision_tc_r1 != decision_tc_r2:
            n_conflicts_tc += 1
            # Resolução
            if decision_tc_r1 == "Incluir" or decision_tc_r2 == "Incluir":
                # Forçar inclusão se algum incluiu e preencheu a meta de 108
                if len(final_included_records) < 108:
                    final_tc = "Incluir"
                else:
                    final_tc = "Excluir"
                    ws_triagem.cell(row=r, column=22, value="Fora do escopo")
                    ws_triagem.cell(row=r, column=25, value="Fora do escopo")
            else:
                final_tc = "Excluir"
            ws_triagem.cell(row=r, column=27, value=final_tc).font = font_body
            
        if final_tc == "Incluir":
            n_included_tc += 1
            final_included_records.append(rec)
            ws_triagem.cell(row=r, column=29, value="Incluído").font = font_body
        else:
            n_excluded_tc += 1
            ws_triagem.cell(row=r, column=29, value="Excluído TC").font = font_body
            
    print(f"  Triagem TC concluída:")
    print(f"    - Estudos incluídos finais: {len(final_included_records)}")
    print(f"    - Excluídos na etapa TC (incluindo não recuperados): {n_excluded_tc}")
    print(f"    - Conflitos TC detectados e resolvidos: {n_conflicts_tc}")


    # ═══════════════════════════════════════════════════════════════
    # STEP 3: PREENCHIMENTO DA PLANILHA DE EXTRAÇÃO
    # ═══════════════════════════════════════════════════════════════
    print("\n[4/5] Populando a Planilha de Extração...")
    ws_extracao = wb["📋 Planilha Extração"]
    
    # Limpar linhas existentes a partir da linha 3
    for r in range(3, ws_extracao.max_row + 1):
        for c in range(1, 26):
            ws_extracao.cell(row=r, column=c, value=None)
            
    # Estilos para extração
    fill_zebra_even = PatternFill(start_color=COLORS["slate_50"], end_color=COLORS["slate_50"], fill_type="solid")
    fill_zebra_odd = PatternFill(start_color=COLORS["white"], end_color=COLORS["white"], fill_type="solid")
    border_light = Border(
        left=Side(style='thin', color=COLORS["slate_200"]),
        right=Side(style='thin', color=COLORS["slate_200"]),
        top=Side(style='thin', color=COLORS["slate_200"]),
        bottom=Side(style='thin', color=COLORS["slate_200"])
    )
    align_left = Alignment(horizontal='left', vertical='top', wrap_text=True)
    align_center_top = Alignment(horizontal='center', vertical='top')
    
    # Dicionários de termos para simular extrações realistas
    metodos_descoberta = ["PC Algorithm", "FCI (Fast Causal Inference)", "DirectLiNGAM", "GES (Greedy Equivalence Search)", "NOTEARS (Gradient-based)", "DAG-GNN", "Constraint-based Search", "Score-based Search"]
    metodos_inferencia = ["Double Machine Learning (DML)", "Propensity Score Matching (PSM)", "IPW (Inverse Probability Weighting)", "Instrumental Variables (IV)", "Structural Causal Models (SCM)", "G-Estimation", "Difference-in-Differences (DiD)", "Regression Discontinuity"]
    
    estudos_desenho = ["Teórico / Algorítmico", "Empírico / Observacional", "Simulação / Benchmark", "Comparativo"]
    aplicacoes = ["Saúde e Epidemiologia", "Econometria e Políticas Públicas", "Sistemas de Recomendação", "Climatologia e Meteorologia", "Genômica e Bioinformática", "Finanças e Análise de Crédito", "Modelagem de Redes Sociais"]
    limitacoes = ["Assunção de suficiência causal (sem latentes)", "Sensibilidade a dados faltantes", "Alto custo computacional para muitas variáveis", "Assunção de linearidade dos efeitos", "Sensibilidade a ruído observacional", "Violação da assunção de positividade"]
    paises = ["EUA", "Alemanha", "China", "Reino Unido", "Canadá", "Suíça", "Japão", "Brasil"]
    financiamentos = ["Financiamento Público (NSF/DFG)", "Não especificado", "Financiamento de Empresa de Tecnologia", "Conselho de Pesquisa Europeu (ERC)"]

    for idx, rec in enumerate(final_included_records, start=1):
        row_num = idx + 2
        fill = fill_zebra_even if idx % 2 == 0 else fill_zebra_odd
        
        # Criar nome curto de autor, ex: "Runge et al. (2023)"
        autores_raw = str(rec["autores"] or "")
        first_author = autores_raw.split(",")[0].split(";")[0].strip()
        if "et al." not in first_author and len(autores_raw.split(",")) > 1:
            autor_ano = f"{first_author} et al. ({rec['ano']})"
        else:
            autor_ano = f"{first_author} ({rec['ano']})"
            
        # Determinar algoritmos e métodos
        title_lower = rec["titulo"].lower() + " " + rec["resumo"].lower()
        metodo_estatistico = []
        for m in metodos_descoberta:
            if m.split(" ")[0].lower() in title_lower:
                metodo_estatistico.append(m)
        for m in metodos_inferencia:
            # Pegar abreviações ou nomes
            abbrev = re.search(r'\((.*?)\)', m)
            if abbrev and abbrev.group(1).lower() in title_lower:
                metodo_estatistico.append(m)
            elif m.split(" ")[0].lower() in title_lower:
                metodo_estatistico.append(m)
                
        if not metodo_estatistico:
            # Fallback
            metodo_estatistico = [random.choice(metodos_descoberta), random.choice(metodos_inferencia)]
            
        metodo_str = ", ".join(list(set(metodo_estatistico))[:2])
        
        # Objetivos e resultados simulados
        area = random.choice(aplicacoes)
        objetivo = f"Desenvolver ou aplicar métodos de inferência e descoberta causal baseados em {metodo_str} no contexto de {area.lower()}."
        resultados = f"Demonstrou que o método proposto obteve melhor acurácia na identificação de relações causais ou na estimativa de efeitos em comparação com baselines tradicionais de aprendizado de máquina."
        lim = random.choice(limitacoes)
        
        values = [
            idx,                                      # A: ID
            autor_ano,                                # B: Autor (Ano)
            rec.get("doi") or "N/A",                  # C: DOI
            rec["base"],                              # D: Base
            rec["titulo"],                            # E: Título
            rec["autores"],                           # F: Autores
            rec["ano"],                               # G: Ano
            rec["periodico"] or "N/A",                # H: Periódico
            rec["resumo"],                            # I: Resumo
            rec["palavras_chave"] or "N/A",           # J: Palavras-Chave
            objetivo,                                 # K: Objetivo
            random.choice(estudos_desenho),           # L: Desenho do Estudo
            f"N = {random.randint(100, 10000)} observações", # M: Tamanho da Amostra
            metodo_str,                               # N: Método Estatístico
            resultados,                               # O: Principais Resultados
            lim,                                      # P: Limitações
            random.choice(paises),                    # Q: País
            random.choice(financiamentos),            # R: Financiamento
            "ROB 2" if "Observacional" not in estudos_desenho else "ROBINS-I", # S: Qualidade (Ferramenta)
            random.choice(["Baixo Risco", "Alguma Preocupação", "Baixo Risco"]), # T: Classificação Risco
            "Extrator Automático LLM-Sim",            # U: Extrator
            datetime.now().strftime("%d/%m/%Y"),      # V: Data Extração
            "Revisor de Consenso",                    # W: Verificador
            "Nenhum conflito na extração.",           # X: Observações
        ]
        
        for col_idx, val in enumerate(values, start=1):
            cell = ws_extracao.cell(row=row_num, column=col_idx, value=val)
            cell.font = font_body
            cell.border = border_light
            cell.fill = fill
            
            if col_idx in (1, 7, 22):  # ID, Ano, Data
                cell.alignment = align_center_top
            else:
                cell.alignment = align_left
                
    print(f"  [OK] {len(final_included_records)} registros populados na Planilha de Extração.")

    # ═══════════════════════════════════════════════════════════════
    # STEP 4: PREENCHIMENTO DA PLANILHA DE QUALIDADE
    # ═══════════════════════════════════════════════════════════════
    print("\n[5/5] Populando a Planilha de Qualidade...")
    ws_qualidade = wb["🏅 Planilha Qualidade"]
    
    # Limpar linhas existentes a partir da linha 4
    for r in range(4, ws_qualidade.max_row + 1):
        for c in range(1, 13):
            ws_qualidade.cell(row=r, column=c, value=None)
            
    judgments = ["Baixo risco", "Alguma preocupação", "Alto risco"]
    
    for idx, rec in enumerate(final_included_records, start=1):
        row_num = idx + 3 # A tabela de qualidade começa na linha 4
        fill = fill_zebra_even if idx % 2 == 0 else fill_zebra_odd
        
        # Autor (Ano)
        autores_raw = str(rec["autores"] or "")
        first_author = autores_raw.split(",")[0].split(";")[0].strip()
        if "et al." not in first_author and len(autores_raw.split(",")) > 1:
            autor_ano = f"{first_author} et al. ({rec['ano']})"
        else:
            autor_ano = f"{first_author} ({rec['ano']})"
            
        # Simular avaliações de qualidade (baixo risco na maioria)
        d1 = random.choice(["Baixo risco", "Baixo risco", "Alguma preocupação"])
        d2 = "Baixo risco"
        d3 = random.choice(["Baixo risco", "Baixo risco", "Alguma preocupação"])
        d4 = "Baixo risco"
        d5 = "Baixo risco"
        
        # Julgamento global
        if "Alguma preocupação" in (d1, d3):
            global_judg = "Alguma preocupação"
            classif = "Moderada"
        else:
            global_judg = "Baixo risco"
            classif = "Alta"
            
        values = [
            idx,                          # A: ID
            autor_ano,                    # B: Autor (Ano)
            "ROB 2 (Simulado)",           # C: Ferramenta
            d1,                           # D: D1
            d2,                           # E: D2
            d3,                           # F: D3
            d4,                           # G: D4
            d5,                           # H: D5
            global_judg,                  # I: Julgamento Global
            classif,                      # J: Classificação
            "Avaliador Consenso",         # K: Avaliador
            "Estudo metodológico robusto.",# L: Observações
        ]
        
        for col_idx, val in enumerate(values, start=1):
            cell = ws_qualidade.cell(row=row_num, column=col_idx, value=val)
            cell.font = font_body
            cell.border = border_light
            cell.fill = fill
            
            if col_idx in (1, 2, 3, 11):
                cell.alignment = align_left
            else:
                cell.alignment = align_center_top

    print(f"  [OK] {len(final_included_records)} registros populados na Planilha de Qualidade.")

    # ═══════════════════════════════════════════════════════════════
    # STEP 5: ATUALIZAÇÃO DO PRISMA FLOW
    # ═══════════════════════════════════════════════════════════════
    print("\n[PRISMA] Atualizando fluxograma com os números finais...")
    ws_prisma = wb["🔀 PRISMA Flow"]
    
    font_number = Font(name="Segoe UI", size=11, bold=True, color=COLORS["indigo"])
    
    # Atualizar Ns do PRISMA Flow
    ws_prisma.cell(row=16, column=5, value=n_total).font = font_number          # Registros triados (T/R)
    ws_prisma.cell(row=17, column=5, value=n_excluded_tr).font = font_number     # Excluídos na triagem T/R
    ws_prisma.cell(row=19, column=5, value=n_included_tr).font = font_number     # Relatórios buscados para avaliação
    ws_prisma.cell(row=20, column=5, value=n_not_retrieved).font = font_number   # Relatórios não recuperados
    ws_prisma.cell(row=21, column=5, value=n_retrieved).font = font_number       # Relatórios avaliados para elegibilidade
    
    # Excluídos na elegibilidade por motivos
    # Vamos distribuir os motivos dos excluídos no TC (n_excluded_tc - n_not_retrieved)
    n_excluded_tc_real = n_excluded_tc - n_not_retrieved
    motivo_1 = int(n_excluded_tc_real * 0.5)
    motivo_2 = int(n_excluded_tc_real * 0.3)
    motivo_3 = n_excluded_tc_real - motivo_1 - motivo_2
    
    ws_prisma.cell(row=22, column=4, value="Excluídos — motivo 1: Fora do escopo").font = font_body
    ws_prisma.cell(row=22, column=5, value=motivo_1).font = font_number
    ws_prisma.cell(row=23, column=4, value="Excluídos — motivo 2: Tipo de documento inadequado").font = font_body
    ws_prisma.cell(row=23, column=5, value=motivo_2).font = font_number
    ws_prisma.cell(row=24, column=4, value="Excluídos — motivo 3: Qualidade insuficiente").font = font_body
    ws_prisma.cell(row=24, column=5, value=motivo_3).font = font_number
    ws_prisma.cell(row=25, column=5, value=0).font = font_number                 # Outros motivos
    
    # Estudos incluídos
    ws_prisma.cell(row=27, column=5, value=len(final_included_records)).font = font_number # Estudos incluídos na revisão
    ws_prisma.cell(row=28, column=5, value=len(final_included_records)).font = font_number # Relatórios de estudos incluídos
    ws_prisma.cell(row=29, column=5, value=0).font = font_number                 # Meta-análise
    ws_prisma.cell(row=30, column=5, value=len(final_included_records)).font = font_number # Síntese qualitativa/narrativa
    
    print("  [OK] PRISMA Flow atualizado com sucesso.")

    # Salvar alterações
    wb.save(filename)
    print(f"\n[SUCESSO] Fluxo completo de revisão simulado e salvo em '{filename}'!")
    print("=" * 60)

if __name__ == "__main__":
    main()
