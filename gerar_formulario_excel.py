"""
Gerador do Formulário de Desenho de Pesquisa para Revisão Sistemática (Excel)
==============================================================================
Gera um workbook .xlsx autoguiado com abas para cada etapa do protocolo,
planilhas de gestão (triagem, extração, qualidade, PRISMA) e uma aba de
configuração que exporta os JSONs para os coletores.

Baseado em:
  - Cochrane Handbook for Systematic Reviews of Interventions (2019)
  - Campbell et al. (2020) — Synthesis without meta-analysis (SWiM)
  - PRISMA 2020 Statement

Uso:
  python gerar_formulario_excel.py
"""

import os
import json
from datetime import datetime
import openpyxl
from openpyxl.styles import (
    Font, Alignment, Border, Side, PatternFill, NamedStyle, Protection
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

# ═══════════════════════════════════════════════════════════════
# DESIGN SYSTEM — Colors & Styles
# ═══════════════════════════════════════════════════════════════

# Color palette
COLORS = {
    "indigo_dark":   "1E1B4B",
    "indigo":        "4338CA",
    "indigo_mid":    "6366F1",
    "indigo_light":  "A5B4FC",
    "indigo_pale":   "E0E7FF",
    "indigo_wash":   "EEF2FF",
    "slate_900":     "0F172A",
    "slate_800":     "1E293B",
    "slate_700":     "334155",
    "slate_600":     "475569",
    "slate_500":     "64748B",
    "slate_400":     "94A3B8",
    "slate_200":     "E2E8F0",
    "slate_100":     "F1F5F9",
    "slate_50":      "F8FAFC",
    "white":         "FFFFFF",
    "black":         "000000",
    # Semantic
    "success":       "059669",
    "success_light": "D1FAE5",
    "success_wash":  "ECFDF5",
    "warning":       "D97706",
    "warning_light": "FEF3C7",
    "warning_wash":  "FFFBEB",
    "error":         "DC2626",
    "error_light":   "FEE2E2",
    "error_wash":    "FEF2F2",
    "info":          "2563EB",
    "info_light":    "DBEAFE",
    "info_wash":     "EFF6FF",
    # PICO cards
    "pico_p":        "DB2777",
    "pico_p_bg":     "FCE7F3",
    "pico_i":        "2563EB",
    "pico_i_bg":     "DBEAFE",
    "pico_c":        "D97706",
    "pico_c_bg":     "FEF3C7",
    "pico_o":        "059669",
    "pico_o_bg":     "D1FAE5",
    # PRISMA stages
    "prisma_id":     "4338CA",
    "prisma_id_bg":  "E0E7FF",
    "prisma_scr":    "D97706",
    "prisma_scr_bg": "FEF3C7",
    "prisma_elig":   "059669",
    "prisma_elig_bg":"D1FAE5",
    "prisma_inc":    "0891B2",
    "prisma_inc_bg": "CFFAFE",
    # Tab colors
    "tab_protocol":  "6366F1",
    "tab_pico":      "8B5CF6",
    "tab_eligib":    "10B981",
    "tab_search":    "F59E0B",
    "tab_sources":   "3B82F6",
    "tab_screening": "EF4444",
    "tab_extraction":"06B6D4",
    "tab_quality":   "8B5CF6",
    "tab_synthesis": "EC4899",
    "tab_config":    "64748B",
    "tab_triagem":   "F97316",
    "tab_extracao":  "0EA5E9",
    "tab_qualidade": "A855F7",
    "tab_prisma":    "14B8A6",
}

# Fill shortcuts
def fill(color_key, fill_type="solid"):
    return PatternFill(start_color=COLORS[color_key], end_color=COLORS[color_key], fill_type=fill_type)

# Font shortcuts
FONT_TITLE = Font(name="Segoe UI", size=18, bold=True, color=COLORS["indigo"])
FONT_SECTION = Font(name="Segoe UI", size=14, bold=True, color=COLORS["indigo_dark"])
FONT_SUBSECTION = Font(name="Segoe UI", size=12, bold=True, color=COLORS["slate_700"])
FONT_LABEL = Font(name="Segoe UI", size=11, bold=True, color=COLORS["slate_800"])
FONT_BODY = Font(name="Segoe UI", size=11, color=COLORS["slate_700"])
FONT_HINT = Font(name="Segoe UI", size=10, italic=True, color=COLORS["slate_500"])
FONT_INPUT = Font(name="Segoe UI", size=11, color=COLORS["indigo_dark"])
FONT_HEADER_COL = Font(name="Segoe UI", size=11, bold=True, color=COLORS["white"])
FONT_GUIDANCE = Font(name="Segoe UI", size=10, color=COLORS["info"])
FONT_GUIDANCE_BOLD = Font(name="Segoe UI", size=10, bold=True, color=COLORS["info"])
FONT_SMALL = Font(name="Segoe UI", size=9, color=COLORS["slate_400"])
FONT_NUMBER = Font(name="Consolas", size=12, bold=True, color=COLORS["indigo_mid"])

# Alignment shortcuts
ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_TOP_LEFT = Alignment(horizontal="left", vertical="top", wrap_text=True)
ALIGN_TOP_CENTER = Alignment(horizontal="center", vertical="top", wrap_text=True)

# Border
THIN_BORDER = Border(
    left=Side(style="thin", color=COLORS["slate_200"]),
    right=Side(style="thin", color=COLORS["slate_200"]),
    top=Side(style="thin", color=COLORS["slate_200"]),
    bottom=Side(style="thin", color=COLORS["slate_200"]),
)
INPUT_BORDER = Border(
    left=Side(style="thin", color=COLORS["indigo_light"]),
    right=Side(style="thin", color=COLORS["indigo_light"]),
    top=Side(style="thin", color=COLORS["indigo_light"]),
    bottom=Side(style="medium", color=COLORS["indigo_mid"]),
)
HEADER_BORDER = Border(
    bottom=Side(style="medium", color=COLORS["indigo"])
)


# ═══════════════════════════════════════════════════════════════
# HELPER FUNCTIONS
# ═══════════════════════════════════════════════════════════════

def set_col_widths(ws, widths_dict):
    """Set column widths from a dict {col_letter: width}."""
    for col, w in widths_dict.items():
        ws.column_dimensions[col].width = w

def write_title_banner(ws, row, text, last_col="H"):
    """Write a large title banner merged across columns."""
    ws.merge_cells(f"B{row}:{last_col}{row}")
    cell = ws[f"B{row}"]
    cell.value = text
    cell.font = FONT_TITLE
    cell.alignment = Alignment(horizontal="left", vertical="center")
    cell.fill = fill("indigo_wash")
    ws.row_dimensions[row].height = 45
    return row + 1

def write_section_header(ws, row, number, text, last_col="H"):
    """Write a numbered section header."""
    ws.merge_cells(f"B{row}:{last_col}{row}")
    cell = ws[f"B{row}"]
    cell.value = f"  {number}   {text}"
    cell.font = FONT_SECTION
    cell.alignment = Alignment(horizontal="left", vertical="center")
    cell.fill = fill("indigo_pale")
    ws.row_dimensions[row].height = 32
    # Left accent
    ws[f"A{row}"].fill = fill("indigo_mid")
    return row + 1

def write_guidance(ws, row, text, last_col="H"):
    """Write a methodology guidance callout."""
    ws.merge_cells(f"C{row}:{last_col}{row}")
    cell = ws[f"C{row}"]
    cell.value = text
    cell.font = FONT_GUIDANCE
    cell.alignment = ALIGN_TOP_LEFT
    cell.fill = fill("info_wash")
    cell.border = Border(left=Side(style="medium", color=COLORS["info"]))
    ws[f"B{row}"].fill = fill("info_wash")
    ws.row_dimensions[row].height = max(38, len(text) // 5)
    return row + 1

def write_label_input(ws, row, label, hint="", input_col="D", last_input_col="H", height=28):
    """Write a label in col B-C and an input cell in col D-H."""
    ws.merge_cells(f"B{row}:C{row}")
    cell_label = ws[f"B{row}"]
    cell_label.value = label
    cell_label.font = FONT_LABEL
    cell_label.alignment = ALIGN_LEFT
    
    ws.merge_cells(f"{input_col}{row}:{last_input_col}{row}")
    cell_input = ws[f"{input_col}{row}"]
    cell_input.fill = fill("slate_50")
    cell_input.border = INPUT_BORDER
    cell_input.font = FONT_INPUT
    cell_input.alignment = ALIGN_LEFT
    cell_input.protection = Protection(locked=False)
    
    ws.row_dimensions[row].height = height
    
    if hint:
        row += 1
        ws.merge_cells(f"D{row}:H{row}")
        cell_hint = ws[f"D{row}"]
        cell_hint.value = hint
        cell_hint.font = FONT_HINT
        cell_hint.alignment = ALIGN_LEFT
        ws.row_dimensions[row].height = 20
    
    return row + 1

def write_label_textarea(ws, row, label, hint="", num_rows=4, last_col="H"):
    """Write a label and a tall merged input area."""
    ws.merge_cells(f"B{row}:C{row}")
    cell_label = ws[f"B{row}"]
    cell_label.value = label
    cell_label.font = FONT_LABEL
    cell_label.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    
    end_row = row + num_rows - 1
    ws.merge_cells(f"D{row}:{last_col}{end_row}")
    cell_input = ws[f"D{row}"]
    cell_input.fill = fill("slate_50")
    cell_input.border = INPUT_BORDER
    cell_input.font = FONT_INPUT
    cell_input.alignment = ALIGN_TOP_LEFT
    cell_input.protection = Protection(locked=False)
    
    for r in range(row, end_row + 1):
        ws.row_dimensions[r].height = 22
    
    next_row = end_row + 1
    if hint:
        ws.merge_cells(f"D{next_row}:{last_col}{next_row}")
        ws[f"D{next_row}"].value = hint
        ws[f"D{next_row}"].font = FONT_HINT
        ws[f"D{next_row}"].alignment = ALIGN_LEFT
        ws.row_dimensions[next_row].height = 20
        next_row += 1
    
    return next_row

def write_spacer(ws, row, height=10):
    ws.row_dimensions[row].height = height
    return row + 1

def write_numbered_list(ws, row, label, items, input_col="D", last_col="H"):
    """Write a label and numbered input rows for a list."""
    ws[f"B{row}"].value = label
    ws[f"B{row}"].font = FONT_LABEL
    ws[f"B{row}"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws.row_dimensions[row].height = 24
    row += 1
    
    start_row = row
    for i, item_text in enumerate(items, 1):
        ws[f"C{row}"].value = f"{i}."
        ws[f"C{row}"].font = FONT_NUMBER
        ws[f"C{row}"].alignment = ALIGN_CENTER
        ws.merge_cells(f"{input_col}{row}:{last_col}{row}")
        cell = ws[f"{input_col}{row}"]
        cell.value = item_text if item_text else ""
        cell.fill = fill("slate_50")
        cell.border = INPUT_BORDER
        cell.font = FONT_INPUT
        cell.alignment = ALIGN_LEFT
        cell.protection = Protection(locked=False)
        ws.row_dimensions[row].height = 26
        row += 1
    
    # Extra blank rows for user additions
    for i in range(len(items) + 1, len(items) + 6):
        ws[f"C{row}"].value = f"{i}."
        ws[f"C{row}"].font = Font(name="Consolas", size=12, color=COLORS["slate_400"])
        ws[f"C{row}"].alignment = ALIGN_CENTER
        ws.merge_cells(f"{input_col}{row}:{last_col}{row}")
        cell = ws[f"{input_col}{row}"]
        cell.fill = fill("slate_50")
        cell.border = INPUT_BORDER
        cell.font = FONT_INPUT
        cell.alignment = ALIGN_LEFT
        cell.protection = Protection(locked=False)
        ws.row_dimensions[row].height = 26
        row += 1
    
    return row

def write_table_header(ws, row, headers, col_start="B"):
    """Write a styled table header row."""
    for i, h in enumerate(headers):
        col = get_column_letter(ord(col_start) - 64 + i)
        cell = ws[f"{col}{row}"]
        cell.value = h
        cell.font = FONT_HEADER_COL
        cell.fill = fill("indigo")
        cell.alignment = ALIGN_CENTER
        cell.border = THIN_BORDER
    ws.row_dimensions[row].height = 30
    return row + 1

def write_table_row(ws, row, values, col_start="B", editable=True, zebra=False):
    """Write a data row (optionally with zebra striping)."""
    bg = "slate_100" if zebra else "white"
    for i, v in enumerate(values):
        col = get_column_letter(ord(col_start) - 64 + i)
        cell = ws[f"{col}{row}"]
        cell.value = v
        cell.font = FONT_BODY if not editable else FONT_INPUT
        cell.fill = fill(bg)
        cell.alignment = ALIGN_LEFT
        cell.border = THIN_BORDER
        if editable:
            cell.protection = Protection(locked=False)
    ws.row_dimensions[row].height = 26
    return row + 1


# ═══════════════════════════════════════════════════════════════
# SHEET BUILDERS
# ═══════════════════════════════════════════════════════════════

def build_instructions(wb):
    """Aba 0: Capa e instruções."""
    ws = wb.active
    ws.title = "📖 Instruções"
    ws.sheet_properties.tabColor = COLORS["tab_protocol"]
    ws.sheet_view.showGridLines = False
    set_col_widths(ws, {"A": 4, "B": 28, "C": 8, "D": 20, "E": 20, "F": 20, "G": 20, "H": 20})
    
    row = 2
    row = write_title_banner(ws, row, "Formulário de Desenho de Pesquisa — Revisão Sistemática da Literatura")
    row = write_spacer(ws, row)
    
    ws.merge_cells(f"B{row}:H{row}")
    ws[f"B{row}"].value = "Configure cada etapa do protocolo de pesquisa e exporte automaticamente os arquivos para os coletores e planilhas de gestão."
    ws[f"B{row}"].font = FONT_BODY
    ws[f"B{row}"].alignment = ALIGN_LEFT
    row += 2
    
    # Instruction table
    instructions = [
        ("📌", "1. Protocolo",        "Identifique o projeto, equipe e registro PROSPERO."),
        ("❓", "2. Questão (PICO)",    "Defina a pergunta de pesquisa com framework estruturado."),
        ("✅", "3. Elegibilidade",     "Critérios de inclusão/exclusão, tipos de estudo e idiomas."),
        ("🔍", "4. Estratégia de Busca","Monte os blocos booleanos e a expressão de pesquisa."),
        ("🗄️", "5. Fontes de Dados",   "Selecione as bases e configure as chaves de API."),
        ("🔬", "6. Triagem",           "Defina revisores, conflitos e piloto de calibração."),
        ("📊", "7. Extração de Dados", "Variáveis para o formulário padronizado de extração."),
        ("🏅", "8. Avaliação Qualidade","Ferramenta de risco de viés (ROB 2, NOS, JBI, etc.)."),
        ("📈", "9. Síntese",           "Tipo de síntese e software de análise planejado."),
    ]
    
    row = write_table_header(ws, row, ["", "Aba do Formulário", "O que preencher"], col_start="B")
    for i, (icon, aba, desc) in enumerate(instructions):
        vals = [icon, aba, desc]
        row = write_table_row(ws, row, vals, col_start="B", editable=False, zebra=i % 2 == 1)
    
    row = write_spacer(ws, row, 20)
    
    ws.merge_cells(f"B{row}:H{row}")
    cell = ws[f"B{row}"]
    cell.value = "PLANILHAS DE GESTÃO GERADAS"
    cell.font = FONT_SECTION
    cell.fill = fill("success_light")
    ws.row_dimensions[row].height = 32
    row += 1
    
    management_sheets = [
        ("📊", "Planilha Triagem",     "Screening por título/resumo e texto completo com 2 revisores."),
        ("📋", "Planilha Extração",    "Formulário padronizado com as variáveis definidas na Etapa 7."),
        ("🏅", "Planilha Qualidade",   "Avaliação de risco de viés por domínios (ferramenta da Etapa 8)."),
        ("🔀", "PRISMA Flow",          "Dados para preenchimento do fluxograma PRISMA 2020."),
    ]
    row = write_table_header(ws, row, ["", "Planilha", "Finalidade"], col_start="B")
    for i, (icon, aba, desc) in enumerate(management_sheets):
        row = write_table_row(ws, row, [icon, aba, desc], col_start="B", editable=False, zebra=i % 2 == 1)
    
    row += 2
    row = write_guidance(ws, row,
        "📖 Referências: Cochrane Handbook v6.0 (Higgins et al., 2019) · Campbell et al. (2020), BMJ 368:l6890 · PRISMA 2020 (Page et al., 2021)"
    )

    row += 1
    ws.merge_cells(f"B{row}:H{row}")
    ws[f"B{row}"].value = "ℹ️  Células com fundo cinza claro e borda inferior azul são campos editáveis. Preencha-os com os dados do seu projeto."
    ws[f"B{row}"].font = FONT_HINT
    ws[f"B{row}"].fill = fill("warning_wash")
    ws[f"B{row}"].alignment = ALIGN_LEFT
    ws.row_dimensions[row].height = 28


def build_protocolo(wb):
    """Aba 1: Identificação do Protocolo."""
    ws = wb.create_sheet("1. Protocolo")
    ws.sheet_properties.tabColor = COLORS["tab_protocol"]
    ws.sheet_view.showGridLines = False
    set_col_widths(ws, {"A": 4, "B": 22, "C": 8, "D": 20, "E": 20, "F": 20, "G": 20, "H": 20})
    
    row = 2
    row = write_title_banner(ws, row, "Etapa 1 — Identificação do Protocolo")
    row = write_spacer(ws, row)
    row = write_guidance(ws, row,
        "📖 Cochrane (2019), Cap. 1: O protocolo deve ser registrado previamente (ex.: PROSPERO) para garantir transparência e evitar duplicação."
    )
    row = write_spacer(ws, row)
    
    row = write_section_header(ws, row, "1.1", "Dados do Projeto")
    row = write_spacer(ws, row)
    row = write_label_textarea(ws, row, "Título da Revisão *", 
        hint='Deve conter a estrutura PICO. Ex.: "Efeitos de [Intervenção] em [Desfecho] para [População]: Uma Revisão Sistemática"',
        num_rows=2)
    row = write_spacer(ws, row)
    row = write_label_input(ws, row, "Pesquisador Principal *")
    row = write_label_input(ws, row, "E-mail de contato")
    row = write_label_input(ws, row, "Instituição / Programa", hint="Ex.: PPGES - UFES")
    row = write_label_input(ws, row, "Data do Protocolo", hint="DD/MM/AAAA")
    row = write_label_input(ws, row, "Registro PROSPERO", hint="Ex.: CRD42024000000. Deixe em branco se não registrado.")
    
    row = write_spacer(ws, row, 20)
    row = write_section_header(ws, row, "1.2", "Equipe de Revisão")
    row = write_spacer(ws, row)
    row = write_guidance(ws, row,
        "Mínimo recomendado: 2 revisores independentes para triagem (Cochrane, 2019, Cap. 4.6)."
    )
    row = write_spacer(ws, row)
    
    # Team table
    headers = ["#", "Nome Completo", "Função", "Instituição", "E-mail"]
    row = write_table_header(ws, row, headers, col_start="B")
    roles_dv = DataValidation(
        type="list",
        formula1='"Revisor,Orientador,Co-orientador,Bibliotecário,Estatístico,Analista"',
        allow_blank=True
    )
    roles_dv.error = "Selecione uma função válida"
    roles_dv.errorTitle = "Função inválida"
    ws.add_data_validation(roles_dv)
    
    for i in range(1, 9):
        vals = [str(i), "", "", "", ""]
        row = write_table_row(ws, row, vals, col_start="B", editable=True, zebra=i % 2 == 0)
        roles_dv.add(ws[f"D{row-1}"])


def build_pico(wb):
    """Aba 2: Questão de Pesquisa (PICO)."""
    ws = wb.create_sheet("2. Questão PICO")
    ws.sheet_properties.tabColor = COLORS["tab_pico"]
    ws.sheet_view.showGridLines = False
    set_col_widths(ws, {"A": 4, "B": 22, "C": 8, "D": 20, "E": 20, "F": 20, "G": 20, "H": 20})
    
    row = 2
    row = write_title_banner(ws, row, "Etapa 2 — Questão de Pesquisa")
    row = write_spacer(ws, row)
    row = write_guidance(ws, row,
        "📖 Cochrane (2019), Cap. 2: A pergunta deve ser formulável dentro de um framework (PICO para intervenção, PICo para exploratória, SPIDER para qualitativa, PEO para exposição)."
    )
    row = write_spacer(ws, row)
    
    row = write_section_header(ws, row, "2.1", "Framework de Estruturação")
    row = write_spacer(ws, row)
    
    # Framework selector
    row = write_label_input(ws, row, "Framework escolhido *",
        hint="PICO (Intervenção) | PICo (Exploratória) | SPIDER (Qualitativa) | PEO (Exposição)")
    fw_dv = DataValidation(type="list", formula1='"PICO,PICo,SPIDER,PEO"', allow_blank=False)
    fw_dv.error = "Selecione um framework válido"
    ws.add_data_validation(fw_dv)
    fw_dv.add(ws[f"D{row-2}"])
    ws[f"D{row-2}"].value = "PICo"
    
    row = write_spacer(ws, row, 20)
    row = write_section_header(ws, row, "2.2", "Componentes do Framework")
    row = write_spacer(ws, row)
    
    # PICO cards as colored blocks
    pico_items = [
        ("P", "População / Problema", "pico_p", "pico_p_bg",
         "Quem são os participantes ou qual é o contexto do problema?"),
        ("I", "Intervenção / Interesse", "pico_i", "pico_i_bg",
         "Qual fenômeno de interesse, intervenção ou exposição?"),
        ("C", "Contexto / Comparação", "pico_c", "pico_c_bg",
         "Qual o contexto ou grupo comparador?"),
        ("O", "Resultado (Outcome)", "pico_o", "pico_o_bg",
         "Quais desfechos ou resultados são avaliados?"),
    ]
    
    for letter, title, color_key, bg_key, hint_text in pico_items:
        # Header bar
        ws.merge_cells(f"B{row}:C{row}")
        cell_letter = ws[f"B{row}"]
        cell_letter.value = f"  {letter} — {title}"
        cell_letter.font = Font(name="Segoe UI", size=12, bold=True, color=COLORS[color_key])
        cell_letter.fill = fill(bg_key)
        cell_letter.alignment = ALIGN_LEFT
        ws.merge_cells(f"D{row}:H{row}")
        ws[f"D{row}"].fill = fill(bg_key)
        ws.row_dimensions[row].height = 30
        row += 1
        
        # Hint
        ws.merge_cells(f"C{row}:H{row}")
        ws[f"C{row}"].value = hint_text
        ws[f"C{row}"].font = FONT_HINT
        ws.row_dimensions[row].height = 20
        row += 1
        
        # Input area
        row = write_label_textarea(ws, row, "", num_rows=3)
        row = write_spacer(ws, row)
    
    row = write_spacer(ws, row, 15)
    row = write_section_header(ws, row, "2.3", "Pergunta de Pesquisa")
    row = write_spacer(ws, row)
    row = write_label_textarea(ws, row, "Pergunta completa *",
        hint="Formule a pergunta de pesquisa como uma sentença completa integrando os componentes acima.",
        num_rows=4)


def build_elegibilidade(wb):
    """Aba 3: Critérios de Elegibilidade."""
    ws = wb.create_sheet("3. Elegibilidade")
    ws.sheet_properties.tabColor = COLORS["tab_eligib"]
    ws.sheet_view.showGridLines = False
    set_col_widths(ws, {"A": 4, "B": 22, "C": 8, "D": 20, "E": 20, "F": 20, "G": 20, "H": 20})
    
    row = 2
    row = write_title_banner(ws, row, "Etapa 3 — Critérios de Elegibilidade")
    row = write_spacer(ws, row)
    row = write_guidance(ws, row,
        "📖 Campbell (2020): Os critérios devem ser definidos a priori para evitar viés de seleção. Documente tipos de estudo, participantes, intervenção e desfechos aceitáveis."
    )
    row = write_spacer(ws, row)
    
    # Inclusion
    row = write_section_header(ws, row, "3.1", "Critérios de Inclusão")
    row = write_spacer(ws, row)
    row = write_numbered_list(ws, row, "Inclusão *", [
        "Artigos publicados em periódicos revisados por pares",
        "Estudos que abordem o tema de interesse definido no PICO",
        "",
        "",
    ])
    
    row = write_spacer(ws, row, 20)
    
    # Exclusion
    row = write_section_header(ws, row, "3.2", "Critérios de Exclusão")
    row = write_spacer(ws, row)
    row = write_numbered_list(ws, row, "Exclusão *", [
        "Artigos de opinião, editoriais, cartas ao editor",
        "Estudos sem texto completo disponível",
        "",
        "",
    ])
    
    row = write_spacer(ws, row, 20)
    
    # Study types
    row = write_section_header(ws, row, "3.3", "Tipos de Estudo e Idiomas")
    row = write_spacer(ws, row)
    
    study_types = [
        "Ensaio clínico randomizado (ECR)", "Estudo de coorte", "Caso-controle",
        "Estudo transversal", "Quasi-experimental", "Estudo qualitativo",
        "Revisão de escopo", "Métodos mistos"
    ]
    headers = ["#", "Tipo de Estudo", "Aceito? (S/N)"]
    row = write_table_header(ws, row, headers, col_start="B")
    yn_dv = DataValidation(type="list", formula1='"Sim,Não"', allow_blank=True)
    ws.add_data_validation(yn_dv)
    
    for i, st in enumerate(study_types, 1):
        row = write_table_row(ws, row, [str(i), st, ""], col_start="B", editable=True, zebra=i % 2 == 0)
        yn_dv.add(ws[f"D{row-1}"])
    
    row = write_spacer(ws, row, 15)
    
    languages = ["Português", "Inglês", "Espanhol", "Francês", "Sem restrição de idioma"]
    headers2 = ["#", "Idioma", "Aceito? (S/N)"]
    row = write_table_header(ws, row, headers2, col_start="B")
    yn_dv2 = DataValidation(type="list", formula1='"Sim,Não"', allow_blank=True)
    ws.add_data_validation(yn_dv2)
    
    for i, lang in enumerate(languages, 1):
        default = "Sim" if lang in ["Português", "Inglês", "Espanhol"] else ""
        row = write_table_row(ws, row, [str(i), lang, default], col_start="B", editable=True, zebra=i % 2 == 0)
        yn_dv2.add(ws[f"D{row-1}"])
    
    # Temporal
    row = write_spacer(ws, row, 15)
    row = write_label_input(ws, row, "Ano inicial", hint="Ex.: 2000")
    row = write_label_input(ws, row, "Ano final", hint="Ex.: 2026")


def build_estrategia_busca(wb):
    """Aba 4: Estratégia de Busca."""
    ws = wb.create_sheet("4. Estratégia Busca")
    ws.sheet_properties.tabColor = COLORS["tab_search"]
    ws.sheet_view.showGridLines = False
    set_col_widths(ws, {"A": 4, "B": 22, "C": 8, "D": 20, "E": 20, "F": 20, "G": 20, "H": 20})
    
    row = 2
    row = write_title_banner(ws, row, "Etapa 4 — Estratégia de Busca")
    row = write_spacer(ws, row)
    row = write_guidance(ws, row,
        '📖 Cochrane (2019), Cap. 4: A estratégia deve ser altamente sensível. Use termos livres, descritores (MeSH/DeCS) e sinônimos. '
        'Combine blocos conceituais com AND; dentro de cada bloco, ligue sinônimos com OR.'
    )
    row = write_spacer(ws, row)
    
    # Block 1
    row = write_section_header(ws, row, "4.1", "Bloco 1 — População / Problema")
    row = write_spacer(ws, row)
    ws.merge_cells(f"C{row}:H{row}")
    ws[f"C{row}"].value = "Liste sinônimos, termos MeSH/DeCS e variações. Cada termo neste bloco será ligado por OR."
    ws[f"C{row}"].font = FONT_HINT
    row += 1
    row = write_numbered_list(ws, row, "Termos do Bloco 1 *", ["", "", "", ""])
    
    row = write_spacer(ws, row, 15)
    
    # Block 2
    row = write_section_header(ws, row, "4.2", "Bloco 2 — Intervenção / Interesse")
    row = write_spacer(ws, row)
    ws.merge_cells(f"C{row}:H{row}")
    ws[f"C{row}"].value = "Termos do conceito de intervenção ou fenômeno de interesse."
    ws[f"C{row}"].font = FONT_HINT
    row += 1
    row = write_numbered_list(ws, row, "Termos do Bloco 2 *", ["", "", "", ""])
    
    row = write_spacer(ws, row, 15)
    
    # Block 3
    row = write_section_header(ws, row, "4.3", "Bloco 3 — Contexto / Desfecho (Opcional)")
    row = write_spacer(ws, row)
    ws.merge_cells(f"C{row}:H{row}")
    ws[f"C{row}"].value = "Termos adicionais para refinar. Deixe em branco se não for necessário restringir."
    ws[f"C{row}"].font = FONT_HINT
    row += 1
    row = write_numbered_list(ws, row, "Termos do Bloco 3", ["", "", ""])
    
    row = write_spacer(ws, row, 20)
    
    # Operator
    row = write_section_header(ws, row, "4.4", "Montagem da Expressão Booleana")
    row = write_spacer(ws, row)
    row = write_label_input(ws, row, "Operador entre blocos",
        hint="AND ou AND NOT")
    op_dv = DataValidation(type="list", formula1='"AND,AND NOT"', allow_blank=False)
    ws.add_data_validation(op_dv)
    op_dv.add(ws[f"D{row-2}"])
    ws[f"D{row-2}"].value = "AND"
    
    row = write_spacer(ws, row)
    row = write_label_textarea(ws, row, "Expressão final *",
        hint="Cole ou monte a expressão booleana completa. Ex.: (\"saneamento\" OR \"water sanitation\") AND (\"inferência causal\" OR \"causal inference\")",
        num_rows=4)
    
    row = write_spacer(ws, row)
    row = write_label_input(ws, row, "Campo Scopus",
        hint="TITLE-ABS-KEY | TITLE | ABS | KEY | ALL")
    field_dv = DataValidation(type="list", formula1='"TITLE-ABS-KEY,TITLE,ABS,KEY,ALL"', allow_blank=True)
    ws.add_data_validation(field_dv)
    field_dv.add(ws[f"D{row-2}"])
    ws[f"D{row-2}"].value = "TITLE-ABS-KEY"
    
    row = write_spacer(ws, row)
    row = write_label_input(ws, row, "Ano inicial", hint="Ex.: 2000")
    row = write_label_input(ws, row, "Ano final", hint="Ex.: 2026")


def build_fontes(wb):
    """Aba 5: Fontes de Dados."""
    ws = wb.create_sheet("5. Fontes de Dados")
    ws.sheet_properties.tabColor = COLORS["tab_sources"]
    ws.sheet_view.showGridLines = False
    set_col_widths(ws, {"A": 4, "B": 22, "C": 8, "D": 20, "E": 20, "F": 20, "G": 20, "H": 20})
    
    row = 2
    row = write_title_banner(ws, row, "Etapa 5 — Fontes de Dados")
    row = write_spacer(ws, row)
    row = write_guidance(ws, row,
        "📖 Cochrane (2019), Cap. 4.3: No mínimo MEDLINE e Cochrane Library. Campbell (2020) recomenda ≥3 bases. As bases escolhidas determinam os configs gerados."
    )
    row = write_spacer(ws, row)
    
    row = write_section_header(ws, row, "5.1", "Bases Automatizadas (Software)")
    row = write_spacer(ws, row)
    
    bases = [
        ("BDTD", "Teses e Dissertações brasileiras (IBICT)", "Sim"),
        ("SciELO", "Periódicos Latino-americanos de acesso aberto", "Sim"),
        ("Scopus", "Base global Elsevier (requer API key)", "Sim"),
        ("OpenAlex", "Agregador global aberto (OurResearch)", "Sim"),
    ]
    headers = ["#", "Base", "Descrição", "Usar? (S/N)"]
    row = write_table_header(ws, row, headers, col_start="B")
    use_dv = DataValidation(type="list", formula1='"Sim,Não"', allow_blank=True)
    ws.add_data_validation(use_dv)
    
    for i, (name, desc, default) in enumerate(bases, 1):
        row_data = [str(i), name, desc, default]
        row = write_table_row(ws, row, row_data, col_start="B", editable=True, zebra=i % 2 == 0)
        use_dv.add(ws[f"E{row-1}"])
    
    row = write_spacer(ws, row, 20)
    row = write_section_header(ws, row, "5.2", "Bases Consultadas Manualmente")
    row = write_spacer(ws, row)
    
    row = write_numbered_list(ws, row, "Bases manuais", [
        "PubMed / MEDLINE",
        "Web of Science",
        "",
        "",
    ])
    
    row = write_spacer(ws, row, 20)
    row = write_section_header(ws, row, "5.3", "Configurações de API")
    row = write_spacer(ws, row)
    row = write_label_input(ws, row, "Scopus API Key",
        hint="Obtenha em https://dev.elsevier.com/. Deixe em branco para usar o config atual.")
    row = write_label_input(ws, row, "E-mail Polite Pool",
        hint="Para OpenAlex/SciELO. Garante limites ampliados.")


def build_triagem(wb):
    """Aba 6: Triagem e Seleção."""
    ws = wb.create_sheet("6. Triagem")
    ws.sheet_properties.tabColor = COLORS["tab_screening"]
    ws.sheet_view.showGridLines = False
    set_col_widths(ws, {"A": 4, "B": 22, "C": 8, "D": 20, "E": 20, "F": 20, "G": 20, "H": 20})
    
    row = 2
    row = write_title_banner(ws, row, "Etapa 6 — Triagem e Seleção")
    row = write_spacer(ws, row)
    row = write_guidance(ws, row,
        "📖 Cochrane (2019), Cap. 4.6: Triagem em 2 fases — (1) título/resumo, (2) texto completo. "
        "Pelo menos 2 revisores independentes. Conflitos resolvidos por consenso ou 3º revisor."
    )
    row = write_spacer(ws, row)
    
    row = write_section_header(ws, row, "6.1", "Processo de Triagem")
    row = write_spacer(ws, row)
    
    row = write_label_input(ws, row, "Nº de revisores *",
        hint="Recomendado: 2 revisores independentes")
    rev_dv = DataValidation(type="list", formula1='"1,2,3 ou mais"', allow_blank=False)
    ws.add_data_validation(rev_dv)
    rev_dv.add(ws[f"D{row-2}"])
    ws[f"D{row-2}"].value = "2"
    
    row = write_label_input(ws, row, "Resolução de conflitos")
    conf_dv = DataValidation(type="list",
        formula1='"Consenso entre revisores,Terceiro revisor independente,Voto majoritário,Decisão do orientador"',
        allow_blank=True)
    ws.add_data_validation(conf_dv)
    conf_dv.add(ws[f"D{row-1}"])
    ws[f"D{row-1}"].value = "Consenso entre revisores"
    
    row = write_label_input(ws, row, "Software de triagem")
    sw_dv = DataValidation(type="list",
        formula1='"Planilhas geradas pelo coletor (Excel),Rayyan,Covidence,ASReview,Parsifal,Mendeley,Zotero,EndNote"',
        allow_blank=True)
    ws.add_data_validation(sw_dv)
    sw_dv.add(ws[f"D{row-1}"])
    ws[f"D{row-1}"].value = "Planilhas geradas pelo coletor (Excel)"
    
    row = write_spacer(ws, row, 15)
    row = write_section_header(ws, row, "6.2", "Piloto de Calibração")
    row = write_spacer(ws, row)
    row = write_guidance(ws, row,
        "Cochrane recomenda calibrar os critérios com um lote-piloto antes da triagem completa para aferir concordância inter-avaliador."
    )
    row = write_spacer(ws, row)
    row = write_label_input(ws, row, "Tamanho do lote-piloto", hint="Nº de estudos para calibrar (ex.: 30)")
    ws[f"D{row-2}"].value = "30"
    row = write_label_input(ws, row, "Meta de concordância", hint="Kappa mínimo aceitável (ex.: ≥ 0.80)")
    ws[f"D{row-2}"].value = "≥ 0.80"


def build_extracao_dados(wb):
    """Aba 7: Extração de Dados."""
    ws = wb.create_sheet("7. Extração Dados")
    ws.sheet_properties.tabColor = COLORS["tab_extraction"]
    ws.sheet_view.showGridLines = False
    set_col_widths(ws, {"A": 4, "B": 22, "C": 8, "D": 20, "E": 20, "F": 20, "G": 20, "H": 20})
    
    row = 2
    row = write_title_banner(ws, row, "Etapa 7 — Extração de Dados")
    row = write_spacer(ws, row)
    row = write_guidance(ws, row,
        "📖 Cochrane (2019), Cap. 5: Use formulários padronizados. Extraia dados bibliográficos, características do estudo, "
        "intervenção, desfechos e resultados numéricos. Faça extração dupla para minimizar erros."
    )
    row = write_spacer(ws, row)
    
    row = write_section_header(ws, row, "7.1", "Variáveis para Extração")
    row = write_spacer(ws, row)
    
    variables = [
        ("Título", "Auto", "Sim"),
        ("Autores", "Auto", "Sim"),
        ("Ano de publicação", "Auto", "Sim"),
        ("DOI", "Auto", "Sim"),
        ("Periódico / Fonte", "Auto", "Sim"),
        ("Resumo (Abstract)", "Auto", "Sim"),
        ("Palavras-chave", "Auto", "Sim"),
        ("Objetivo do estudo", "Manual", ""),
        ("Desenho do estudo", "Manual", ""),
        ("Tamanho da amostra", "Manual", ""),
        ("Método estatístico principal", "Manual", ""),
        ("Principais resultados", "Manual", ""),
        ("Limitações reportadas", "Manual", ""),
        ("País do estudo", "Manual", ""),
        ("Financiamento / Conflito de interesse", "Manual", ""),
    ]
    
    headers = ["#", "Variável", "Tipo", "Extrair? (S/N)"]
    row = write_table_header(ws, row, headers, col_start="B")
    extract_dv = DataValidation(type="list", formula1='"Sim,Não"', allow_blank=True)
    ws.add_data_validation(extract_dv)
    
    for i, (var_name, var_type, default) in enumerate(variables, 1):
        row = write_table_row(ws, row, [str(i), var_name, var_type, default],
            col_start="B", editable=True, zebra=i % 2 == 0)
        extract_dv.add(ws[f"E{row-1}"])
    
    row = write_spacer(ws, row, 15)
    row = write_section_header(ws, row, "7.2", "Variáveis Personalizadas")
    row = write_spacer(ws, row)
    row = write_numbered_list(ws, row, "Variáveis adicionais", ["", "", "", ""])
    
    row = write_spacer(ws, row, 15)
    row = write_section_header(ws, row, "7.3", "Extração Dupla")
    row = write_spacer(ws, row)
    row = write_label_input(ws, row, "Extração dupla?",
        hint="Sim (recomendado) | Não | Parcial (amostra de verificação)")
    dup_dv = DataValidation(type="list", formula1='"Sim (recomendado),Não,Parcial (amostra de verificação)"', allow_blank=True)
    ws.add_data_validation(dup_dv)
    dup_dv.add(ws[f"D{row-2}"])
    ws[f"D{row-2}"].value = "Sim (recomendado)"


def build_qualidade(wb):
    """Aba 8: Avaliação de Qualidade."""
    ws = wb.create_sheet("8. Qualidade")
    ws.sheet_properties.tabColor = COLORS["tab_quality"]
    ws.sheet_view.showGridLines = False
    set_col_widths(ws, {"A": 4, "B": 22, "C": 8, "D": 20, "E": 20, "F": 20, "G": 20, "H": 20})
    
    row = 2
    row = write_title_banner(ws, row, "Etapa 8 — Avaliação da Qualidade Metodológica")
    row = write_spacer(ws, row)
    row = write_guidance(ws, row,
        "📖 Cochrane (2019), Cap. 8: Use ferramentas validadas. ROB 2 para ECRs, ROBINS-I para não randomizados, "
        "Newcastle-Ottawa para coorte/caso-controle, JBI/MMAT/CASP conforme o desenho."
    )
    row = write_spacer(ws, row)
    
    row = write_section_header(ws, row, "8.1", "Ferramenta de Avaliação")
    row = write_spacer(ws, row)
    
    row = write_label_input(ws, row, "Ferramenta de viés *")
    tool_dv = DataValidation(type="list",
        formula1='"ROB 2 (Cochrane - ECRs),ROBINS-I (Não randomizados),Newcastle-Ottawa Scale (NOS),JBI Critical Appraisal,MMAT (Métodos mistos),CASP,GRADE,Não aplicável (revisão de escopo),Outra"',
        allow_blank=False)
    ws.add_data_validation(tool_dv)
    tool_dv.add(ws[f"D{row-1}"])
    
    row = write_label_input(ws, row, "Classificação")
    class_dv = DataValidation(type="list",
        formula1='"Alto/Moderado/Baixo risco,A/B/C,Pontuação numérica"',
        allow_blank=True)
    ws.add_data_validation(class_dv)
    class_dv.add(ws[f"D{row-1}"])
    ws[f"D{row-1}"].value = "Alto/Moderado/Baixo risco"
    
    row = write_spacer(ws, row)
    row = write_label_textarea(ws, row, "Observações",
        hint="Domínios específicos, adaptações da ferramenta, etc.", num_rows=4)
    
    row = write_spacer(ws, row, 20)
    row = write_section_header(ws, row, "8.2", "Domínios por Ferramenta (Referência)")
    row = write_spacer(ws, row)
    
    # Reference table of domains per tool
    ref_data = [
        ("ROB 2", "D1: Randomização | D2: Desvios da intervenção | D3: Dados faltantes | D4: Mensuração do desfecho | D5: Seleção do resultado"),
        ("ROBINS-I", "D1: Confundimento | D2: Seleção | D3: Classificação | D4: Desvios | D5: Dados faltantes | D6: Mensuração | D7: Seleção do resultado"),
        ("NOS", "Seleção (4 itens) | Comparabilidade (2 itens) | Desfecho/Exposição (3 itens) — máx. 9 estrelas"),
        ("JBI", "8-13 critérios específicos por tipo de estudo (checklist dedicado)"),
        ("MMAT", "2 critérios de triagem + 5 critérios específicos ao desenho"),
    ]
    headers = ["Ferramenta", "Domínios de Avaliação"]
    row = write_table_header(ws, row, headers, col_start="B")
    for i, (tool, domains) in enumerate(ref_data):
        ws.merge_cells(f"C{row}:H{row}")
        row = write_table_row(ws, row, [tool, domains], col_start="B", editable=False, zebra=i % 2 == 0)


def build_sintese(wb):
    """Aba 9: Síntese dos Resultados."""
    ws = wb.create_sheet("9. Síntese")
    ws.sheet_properties.tabColor = COLORS["tab_synthesis"]
    ws.sheet_view.showGridLines = False
    set_col_widths(ws, {"A": 4, "B": 22, "C": 8, "D": 20, "E": 20, "F": 20, "G": 20, "H": 20})
    
    row = 2
    row = write_title_banner(ws, row, "Etapa 9 — Síntese dos Resultados")
    row = write_spacer(ws, row)
    row = write_guidance(ws, row,
        "📖 Cochrane (2019), Cap. 10-12: Meta-análise é viável quando estudos são homogêneos. "
        "Caso contrário, opte por síntese narrativa (Campbell, 2020 — SWiM). Avalie I² e planeje análises de sensibilidade."
    )
    row = write_spacer(ws, row)
    
    row = write_section_header(ws, row, "9.1", "Tipo de Síntese")
    row = write_spacer(ws, row)
    
    synthesis_types = [
        "Meta-análise quantitativa",
        "Síntese narrativa estruturada",
        "Síntese temática (qualitativa)",
        "Análise bibliométrica",
        "Mapeamento sistemático",
    ]
    headers = ["#", "Tipo de Síntese", "Planejado? (S/N)"]
    row = write_table_header(ws, row, headers, col_start="B")
    plan_dv = DataValidation(type="list", formula1='"Sim,Não"', allow_blank=True)
    ws.add_data_validation(plan_dv)
    
    for i, st in enumerate(synthesis_types, 1):
        row = write_table_row(ws, row, [str(i), st, ""], col_start="B", editable=True, zebra=i % 2 == 0)
        plan_dv.add(ws[f"D{row-1}"])
    
    row = write_spacer(ws, row, 15)
    row = write_section_header(ws, row, "9.2", "Software e Análises")
    row = write_spacer(ws, row)
    
    row = write_label_input(ws, row, "Software de análise")
    sw_dv = DataValidation(type="list",
        formula1='"R (meta / metafor),Python (pandas + statsmodels),RevMan (Cochrane),STATA,Não definido"',
        allow_blank=True)
    ws.add_data_validation(sw_dv)
    sw_dv.add(ws[f"D{row-1}"])
    
    row = write_spacer(ws, row)
    row = write_label_textarea(ws, row, "Análises adicionais",
        hint="Análises de subgrupo, sensibilidade, funnel plot, etc.", num_rows=4)


# ═══════════════════════════════════════════════════════════════
# MANAGEMENT SHEETS
# ═══════════════════════════════════════════════════════════════

def build_planilha_triagem(wb):
    """Planilha de Triagem (Screening)."""
    ws = wb.create_sheet("📊 Planilha Triagem")
    ws.sheet_properties.tabColor = COLORS["tab_triagem"]
    ws.sheet_view.showGridLines = True
    
    headers = [
        "ID", "Base", "DOI", "Título", "Autores", "Ano", "Periódico",
        "Resumo", "Palavras-Chave", "Idioma",
        "Revisor 1 (T/R)", "Decisão R1", "Motivo Exclusão R1",
        "Revisor 2 (T/R)", "Decisão R2", "Motivo Exclusão R2",
        "Conflito?", "Resolução", "Decisão Final T/R",
        "Revisor 1 (TC)", "Decisão TC R1", "Motivo Exclusão TC R1",
        "Revisor 2 (TC)", "Decisão TC R2", "Motivo Exclusão TC R2",
        "Conflito TC?", "Resolução TC", "Decisão Final TC",
        "Etapa PRISMA", "Observações"
    ]
    
    # Column widths
    widths = [6, 12, 18, 40, 30, 8, 25, 50, 30, 8,
              16, 14, 25, 16, 14, 25,
              10, 20, 16,
              16, 14, 25, 16, 14, 25,
              10, 20, 16,
              18, 30]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    
    # Title row
    ws.merge_cells("A1:AD1")
    ws["A1"].value = "PLANILHA DE TRIAGEM — REVISÃO SISTEMÁTICA (Screening por Título/Resumo e Texto Completo)"
    ws["A1"].font = Font(name="Segoe UI", size=14, bold=True, color=COLORS["white"])
    ws["A1"].fill = fill("indigo")
    ws["A1"].alignment = ALIGN_CENTER
    ws.row_dimensions[1].height = 38
    
    # Color groups for headers
    color_groups = {
        range(0, 10): "slate_800",        # Bibliographic
        range(10, 19): "warning",          # Screening T/R
        range(19, 28): "success",          # Screening TC
        range(28, 30): "indigo_mid",       # PRISMA + obs
    }
    
    # Header row
    row = 2
    for i, h in enumerate(headers):
        col = get_column_letter(i + 1)
        cell = ws[f"{col}{row}"]
        cell.value = h
        cell.font = FONT_HEADER_COL
        cell.alignment = ALIGN_CENTER
        cell.border = THIN_BORDER
        for r, color in color_groups.items():
            if i in r:
                cell.fill = fill(color)
                break
    ws.row_dimensions[row].height = 35
    
    # Data validations
    decision_dv = DataValidation(type="list", formula1='"Incluir,Excluir,Dúvida"', allow_blank=True)
    conflict_dv = DataValidation(type="list", formula1='"Sim,Não"', allow_blank=True)
    final_dv = DataValidation(type="list", formula1='"Incluído,Excluído"', allow_blank=True)
    prisma_dv = DataValidation(type="list", formula1='"Identificação,Triagem,Elegibilidade,Incluído,Excluído — duplicata,Excluído — T/R,Excluído — TC,Não recuperado"', allow_blank=True)
    
    ws.add_data_validation(decision_dv)
    ws.add_data_validation(conflict_dv)
    ws.add_data_validation(final_dv)
    ws.add_data_validation(prisma_dv)
    
    # Conditional formatting
    red_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    green_fill = PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid")
    yellow_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    
    # Prepare 200 blank rows with formatting
    for r in range(3, 203):
        for i in range(len(headers)):
            col = get_column_letter(i + 1)
            cell = ws[f"{col}{r}"]
            cell.font = FONT_INPUT
            cell.alignment = ALIGN_LEFT
            cell.border = THIN_BORDER
            cell.protection = Protection(locked=False)
            if r % 2 == 0:
                cell.fill = fill("slate_50")
        
        ws[f"A{r}"].value = r - 2
        ws[f"A{r}"].font = Font(name="Consolas", size=10, color=COLORS["slate_400"])
        ws[f"A{r}"].alignment = ALIGN_CENTER
        
        # Apply validations to specific columns
        for col_letter in ["L", "O"]:  # Decision R1, R2 (T/R)
            decision_dv.add(ws[f"{col_letter}{r}"])
        for col_letter in ["U", "X"]:  # Decision R1, R2 (TC)
            decision_dv.add(ws[f"{col_letter}{r}"])
        for col_letter in ["Q", "Z"]:  # Conflict
            conflict_dv.add(ws[f"{col_letter}{r}"])
        for col_letter in ["S", "AB"]:  # Final decision
            final_dv.add(ws[f"{col_letter}{r}"])
        prisma_dv.add(ws[f"AC{r}"])  # PRISMA stage
    
    # Conditional formatting for decision columns
    for col_letter in ["L", "O", "S", "U", "X", "AB"]:
        col_range = f"{col_letter}3:{col_letter}202"
        ws.conditional_formatting.add(col_range, CellIsRule(operator="equal", formula=['"Incluir"'], fill=green_fill))
        ws.conditional_formatting.add(col_range, CellIsRule(operator="equal", formula=['"Incluído"'], fill=green_fill))
        ws.conditional_formatting.add(col_range, CellIsRule(operator="equal", formula=['"Excluir"'], fill=red_fill))
        ws.conditional_formatting.add(col_range, CellIsRule(operator="equal", formula=['"Excluído"'], fill=red_fill))
        ws.conditional_formatting.add(col_range, CellIsRule(operator="equal", formula=['"Dúvida"'], fill=yellow_fill))
    
    # Freeze panes
    ws.freeze_panes = "E3"
    
    # Auto-filter
    ws.auto_filter.ref = f"A2:AD202"


def build_planilha_extracao(wb):
    """Planilha de Extração de Dados."""
    ws = wb.create_sheet("📋 Planilha Extração")
    ws.sheet_properties.tabColor = COLORS["tab_extracao"]
    ws.sheet_view.showGridLines = True
    
    headers = [
        "ID", "Autor (Ano)", "DOI", "Base",
        "Título", "Autores", "Ano", "Periódico", "Resumo", "Palavras-Chave",
        "Objetivo", "Desenho do Estudo", "Tamanho da Amostra",
        "Método Estatístico", "Principais Resultados", "Limitações",
        "País", "Financiamento",
        "Qualidade (Ferramenta)", "Classificação Risco",
        "Extrator", "Data Extração", "Verificador", "Observações"
    ]
    
    widths = [6, 20, 18, 12,
              40, 30, 8, 25, 50, 25,
              35, 20, 14,
              25, 40, 30,
              14, 20,
              20, 18,
              16, 14, 16, 30]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    
    # Title
    last_col = get_column_letter(len(headers))
    ws.merge_cells(f"A1:{last_col}1")
    ws["A1"].value = "PLANILHA DE EXTRAÇÃO DE DADOS — REVISÃO SISTEMÁTICA"
    ws["A1"].font = Font(name="Segoe UI", size=14, bold=True, color=COLORS["white"])
    ws["A1"].fill = fill("prisma_inc")
    ws["A1"].alignment = ALIGN_CENTER
    ws.row_dimensions[1].height = 38
    
    # Header row
    row = 2
    color_groups = {
        range(0, 4): "slate_800",
        range(4, 10): "indigo_mid",
        range(10, 18): "success",
        range(18, 20): "tab_quality",
        range(20, 24): "slate_600",
    }
    for i, h in enumerate(headers):
        col = get_column_letter(i + 1)
        cell = ws[f"{col}{row}"]
        cell.value = h
        cell.font = FONT_HEADER_COL
        cell.alignment = ALIGN_CENTER
        cell.border = THIN_BORDER
        for r, color in color_groups.items():
            if i in r:
                cell.fill = fill(color)
                break
    ws.row_dimensions[row].height = 35
    
    # Data validations
    risk_dv = DataValidation(type="list", formula1='"Alto risco,Moderado risco,Baixo risco,Não avaliado"', allow_blank=True)
    ws.add_data_validation(risk_dv)
    
    # 100 blank rows
    for r in range(3, 103):
        for i in range(len(headers)):
            col = get_column_letter(i + 1)
            cell = ws[f"{col}{r}"]
            cell.font = FONT_INPUT
            cell.alignment = ALIGN_LEFT
            cell.border = THIN_BORDER
            cell.protection = Protection(locked=False)
            if r % 2 == 0:
                cell.fill = fill("slate_50")
        ws[f"A{r}"].value = r - 2
        ws[f"A{r}"].font = Font(name="Consolas", size=10, color=COLORS["slate_400"])
        ws[f"A{r}"].alignment = ALIGN_CENTER
        risk_dv.add(ws[f"T{r}"])
    
    ws.freeze_panes = "E3"
    ws.auto_filter.ref = f"A2:{last_col}102"


def build_planilha_qualidade(wb):
    """Planilha de Avaliação de Qualidade."""
    ws = wb.create_sheet("🏅 Planilha Qualidade")
    ws.sheet_properties.tabColor = COLORS["tab_qualidade"]
    ws.sheet_view.showGridLines = True
    
    # ROB 2 domains as default (most common)
    headers = [
        "ID", "Autor (Ano)", "Ferramenta",
        "D1: Randomização", "D2: Desvios Intervenção", "D3: Dados Faltantes",
        "D4: Mensuração Desfecho", "D5: Seleção Resultado",
        "Julgamento Global", "Classificação", "Avaliador", "Observações"
    ]
    
    widths = [6, 20, 22, 18, 22, 18, 22, 18, 18, 16, 16, 30]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    
    last_col = get_column_letter(len(headers))
    ws.merge_cells(f"A1:{last_col}1")
    ws["A1"].value = "AVALIAÇÃO DE QUALIDADE METODOLÓGICA — RISCO DE VIÉS"
    ws["A1"].font = Font(name="Segoe UI", size=14, bold=True, color=COLORS["white"])
    ws["A1"].fill = fill("tab_quality")
    ws["A1"].alignment = ALIGN_CENTER
    ws.row_dimensions[1].height = 38
    
    # Sub-header with instructions
    ws.merge_cells(f"A2:{last_col}2")
    ws["A2"].value = "Domínios ROB 2 pré-configurados. Adapte as colunas D-H conforme a ferramenta escolhida na Etapa 8. Use os mesmos rótulos para todas as linhas."
    ws["A2"].font = FONT_HINT
    ws["A2"].fill = fill("warning_wash")
    ws["A2"].alignment = ALIGN_LEFT
    ws.row_dimensions[2].height = 24
    
    row = 3
    for i, h in enumerate(headers):
        col = get_column_letter(i + 1)
        cell = ws[f"{col}{row}"]
        cell.value = h
        cell.font = FONT_HEADER_COL
        cell.fill = fill("tab_quality")
        cell.alignment = ALIGN_CENTER
        cell.border = THIN_BORDER
    ws.row_dimensions[row].height = 35
    
    # Validations
    risk_dv = DataValidation(type="list", formula1='"Baixo risco,Alguma preocupação,Alto risco"', allow_blank=True)
    class_dv = DataValidation(type="list", formula1='"Baixo risco,Moderado risco,Alto risco"', allow_blank=True)
    ws.add_data_validation(risk_dv)
    ws.add_data_validation(class_dv)
    
    # Conditional formatting
    red_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    green_fill = PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid")
    yellow_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    
    for r in range(4, 104):
        for i in range(len(headers)):
            col = get_column_letter(i + 1)
            cell = ws[f"{col}{r}"]
            cell.font = FONT_INPUT
            cell.alignment = ALIGN_LEFT if i > 2 else ALIGN_CENTER
            cell.border = THIN_BORDER
            cell.protection = Protection(locked=False)
            if r % 2 == 0:
                cell.fill = fill("slate_50")
        
        ws[f"A{r}"].value = r - 3
        ws[f"A{r}"].font = Font(name="Consolas", size=10, color=COLORS["slate_400"])
        
        # Apply validations to domain columns (D-I)
        for col_letter in ["D", "E", "F", "G", "H", "I"]:
            risk_dv.add(ws[f"{col_letter}{r}"])
        class_dv.add(ws[f"J{r}"])
    
    # Conditional formatting on domain columns
    for col_letter in ["D", "E", "F", "G", "H", "I", "J"]:
        col_range = f"{col_letter}4:{col_letter}103"
        ws.conditional_formatting.add(col_range, CellIsRule(operator="equal", formula=['"Baixo risco"'], fill=green_fill))
        ws.conditional_formatting.add(col_range, CellIsRule(operator="equal", formula=['"Alguma preocupação"'], fill=yellow_fill))
        ws.conditional_formatting.add(col_range, CellIsRule(operator="equal", formula=['"Alto risco"'], fill=red_fill))
        ws.conditional_formatting.add(col_range, CellIsRule(operator="equal", formula=['"Moderado risco"'], fill=yellow_fill))
    
    ws.freeze_panes = "C4"
    ws.auto_filter.ref = f"A3:{last_col}103"


def build_prisma_flow(wb):
    """Planilha PRISMA Flow Data."""
    ws = wb.create_sheet("🔀 PRISMA Flow")
    ws.sheet_properties.tabColor = COLORS["tab_prisma"]
    ws.sheet_view.showGridLines = False
    set_col_widths(ws, {"A": 4, "B": 6, "C": 38, "D": 50, "E": 16, "F": 20, "G": 30})
    
    # Title
    ws.merge_cells("B1:G1")
    ws["B1"].value = "DADOS PARA O FLUXOGRAMA PRISMA 2020"
    ws["B1"].font = Font(name="Segoe UI", size=14, bold=True, color=COLORS["white"])
    ws["B1"].fill = fill("tab_prisma")
    ws["B1"].alignment = ALIGN_CENTER
    ws.row_dimensions[1].height = 38
    
    ws.merge_cells("B2:G2")
    ws["B2"].value = "Preencha a coluna N conforme o processo de triagem avança. Referência: Page et al. (2021) — PRISMA 2020, BMJ 372:n71."
    ws["B2"].font = FONT_HINT
    ws["B2"].fill = fill("info_wash")
    ws["B2"].alignment = ALIGN_LEFT
    ws.row_dimensions[2].height = 24
    
    row = 4
    # Headers
    headers = ["", "Etapa PRISMA", "Descrição", "N", "Observações"]
    for i, h in enumerate(headers):
        col = get_column_letter(i + 2)  # Start at B
        cell = ws[f"{col}{row}"]
        cell.value = h
        cell.font = FONT_HEADER_COL
        cell.fill = fill("tab_prisma")
        cell.alignment = ALIGN_CENTER
        cell.border = THIN_BORDER
    ws.row_dimensions[row].height = 30
    row += 1
    
    # PRISMA items with stage colors
    prisma_items = [
        # Identification
        ("prisma_id", "prisma_id_bg", "IDENTIFICAÇÃO", [
            "Registros identificados via BDTD",
            "Registros identificados via SciELO",
            "Registros identificados via Scopus",
            "Registros identificados via OpenAlex",
            "Registros de outras bases (manual)",
            "Total de registros identificados",
            "Registros removidos antes da triagem (duplicatas)",
            "Registros removidos por automação (inelegíveis)",
            "Registros removidos por outros motivos",
        ]),
        # Screening
        ("prisma_scr", "prisma_scr_bg", "TRIAGEM", [
            "Registros triados (título e resumo)",
            "Registros excluídos na triagem T/R",
        ]),
        # Eligibility
        ("prisma_elig", "prisma_elig_bg", "ELEGIBILIDADE", [
            "Relatórios buscados para avaliação (texto completo)",
            "Relatórios não recuperados",
            "Relatórios avaliados para elegibilidade",
            "Excluídos — motivo 1: (especifique)",
            "Excluídos — motivo 2: (especifique)",
            "Excluídos — motivo 3: (especifique)",
            "Excluídos — outros motivos",
        ]),
        # Included
        ("prisma_inc", "prisma_inc_bg", "INCLUÍDOS", [
            "Estudos incluídos na revisão",
            "Relatórios de estudos incluídos",
            "Estudos incluídos na síntese quantitativa (meta-análise)",
            "Estudos incluídos na síntese qualitativa/narrativa",
        ]),
    ]
    
    item_num = 1
    for color_key, bg_key, stage_name, items in prisma_items:
        # Stage header
        ws.merge_cells(f"B{row}:F{row}")
        cell = ws[f"B{row}"]
        cell.value = f"  {stage_name}"
        cell.font = Font(name="Segoe UI", size=11, bold=True, color=COLORS[color_key])
        cell.fill = fill(bg_key)
        cell.alignment = ALIGN_LEFT
        ws[f"G{row}"].fill = fill(bg_key)
        ws.row_dimensions[row].height = 28
        row += 1
        
        for desc in items:
            vals = [str(item_num), stage_name, desc, "", ""]
            for i, v in enumerate(vals):
                col = get_column_letter(i + 2)
                cell = ws[f"{col}{row}"]
                cell.value = v
                cell.font = FONT_INPUT if i >= 3 else FONT_BODY
                cell.alignment = ALIGN_LEFT
                cell.border = THIN_BORDER
                if i >= 3:
                    cell.protection = Protection(locked=False)
                    cell.fill = fill("slate_50")
                if row % 2 == 0 and i < 3:
                    cell.fill = fill("slate_50")
            ws.row_dimensions[row].height = 26
            item_num += 1
            row += 1
    
    # Summary formulas
    row += 1
    ws.merge_cells(f"B{row}:G{row}")
    ws[f"B{row}"].value = "ℹ️  Preencha a coluna N à medida que o processo de revisão avança. Use esses dados para montar o diagrama PRISMA 2020."
    ws[f"B{row}"].font = FONT_HINT
    ws[f"B{row}"].fill = fill("warning_wash")
    ws[f"B{row}"].alignment = ALIGN_LEFT


# ═══════════════════════════════════════════════════════════════
# MAIN — Generate workbook
# ═══════════════════════════════════════════════════════════════

def main():
    print("=" * 60)
    print("  GERADOR DO FORMULÁRIO DE DESENHO DE PESQUISA")
    print("  Revisão Sistemática da Literatura")
    print("=" * 60)
    print()
    
    wb = openpyxl.Workbook()
    
    print("[1/13] Criando aba de instruções...")
    build_instructions(wb)
    
    print("[2/13] Criando aba 1: Protocolo...")
    build_protocolo(wb)
    
    print("[3/13] Criando aba 2: Questão PICO...")
    build_pico(wb)
    
    print("[4/13] Criando aba 3: Elegibilidade...")
    build_elegibilidade(wb)
    
    print("[5/13] Criando aba 4: Estratégia de Busca...")
    build_estrategia_busca(wb)
    
    print("[6/13] Criando aba 5: Fontes de Dados...")
    build_fontes(wb)
    
    print("[7/13] Criando aba 6: Triagem...")
    build_triagem(wb)
    
    print("[8/13] Criando aba 7: Extração de Dados...")
    build_extracao_dados(wb)
    
    print("[9/13] Criando aba 8: Qualidade...")
    build_qualidade(wb)
    
    print("[10/13] Criando aba 9: Síntese...")
    build_sintese(wb)
    
    print("[11/13] Criando Planilha de Triagem...")
    build_planilha_triagem(wb)
    
    print("[12/13] Criando Planilha de Extração...")
    build_planilha_extracao(wb)
    
    print("[13/13] Criando Planilha de Qualidade e PRISMA Flow...")
    build_planilha_qualidade(wb)
    build_prisma_flow(wb)
    
    # Output path
    output_dir = os.path.dirname(os.path.abspath(__file__))
    output_path = os.path.join(output_dir, "Formulario_Desenho_Pesquisa_RSL.xlsx")
    
    wb.save(output_path)
    
    print()
    print("[OK] Arquivo gerado com sucesso!")
    print(f"   Arquivo: {output_path}")
    print()
    print(f"   Total de abas: {len(wb.sheetnames)}")
    for i, name in enumerate(wb.sheetnames):
        try:
            print(f"     {i+1}. {name}")
        except UnicodeEncodeError:
            # Fallback for terminals that don't support emojis
            clean_name = name.encode('ascii', errors='ignore').decode('ascii').strip()
            print(f"     {i+1}. {clean_name}")
    print()
    print("=" * 60)


if __name__ == "__main__":
    main()
