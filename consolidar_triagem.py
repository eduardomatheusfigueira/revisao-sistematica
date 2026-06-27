"""
Consolidador e Preparador de Triagem para Revisão Sistemática
==============================================================
Lê os dados coletados de todas as bases (OpenAlex, SciELO, BDTD, Scopus),
consolida em formato unificado, realiza deduplicação por DOI e similaridade
de título, e popula a aba "Planilha Triagem" do formulário Excel.

Também atualiza a aba "PRISMA Flow" com os números reais.

Uso:
  python consolidar_triagem.py
"""

import os
import csv
import json
import re
import sys
from datetime import datetime
from difflib import SequenceMatcher
from collections import Counter, defaultdict

import openpyxl
from openpyxl.styles import (
    Font, Alignment, Border, Side, PatternFill, Protection
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule


# ═══════════════════════════════════════════════════════════════
# CONFIGURATION
# ═══════════════════════════════════════════════════════════════

FORMULARIO_PATH = "Formulario_Desenho_Pesquisa_RSL.xlsx"

# Data sources — map base name to (CSV path, ID column name)
DATA_SOURCES = {
    "OpenAlex": {
        "csv": "openalex_outputs/openalex_clean_data.csv",
        "id_col": "ID OpenAlex",
    },
    "SciELO": {
        "csv": "scielo_outputs/scielo_clean_data.csv",
        "id_col": "ID SciELO (PID)",
    },
    "BDTD": {
        "csv": "data_outputs/bdtd_clean_data.csv",
        "id_col": "ID BDTD",
    },
    # Scopus can be added later
}

# Deduplication thresholds
DOI_DEDUP = True
TITLE_SIMILARITY_THRESHOLD = 0.90  # Levenshtein-like ratio for fuzzy match

# Design System colors (matching the formulário)
COLORS = {
    "indigo_dark":   "1E1B4B",
    "indigo":        "4338CA",
    "indigo_mid":    "6366F1",
    "indigo_light":  "A5B4FC",
    "indigo_pale":   "E0E7FF",
    "indigo_wash":   "EEF2FF",
    "slate_900":     "0F172A",
    "slate_700":     "334155",
    "slate_600":     "475569",
    "slate_200":     "E2E8F0",
    "slate_100":     "F1F5F9",
    "slate_50":      "F8FAFC",
    "white":         "FFFFFF",
    "success":       "059669",
    "success_light": "D1FAE5",
    "warning":       "D97706",
    "warning_light": "FEF3C7",
    "error":         "DC2626",
    "error_light":   "FEE2E2",
    "amber_500":     "F59E0B",
    "amber_light":   "FEF3C7",
}


# ═══════════════════════════════════════════════════════════════
# DATA LOADING
# ═══════════════════════════════════════════════════════════════

def load_csv_records(csv_path, base_name, id_col):
    """Load records from a CSV file and normalize to a standard schema."""
    records = []
    if not os.path.exists(csv_path):
        print(f"  [SKIP] {base_name}: arquivo '{csv_path}' não encontrado.")
        return records

    with open(csv_path, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            record = {
                "base": base_name,
                "id_original": row.get(id_col, "").strip(),
                "doi": normalize_doi(row.get("DOI", "")),
                "titulo": clean_text(row.get("Título", row.get("T\u00edtulo", ""))),
                "autores": clean_text(row.get("Autores", "")),
                "ano": row.get("Ano de Publicação", row.get("Ano de Publica\u00e7\u00e3o", "")),
                "periodico": clean_text(row.get("Periódico (Nome)", row.get("Peri\u00f3dico (Nome)", ""))),
                "resumo": clean_text(row.get("Resumo", "")),
                "palavras_chave": clean_text(row.get("Palavras-chave / Conceitos",
                                    row.get("Fontes / Bases Originais", ""))),
                "idioma": row.get("Idioma", "").strip(),
                "url": row.get("Link / URL", "").strip(),
                "pdf_url": row.get("PDF URL", "").strip(),
            }
            if record["titulo"] and record["titulo"] != "N/A":
                records.append(record)

    return records


# ═══════════════════════════════════════════════════════════════
# TEXT UTILITIES
# ═══════════════════════════════════════════════════════════════

def clean_text(text):
    """Clean and normalize text."""
    if not text or text.strip() in ("", "N/A", "None"):
        return ""
    text = re.sub(r'<[^>]+>', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


def normalize_doi(doi):
    """Normalize a DOI for comparison."""
    if not doi or doi.strip() in ("", "N/A", "None"):
        return ""
    doi = doi.strip().lower()
    # Remove common prefixes
    doi = re.sub(r'^https?://doi\.org/', '', doi)
    doi = re.sub(r'^https?://dx\.doi\.org/', '', doi)
    doi = re.sub(r'^doi:', '', doi)
    return doi.strip()


def normalize_title(title):
    """Normalize a title for fuzzy comparison."""
    if not title:
        return ""
    t = title.lower().strip()
    # Remove punctuation
    t = re.sub(r'[^\w\s]', '', t)
    # Collapse whitespace
    t = re.sub(r'\s+', ' ', t)
    return t.strip()


def title_similarity(t1, t2):
    """Compute similarity ratio between two normalized titles."""
    return SequenceMatcher(None, t1, t2).ratio()


# ═══════════════════════════════════════════════════════════════
# DEDUPLICATION ENGINE
# ═══════════════════════════════════════════════════════════════

def deduplicate_records(all_records):
    """
    Remove duplicates using a two-phase approach:
    1. Exact DOI match
    2. Fuzzy title similarity for records without DOI
    
    Returns: (unique_records, duplicates_removed, duplicate_pairs)
    """
    print("\n[DEDUPLICAÇÃO] Iniciando processo...")
    
    unique = []
    duplicates_removed = []
    duplicate_pairs = []  # (removed_record, kept_record_index)
    
    # Index by DOI
    doi_index = {}  # doi -> index in unique list
    
    # Phase 1: DOI-based dedup
    doi_dupes = 0
    for record in all_records:
        doi = record["doi"]
        if doi and doi in doi_index:
            # Duplicate by DOI
            kept_idx = doi_index[doi]
            duplicates_removed.append(record)
            duplicate_pairs.append((record, unique[kept_idx]))
            doi_dupes += 1
            # Merge: keep the one with more info (longer abstract)
            if len(record.get("resumo", "")) > len(unique[kept_idx].get("resumo", "")):
                old_base = unique[kept_idx]["base"]
                unique[kept_idx].update(record)
                unique[kept_idx]["base"] = f"{old_base}; {record['base']}"
        elif doi:
            doi_index[doi] = len(unique)
            unique.append(record)
        else:
            unique.append(record)
    
    print(f"  Fase 1 (DOI exato): {doi_dupes} duplicatas removidas")
    
    # Phase 2: Title similarity for records without DOI
    title_dupes = 0
    no_doi_indices = [i for i, r in enumerate(unique) if not r["doi"]]
    has_doi_indices = [i for i, r in enumerate(unique) if r["doi"]]
    
    # Compare no-DOI records against all DOI records
    to_remove = set()
    for i in no_doi_indices:
        if i in to_remove:
            continue
        norm_title_i = normalize_title(unique[i]["titulo"])
        if not norm_title_i:
            continue
        
        # Compare against DOI records
        for j in has_doi_indices:
            norm_title_j = normalize_title(unique[j]["titulo"])
            if norm_title_j and title_similarity(norm_title_i, norm_title_j) >= TITLE_SIMILARITY_THRESHOLD:
                to_remove.add(i)
                duplicate_pairs.append((unique[i], unique[j]))
                title_dupes += 1
                break
        
        if i not in to_remove:
            # Compare against other no-DOI records
            for j in no_doi_indices:
                if j <= i or j in to_remove:
                    continue
                norm_title_j = normalize_title(unique[j]["titulo"])
                if norm_title_j and title_similarity(norm_title_i, norm_title_j) >= TITLE_SIMILARITY_THRESHOLD:
                    to_remove.add(j)
                    duplicate_pairs.append((unique[j], unique[i]))
                    title_dupes += 1
    
    print(f"  Fase 2 (similaridade de título ≥ {TITLE_SIMILARITY_THRESHOLD:.0%}): {title_dupes} duplicatas removidas")
    
    # Remove title-based duplicates
    for idx in sorted(to_remove, reverse=True):
        duplicates_removed.append(unique[idx])
    
    final_unique = [r for i, r in enumerate(unique) if i not in to_remove]
    
    total_dupes = doi_dupes + title_dupes
    print(f"  Total de duplicatas removidas: {total_dupes}")
    print(f"  Registros únicos restantes: {len(final_unique)}")
    
    return final_unique, duplicates_removed, duplicate_pairs


# ═══════════════════════════════════════════════════════════════
# AUTOMATED PRE-SCREENING (ELIGIBILITY FILTERS)
# ═══════════════════════════════════════════════════════════════

def auto_screen_records(records):
    """
    Apply automated pre-screening filters based on eligibility criteria:
    1. Language filter (keep: en, pt, es)
    2. Year filter (keep: 2020-2026)
    3. Remove records without title or abstract
    
    Returns: (eligible_records, ineligible_records, reasons_counter)
    """
    print("\n[PRÉ-TRIAGEM AUTOMATIZADA] Aplicando filtros de elegibilidade...")
    
    eligible = []
    ineligible = []
    reasons = Counter()
    
    allowed_languages = {"en", "pt", "es", "por", "eng", "spa", "português", "english", "spanish", ""}
    
    for record in records:
        # Check language
        lang = record.get("idioma", "").strip().lower()
        if lang and lang not in allowed_languages:
            reasons["Idioma não elegível"] += 1
            record["motivo_exclusao_auto"] = f"Idioma: {record['idioma']}"
            ineligible.append(record)
            continue
        
        # Check year
        try:
            year = int(record.get("ano", 0))
            if year < 2020 or year > 2026:
                reasons["Fora do recorte temporal"] += 1
                record["motivo_exclusao_auto"] = f"Ano: {year}"
                ineligible.append(record)
                continue
        except (ValueError, TypeError):
            pass  # Keep records with unparseable years
        
        # Check minimum content (must have title)
        if not record.get("titulo"):
            reasons["Sem título"] += 1
            record["motivo_exclusao_auto"] = "Sem título"
            ineligible.append(record)
            continue
        
        eligible.append(record)
    
    print(f"  Registros elegíveis: {len(eligible)}")
    print(f"  Registros excluídos automaticamente: {len(ineligible)}")
    for reason, count in reasons.most_common():
        print(f"    - {reason}: {count}")
    
    return eligible, ineligible, reasons


# ═══════════════════════════════════════════════════════════════
# EXCEL POPULATION
# ═══════════════════════════════════════════════════════════════

def populate_screening_sheet(wb, records):
    """Populate the screening spreadsheet tab with consolidated records."""
    
    sheet_name = "\U0001f4ca Planilha Triagem"
    ws = wb[sheet_name]
    
    print(f"\n[PLANILHA TRIAGEM] Populando {len(records)} registros...")
    
    # Design styles
    font_body = Font(name="Segoe UI", size=9, color=COLORS["slate_900"])
    font_body_small = Font(name="Segoe UI", size=8, color=COLORS["slate_600"])
    align_left = Alignment(horizontal='left', vertical='center', wrap_text=False)
    align_center = Alignment(horizontal='center', vertical='center')
    align_wrap = Alignment(horizontal='left', vertical='top', wrap_text=True)
    
    fill_zebra_even = PatternFill(start_color=COLORS["slate_50"], end_color=COLORS["slate_50"], fill_type="solid")
    fill_zebra_odd = PatternFill(start_color=COLORS["white"], end_color=COLORS["white"], fill_type="solid")
    
    border_light = Border(
        left=Side(style='thin', color=COLORS["slate_200"]),
        right=Side(style='thin', color=COLORS["slate_200"]),
        top=Side(style='thin', color=COLORS["slate_200"]),
        bottom=Side(style='thin', color=COLORS["slate_200"])
    )
    
    # Clear existing data rows (starting from row 3)
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row):
        for cell in row:
            cell.value = None
    
    # Column mapping (matches the header in row 2)
    # A=ID, B=Base, C=DOI, D=Título, E=Autores, F=Ano, G=Periódico,
    # H=Resumo, I=Palavras-Chave, J=Idioma, K-S=Revisores..., T=Observações
    
    for idx, record in enumerate(records, start=1):
        row_num = idx + 2  # Data starts at row 3
        fill = fill_zebra_even if idx % 2 == 0 else fill_zebra_odd
        
        values = [
            idx,                              # A: ID
            record["base"],                   # B: Base
            record.get("doi", ""),            # C: DOI
            record.get("titulo", ""),         # D: Título
            record.get("autores", ""),        # E: Autores
            record.get("ano", ""),            # F: Ano
            record.get("periodico", ""),      # G: Periódico
            record.get("resumo", ""),         # H: Resumo
            record.get("palavras_chave", ""), # I: Palavras-Chave
            record.get("idioma", ""),         # J: Idioma
        ]
        
        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row=row_num, column=col_idx, value=value)
            cell.font = font_body
            cell.border = border_light
            cell.fill = fill
            
            if col_idx in (1, 6, 10):  # ID, Ano, Idioma
                cell.alignment = align_center
            elif col_idx in (4, 8):  # Título, Resumo
                cell.alignment = align_wrap
            else:
                cell.alignment = align_left
        
        # Set remaining columns (K through AD) with empty styling
        for col_idx in range(11, 31):
            cell = ws.cell(row=row_num, column=col_idx)
            cell.font = font_body
            cell.border = border_light
            cell.fill = fill
            cell.alignment = align_center
        
        # Set PRISMA stage default
        ws.cell(row=row_num, column=29, value="Triagem T/R").font = font_body_small
    
    # Add data validation for decision columns
    decision_list = '"Incluir,Excluir,Dúvida"'
    dv = DataValidation(type="list", formula1=decision_list, allow_blank=True)
    dv.error = "Selecione: Incluir, Excluir ou Dúvida"
    dv.errorTitle = "Decisão inválida"
    dv.prompt = "Selecione a decisão de triagem"
    dv.promptTitle = "Decisão"
    
    last_row = len(records) + 2
    
    # Apply to Decision columns: L (R1), O (R2), U (TC R1), X (TC R2)
    for col_letter in ["L", "O", "U", "X"]:
        dv_range = f"{col_letter}3:{col_letter}{last_row}"
        dv_copy = DataValidation(type="list", formula1=decision_list, allow_blank=True)
        dv_copy.error = dv.error
        dv_copy.errorTitle = dv.errorTitle
        ws.add_data_validation(dv_copy)
        dv_copy.add(f"{col_letter}3:{col_letter}{last_row}")
    
    # Exclusion motives validation
    motives_list = '"Idioma,Fora do escopo,Tipo de documento inadequado,Duplicata,Sem acesso,Qualidade insuficiente,Outro"'
    for col_letter in ["M", "P", "V", "Y"]:
        dv_motive = DataValidation(type="list", formula1=motives_list, allow_blank=True)
        dv_motive.error = "Selecione um motivo de exclusão"
        ws.add_data_validation(dv_motive)
        dv_motive.add(f"{col_letter}3:{col_letter}{last_row}")
    
    # Conflict column formula (compare R1 and R2)
    for row_num in range(3, last_row + 1):
        # Conflict T/R (col Q = 17): IF L != O then "SIM"
        ws.cell(row=row_num, column=17).value = f'=IF(AND(L{row_num}<>"",O{row_num}<>""),IF(L{row_num}<>O{row_num},"SIM","NÃO"),"")'
        ws.cell(row=row_num, column=17).font = font_body
        
        # Final Decision T/R (col S = 19)
        ws.cell(row=row_num, column=19).value = f'=IF(R{row_num}<>"",R{row_num},IF(AND(L{row_num}<>"",O{row_num}<>""),IF(L{row_num}=O{row_num},L{row_num},"CONFLITO"),""))'
        ws.cell(row=row_num, column=19).font = font_body
        
        # Conflict TC (col Z = 26)
        ws.cell(row=row_num, column=26).value = f'=IF(AND(U{row_num}<>"",X{row_num}<>""),IF(U{row_num}<>X{row_num},"SIM","NÃO"),"")'
        ws.cell(row=row_num, column=26).font = font_body
        
        # Final Decision TC (col AB = 28)
        ws.cell(row=row_num, column=28).value = f'=IF(AA{row_num}<>"",AA{row_num},IF(AND(U{row_num}<>"",X{row_num}<>""),IF(U{row_num}=X{row_num},U{row_num},"CONFLITO"),""))'
        ws.cell(row=row_num, column=28).font = font_body
    
    # Conditional formatting: highlight "Incluir" in green, "Excluir" in red, "Dúvida" in yellow
    green_fill = PatternFill(start_color=COLORS["success_light"], end_color=COLORS["success_light"], fill_type="solid")
    red_fill = PatternFill(start_color=COLORS["error_light"], end_color=COLORS["error_light"], fill_type="solid")
    yellow_fill = PatternFill(start_color=COLORS["amber_light"], end_color=COLORS["amber_light"], fill_type="solid")
    
    green_font = Font(name="Segoe UI", size=9, color=COLORS["success"], bold=True)
    red_font = Font(name="Segoe UI", size=9, color=COLORS["error"], bold=True)
    yellow_font = Font(name="Segoe UI", size=9, color=COLORS["warning"], bold=True)
    
    for col_letter in ["L", "O", "S", "U", "X", "AB"]:
        cell_range = f"{col_letter}3:{col_letter}{last_row}"
        ws.conditional_formatting.add(cell_range,
            CellIsRule(operator='equal', formula=['"Incluir"'], fill=green_fill, font=green_font))
        ws.conditional_formatting.add(cell_range,
            CellIsRule(operator='equal', formula=['"Excluir"'], fill=red_fill, font=red_font))
        ws.conditional_formatting.add(cell_range,
            CellIsRule(operator='equal', formula=['"Dúvida"'], fill=yellow_fill, font=yellow_font))
    
    print(f"  [OK] {len(records)} registros inseridos na Planilha Triagem.")
    return len(records)


def update_prisma_flow(wb, stats):
    """Update the PRISMA Flow sheet with real numbers."""
    
    sheet_name = "\U0001f500 PRISMA Flow"
    ws = wb[sheet_name]
    
    print("\n[PRISMA FLOW] Atualizando números...")
    
    font_number = Font(name="Segoe UI", size=11, bold=True, color=COLORS["indigo"])
    
    # Row mapping (from inspection):
    # Row 6 col E: BDTD
    # Row 7 col E: SciELO
    # Row 8 col E: Scopus
    # Row 9 col E: OpenAlex
    # Row 10 col E: Outras bases
    # Row 11 col E: Total identificados
    # Row 12 col E: Duplicatas removidas
    # Row 13 col E: Inelegíveis removidos
    # Row 14 col E: Outros removidos
    # Row 16 col E: Triados (T/R)
    
    ws.cell(row=6, column=5, value=stats.get("bdtd", 0)).font = font_number
    ws.cell(row=7, column=5, value=stats.get("scielo", 0)).font = font_number
    ws.cell(row=8, column=5, value=stats.get("scopus", 0)).font = font_number
    ws.cell(row=9, column=5, value=stats.get("openalex", 0)).font = font_number
    ws.cell(row=10, column=5, value=stats.get("outras", 0)).font = font_number
    ws.cell(row=11, column=5, value=stats.get("total_bruto", 0)).font = font_number
    ws.cell(row=12, column=5, value=stats.get("duplicatas", 0)).font = font_number
    ws.cell(row=13, column=5, value=stats.get("inelegiveis_auto", 0)).font = font_number
    ws.cell(row=14, column=5, value=0).font = font_number  # Outros
    ws.cell(row=16, column=5, value=stats.get("para_triagem", 0)).font = font_number
    
    print(f"  [OK] PRISMA Flow atualizado.")


# ═══════════════════════════════════════════════════════════════
# REPORT GENERATION
# ═══════════════════════════════════════════════════════════════

def generate_consolidation_report(stats, duplicate_pairs, output_path):
    """Generate a markdown report of the consolidation process."""
    
    report = []
    report.append("# Relatório de Consolidação e Deduplicação")
    report.append(f"**Data de execução**: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    report.append("")
    
    report.append("## 1. Registros por Base")
    report.append("| Base | Registros Brutos |")
    report.append("| --- | --- |")
    for base, count in stats.get("por_base", {}).items():
        report.append(f"| {base} | {count} |")
    report.append(f"| **TOTAL** | **{stats['total_bruto']}** |")
    report.append("")
    
    report.append("## 2. Deduplicação")
    report.append(f"- Duplicatas removidas (DOI exato): parte da contagem total")
    report.append(f"- Duplicatas removidas (similaridade de título ≥ {TITLE_SIMILARITY_THRESHOLD:.0%}): parte da contagem total")
    report.append(f"- **Total de duplicatas removidas: {stats['duplicatas']}**")
    report.append(f"- Registros únicos após deduplicação: {stats['pos_dedup']}")
    report.append("")
    
    if duplicate_pairs:
        report.append("### Exemplos de Duplicatas Detectadas (primeiras 20)")
        report.append("| Removido (Base) | Título Removido | Mantido (Base) | DOI |")
        report.append("| --- | --- | --- | --- |")
        for removed, kept in duplicate_pairs[:20]:
            title_short = (removed['titulo'][:60] + '...') if len(removed.get('titulo','')) > 60 else removed.get('titulo','')
            report.append(f"| {removed['base']} | {title_short} | {kept['base']} | {removed.get('doi','')} |")
        report.append("")
    
    report.append("## 3. Pré-Triagem Automatizada")
    report.append(f"- Registros excluídos automaticamente: {stats['inelegiveis_auto']}")
    if stats.get("motivos_exclusao"):
        for reason, count in stats["motivos_exclusao"].items():
            report.append(f"  - {reason}: {count}")
    report.append(f"- **Registros encaminhados para triagem manual: {stats['para_triagem']}**")
    report.append("")
    
    report.append("## 4. Resumo do Fluxo PRISMA")
    report.append("```")
    report.append(f"Identificados:      {stats['total_bruto']}")
    report.append(f"  - Duplicatas:     -{stats['duplicatas']}")
    report.append(f"  - Inelegíveis:    -{stats['inelegiveis_auto']}")
    report.append(f"Para triagem T/R:    {stats['para_triagem']}")
    report.append("```")
    report.append("")
    
    report.append("## 5. Distribuição dos Registros para Triagem")
    
    # Year distribution
    if stats.get("dist_ano"):
        report.append("### Por Ano")
        report.append("| Ano | Quantidade |")
        report.append("| --- | --- |")
        for year in sorted(stats["dist_ano"].keys(), reverse=True):
            report.append(f"| {year} | {stats['dist_ano'][year]} |")
        report.append("")
    
    # Base distribution
    if stats.get("dist_base"):
        report.append("### Por Base de Origem")
        report.append("| Base | Quantidade | Percentual |")
        report.append("| --- | --- | --- |")
        total = stats["para_triagem"]
        for base, count in sorted(stats["dist_base"].items(), key=lambda x: -x[1]):
            pct = (count / total * 100) if total > 0 else 0
            report.append(f"| {base} | {count} | {pct:.1f}% |")
        report.append("")
    
    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(report))
    
    print(f"\n[RELATÓRIO] Salvo em: {output_path}")


# ═══════════════════════════════════════════════════════════════
# DUPLICATES LOG (CSV)
# ═══════════════════════════════════════════════════════════════

def save_duplicates_log(duplicate_pairs, output_path):
    """Save a CSV log of all detected duplicates for audit."""
    if not duplicate_pairs:
        return
    
    with open(output_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["Título Removido", "Base Removido", "DOI Removido",
                         "Título Mantido", "Base Mantido", "DOI Mantido"])
        for removed, kept in duplicate_pairs:
            writer.writerow([
                removed.get("titulo", ""), removed.get("base", ""), removed.get("doi", ""),
                kept.get("titulo", ""), kept.get("base", ""), kept.get("doi", "")
            ])
    
    print(f"[LOG] Log de duplicatas salvo em: {output_path}")


# ═══════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════

def main():
    # Fix Windows terminal encoding
    import sys
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    
    print("=" * 60)
    print("  CONSOLIDADOR E PREPARADOR DE TRIAGEM")
    print("  Revisão Sistemática da Literatura")
    print("=" * 60)
    
    # ─── Step 1: Load all data ───
    print("\n[CARREGAMENTO] Lendo dados de todas as bases...")
    
    all_records = []
    per_base_counts = {}
    
    for base_name, config in DATA_SOURCES.items():
        records = load_csv_records(config["csv"], base_name, config["id_col"])
        per_base_counts[base_name] = len(records)
        all_records.extend(records)
        print(f"  {base_name}: {len(records)} registros carregados")
    
    total_bruto = len(all_records)
    print(f"\n  TOTAL BRUTO: {total_bruto} registros")
    
    if total_bruto == 0:
        print("\n[ERRO] Nenhum registro encontrado. Execute os coletores primeiro.")
        return
    
    # ─── Step 2: Deduplicate ───
    unique_records, duplicates_removed, duplicate_pairs = deduplicate_records(all_records)
    num_duplicates = total_bruto - len(unique_records)
    
    # ─── Step 3: Automated pre-screening ───
    eligible_records, ineligible_records, exclusion_reasons = auto_screen_records(unique_records)
    
    # ─── Step 4: Populate the screening Excel ───
    print(f"\n[EXCEL] Abrindo '{FORMULARIO_PATH}'...")
    wb = openpyxl.load_workbook(FORMULARIO_PATH)
    
    num_populated = populate_screening_sheet(wb, eligible_records)
    
    # ─── Step 5: Update PRISMA Flow ───
    prisma_stats = {
        "bdtd": per_base_counts.get("BDTD", 0),
        "scielo": per_base_counts.get("SciELO", 0),
        "scopus": per_base_counts.get("Scopus", 0),
        "openalex": per_base_counts.get("OpenAlex", 0),
        "outras": 0,
        "total_bruto": total_bruto,
        "duplicatas": num_duplicates,
        "inelegiveis_auto": len(ineligible_records),
        "para_triagem": len(eligible_records),
    }
    
    update_prisma_flow(wb, prisma_stats)
    
    # ─── Step 6: Save Excel ───
    wb.save(FORMULARIO_PATH)
    print(f"\n[OK] Formulário atualizado: {FORMULARIO_PATH}")
    
    # ─── Step 7: Generate reports ───
    stats = {
        "total_bruto": total_bruto,
        "por_base": per_base_counts,
        "duplicatas": num_duplicates,
        "pos_dedup": len(unique_records),
        "inelegiveis_auto": len(ineligible_records),
        "motivos_exclusao": dict(exclusion_reasons),
        "para_triagem": len(eligible_records),
        "dist_ano": Counter(r.get("ano", "N/A") for r in eligible_records),
        "dist_base": Counter(r.get("base", "N/A").split(";")[0].strip() for r in eligible_records),
    }
    
    # Create triagem_outputs directory
    output_dir = "triagem_outputs"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    generate_consolidation_report(stats, duplicate_pairs, os.path.join(output_dir, "relatorio_consolidacao.md"))
    save_duplicates_log(duplicate_pairs, os.path.join(output_dir, "log_duplicatas.csv"))
    
    # Save consolidated data as backup
    consolidated_path = os.path.join(output_dir, "dados_consolidados.json")
    with open(consolidated_path, "w", encoding="utf-8") as f:
        json.dump(eligible_records, f, ensure_ascii=False, indent=2)
    print(f"[BACKUP] Dados consolidados: {consolidated_path}")
    
    # ─── Summary ───
    print("\n" + "=" * 60)
    print("  RESUMO DA CONSOLIDAÇÃO")
    print("=" * 60)
    print(f"  Registros brutos coletados:     {total_bruto}")
    print(f"  Duplicatas removidas:           {num_duplicates}")
    print(f"  Excluídos por inelegibilidade:  {len(ineligible_records)}")
    print(f"  Registros para triagem manual:  {len(eligible_records)}")
    print(f"")
    print(f"  Planilha atualizada: {FORMULARIO_PATH}")
    print(f"  Relatório: {output_dir}/relatorio_consolidacao.md")
    print("=" * 60)


if __name__ == "__main__":
    main()
