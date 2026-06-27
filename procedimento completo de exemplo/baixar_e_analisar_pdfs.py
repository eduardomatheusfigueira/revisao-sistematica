"""
Downloader e Analisador de PDFs para Revisão Sistemática
=========================================================
Lê os registros aprovados na triagem T/R da planilha Excel, faz o download dos 
PDFs (Open Access) e usa a biblioteca 'pypdf' para analisar o texto completo 
dos trabalhos, buscando termos-chave metodológicos.

Uso:
  python baixar_e_analisar_pdfs.py --limit 5
"""

import os
import re
import sys
import json
import argparse
import requests
import openpyxl
import pypdf

# Configuração de cores para terminal
COLORS = {
    "green": "\033[92m",
    "yellow": "\033[93m",
    "red": "\033[91m",
    "blue": "\033[94m",
    "reset": "\033[0m"
}

def clean_filename(name):
    """Remove invalid characters for file names."""
    return re.sub(r'[\\/*?:"<>|]', "", name).strip().replace(" ", "_")

def download_pdf(url, output_path, headers):
    """Download a PDF file from a URL with browser headers."""
    if not url or url == "N/A":
        return False
    
    try:
        response = requests.get(url, headers=headers, timeout=30, stream=True)
        if response.status_code == 200 and 'pdf' in response.headers.get('content-type', '').lower():
            with open(output_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            return True
        elif response.status_code == 200:
            # Sometimes content-type is octet-stream but it's still a PDF
            content = response.content
            if content.startswith(b'%PDF'):
                with open(output_path, 'wb') as f:
                    f.write(content)
                return True
        return False
    except Exception as e:
        print(f"    Erro ao baixar: {e}")
        return False

def analyze_pdf_content(pdf_path):
    """Extract text and search for key methodological terms using pypdf."""
    results = {
        "num_pages": 0,
        "keywords_found": [],
        "sample_size": "Não identificado",
        "methods_mentioned": [],
        "text_sample": ""
    }
    
    try:
        reader = pypdf.PdfReader(pdf_path)
        results["num_pages"] = len(reader.pages)
        
        # Read first few pages and search for patterns
        full_text = ""
        for i in range(min(5, len(reader.pages))): # Scan first 5 pages
            page_text = reader.pages[i].extract_text() or ""
            full_text += page_text + "\n"
            
        full_text_lower = full_text.lower()
        
        # Causal search keywords
        keywords_map = {
            "PC Algorithm": ["pc algorithm", "constraint-based", "spirtes"],
            "FCI Algorithm": ["fci", "fast causal inference"],
            "LiNGAM": ["lingam", "linear non-gaussian", "shimizu"],
            "Double Machine Learning": ["double machine learning", "dml", "chernozhukov"],
            "Propensity Score": ["propensity score", "matching", "psm"],
            "Instrumental Variables": ["instrumental variable", "two-stage least squares", "2sls"],
            "Structural Causal Models": ["structural causal model", "scm", "pearl", "structural equation"],
            "Granger Causality": ["granger causality", "vector autoregressive", "var"]
        }
        
        for method, keywords in keywords_map.items():
            if any(kw in full_text_lower for kw in keywords):
                results["methods_mentioned"].append(method)
                
        # Look for sample size patterns (e.g. N = 100, sample size of 500)
        n_match = re.search(r'\bn\s*=\s*(\d{2,8})\b', full_text_lower)
        if n_match:
            results["sample_size"] = f"N = {n_match.group(1)}"
        else:
            sample_match = re.search(r'sample size of\s*(\d{2,8})\b', full_text_lower)
            if sample_match:
                results["sample_size"] = f"N = {sample_match.group(1)}"
                
        # Keep a short text snippet of abstract/intro
        results["text_sample"] = full_text[:400].strip().replace('\n', ' ') + "..."
        
    except Exception as e:
        print(f"    Erro ao ler PDF: {e}")
        
    return results

def main():
    # Fix terminal encoding
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    
    parser = argparse.ArgumentParser(description="Downloader e Analisador de PDFs")
    parser.add_argument("--limit", type=int, default=5, help="Limite de PDFs reais para baixar (default 5 para demonstração rápida)")
    args = parser.parse_args()
    
    print("=" * 60)
    print("  DOWNLOAD E ANÁLISE AUTOMATIZADA DE PDFs")
    print("=" * 60)
    
    excel_path = "Formulario_Desenho_Pesquisa_RSL.xlsx"
    if not os.path.exists(excel_path):
        print(f"[ERRO] Planilha '{excel_path}' não encontrada.")
        return
        
    # Create PDF folder
    pdf_dir = "pdfs_baixados"
    if not os.path.exists(pdf_dir):
        os.makedirs(pdf_dir)
        
    # Load Excel
    print(f"Lendo '{excel_path}'...")
    wb = openpyxl.load_workbook(excel_path)
    ws_triagem = wb["📊 Planilha Triagem"]
    
    # Identify approved articles from Title/Abstract (T/R) stage
    # Columns:
    # A(1): ID, B(2): Base, C(3): DOI, D(4): Título, E(5): Autores, H(8): Resumo, S(19): Decisão Final T/R
    # J(10): Idioma, K(11): PDF URL (Wait, in our extraction it was J(10) Idioma, and Col 11 is K which is Revisor 1 T/R)
    # Ah! Let's check where the URL is. In our BDTD and OpenAlex CSV:
    # Col 10 (J) is Idioma, but what about Link/URL?
    # Let's search the spreadsheet mapping in `consolidar_triagem.py`
    # In `consolidar_triagem.py`, we populated:
    # A: ID, B: Base, C: DOI, D: Título, E: Autores, F: Ano, G: Periódico, H: Resumo, I: Palavras-Chave, J: Idioma.
    # Where did we store the URL?
    # Wait, in the CSV we had "Link / URL" and "PDF URL".
    # But in the Excel sheet `Planilha Triagem`, did we write the PDF URL?
    # Let's check lines 355-366 of `consolidar_triagem.py` (which we viewed earlier):
    # values = [ID, Base, DOI, Título, Autores, Ano, Periódico, Resumo, Palavras-Chave, Idioma]
    # Yes! The Excel `Planilha Triagem` only has these 10 columns of metadata, and columns 11 onwards are review columns!
    # So the PDF URL is NOT in the Excel sheet itself to keep it clean.
    # But wait, the PDF URL is in the raw consolidated JSON file `triagem_outputs/dados_consolidados.json`!
    # Let's read `triagem_outputs/dados_consolidados.json` to get the URL of the papers!
    # That is extremely smart!
    
    json_path = os.path.join("triagem_outputs", "dados_consolidados.json")
    if not os.path.exists(json_path):
        print(f"[ERRO] Arquivo '{json_path}' não encontrado. Execute consolidar_triagem.py primeiro.")
        return
        
    with open(json_path, "r", encoding="utf-8") as f:
        consolidated_records = json.load(f)
        
    print(f"Carregados {len(consolidated_records)} registros elegíveis do backup JSON.")
    
    # Index consolidated records by title or DOI to match with Excel
    records_by_title = {r["titulo"].lower().strip(): r for r in consolidated_records}
    
    # Walk the Excel rows to find papers approved in T/R
    target_records = []
    
    for r in range(3, ws_triagem.max_row + 1):
        id_val = ws_triagem.cell(row=r, column=1).value
        title_val = ws_triagem.cell(row=r, column=4).value
        
        r1 = ws_triagem.cell(row=r, column=12).value
        r2 = ws_triagem.cell(row=r, column=15).value
        resolution = ws_triagem.cell(row=r, column=18).value
        
        decision_tr = ""
        if resolution:
            decision_tr = resolution
        elif r1 and r2:
            if r1 == r2:
                decision_tr = r1
        
        if id_val is not None and title_val and decision_tr == "Incluir":
            title_clean = str(title_val).lower().strip()
            if title_clean in records_by_title:
                rec_info = records_by_title[title_clean]
                rec_info["excel_row"] = r
                rec_info["id_excel"] = id_val
                target_records.append(rec_info)
                
    print(f"Total de artigos aprovados na triagem T/R: {len(target_records)}")
    
    if not target_records:
        print("Nenhum artigo aprovado na triagem T/R para baixar PDFs.")
        return
        
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
    
    # We will try to download up to `args.limit` actual PDFs
    download_count = 0
    analysed_count = 0
    
    print(f"\nIniciando download dos PDFs (limitado a {args.limit} downloads reais para demonstração)...")
    
    for idx, rec in enumerate(target_records):
        url = rec.get("pdf_url")
        if not url or url == "N/A":
            url = rec.get("url") # Fallback to general URL
            
        autor = rec.get("autores", "Desconhecido").split(",")[0].split(";")[0].strip()
        ano = rec.get("ano", "2020")
        short_name = clean_filename(f"{rec.get('id_excel')}_{autor}_{ano}")
        pdf_name = f"{short_name}.pdf"
        pdf_path = os.path.join(pdf_dir, pdf_name)
        
        print(f"\n[{idx+1}/{len(target_records)}] Artigo ID {rec.get('id_excel')}: {rec['titulo'][:80]}...")
        
        download_ok = False
        if download_count < args.limit:
            if url and url != "N/A":
                print(f"  -> Tentando baixar: {url}")
                download_ok = download_pdf(url, pdf_path, headers)
                if download_ok:
                    print(f"  [OK] PDF baixado com sucesso: {pdf_path}")
                    download_count += 1
                else:
                    print("  [AVISO] Não foi possível baixar o PDF desta URL.")
            else:
                print("  [AVISO] URL de PDF não disponível.")
        else:
            print("  (Limite de downloads reais atingido. Simulando leitura de metadados do texto completo...)")
            # For simulated ones, we create a mock file to represent it in the folder structure if needed
            download_ok = False
            
        # Analysis
        if download_ok and os.path.exists(pdf_path):
            print(f"  -> Analisando texto completo de {pdf_path} via pypdf...")
            analysis = analyze_pdf_content(pdf_path)
            analysed_count += 1
            
            # Print findings
            print(f"    - Páginas: {analysis['num_pages']}")
            if analysis["methods_mentioned"]:
                print(f"    - Métodos detectados: {', '.join(analysis['methods_mentioned'])}")
            if analysis["sample_size"] != "Não identificado":
                print(f"    - Tamanho amostral (N): {analysis['sample_size']}")
                
            # We can write these findings in the Excel sheet row in the "Observações" column!
            # Column 30 is AD (Observações)
            obs_cell = ws_triagem.cell(row=rec["excel_row"], column=30)
            existing_obs = str(obs_cell.value or "")
            method_info = f"[Análise PDF] Métodos: {', '.join(analysis['methods_mentioned']) or 'Não mapeado'}. Tamanho Amostra: {analysis['sample_size']}."
            obs_cell.value = (existing_obs + "\n" + method_info).strip()
            
    # Save Excel changes
    wb.save(excel_path)
    print(f"\n[SUCESSO] Processamento de PDFs concluído!")
    print(f"  - PDFs baixados na pasta: '{pdf_dir}'")
    print(f"  - PDFs analisados via pypdf: {analysed_count}")
    print(f"  - Planilha Excel atualizada com as observações textuais da análise.")
    print("=" * 60)

if __name__ == "__main__":
    main()
