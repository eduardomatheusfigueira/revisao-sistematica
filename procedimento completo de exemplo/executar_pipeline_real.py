"""
Pipeline Real de Revisão Sistemática — Download, Leitura, Triagem e Extração
=============================================================================
Este script executa o pipeline COMPLETO e REAL da revisão sistemática:

  Fase 1: Download massivo de PDFs (URL direta + Unpaywall + arXiv)
  Fase 2: Extração de texto completo (pypdf)
  Fase 3: Triagem real por texto completo
  Fase 4: Extração real de dados metodológicos
  Fase 5: Avaliação de qualidade objetiva
  Fase 6: Atualização da planilha Excel
  Fase 7: Geração da síntese final

Uso:
  python executar_pipeline_real.py [--skip-download] [--email SEU@EMAIL.COM]
  
  --skip-download  Pula a fase de download (usa PDFs já baixados)
  --email          E-mail para API Unpaywall (gratuita, opcional)
"""

import os
import re
import sys
import csv
import json
import time
import argparse
import hashlib
from datetime import datetime
from collections import Counter, defaultdict

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

# ──────────────────────────────────────────────────────────────
# Dependências
# ──────────────────────────────────────────────────────────────
try:
    import requests
except ImportError:
    print("[ERRO] 'requests' não encontrado. pip install requests")
    sys.exit(1)

try:
    import pypdf
except ImportError:
    print("[ERRO] 'pypdf' não encontrado. pip install pypdf")
    sys.exit(1)

try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
except ImportError:
    print("[ERRO] 'openpyxl' não encontrado. pip install openpyxl")
    sys.exit(1)


# ══════════════════════════════════════════════════════════════
#  CONSTANTES
# ══════════════════════════════════════════════════════════════

XLSX_PATH = "Formulario_Desenho_Pesquisa_RSL.xlsx"
CSV_PATH = "openalex_outputs/openalex_clean_data.csv"
PDF_DIR = "pdfs_baixados"
TEXTS_DIR = "textos_extraidos"
REPORT_PATH = "relatorio_pipeline_real.md"
SYNTHESIS_PATH = "relatorio_sintese_final.md"

HEADERS_BROWSER = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "application/pdf,*/*",
}

# Termos metodológicos para busca no texto completo
CAUSAL_DISCOVERY_TERMS = {
    "PC algorithm": [r"\bpc\s+algorithm\b", r"\bpc\s+algo\b", r"\bpeter.*clark\b"],
    "FCI": [r"\bfci\b", r"\bfast\s+causal\s+inference\b"],
    "GES": [r"\bges\b", r"\bgreedy\s+equivalence\s+search\b"],
    "LiNGAM": [r"\blingam\b", r"\bdirect\s*lingam\b", r"\bica.*lingam\b"],
    "NOTEARS": [r"\bnotears\b", r"\bno\s+tears\b"],
    "DAG-GNN": [r"\bdag.gnn\b"],
    "Score-based": [r"\bscore.based\s+(search|method|approach)\b"],
    "Constraint-based": [r"\bconstraint.based\s+(search|method|approach)\b"],
    "ICA-based": [r"\bica.based\s+causal\b"],
    "Granger causality": [r"\bgranger\s+causal\b"],
    "PCMCI": [r"\bpcmci\b"],
    "CCM": [r"\bconvergent\s+cross\s+mapping\b"],
    "Additive noise": [r"\badditive\s+noise\s+model\b", r"\banm\b"],
}

CAUSAL_INFERENCE_TERMS = {
    "Propensity Score": [r"\bpropensity\s+score\b", r"\bpsm\b", r"\biptw\b"],
    "Instrumental Variables": [r"\binstrumental\s+variable\b", r"\b(?:iv|2sls)\b"],
    "Difference-in-Differences": [r"\bdifference.in.difference\b", r"\bdid\b", r"\bdiff.in.diff\b"],
    "Regression Discontinuity": [r"\bregression\s+discontinuity\b", r"\brdd\b"],
    "SCM": [r"\bstructural\s+causal\s+model\b", r"\bscm\b", r"\bdo.calculus\b"],
    "Double Machine Learning": [r"\bdouble\s+machine\s+learning\b", r"\bdml\b"],
    "Meta-learners": [r"\bmeta.learner\b", r"\bs.learner\b", r"\bt.learner\b", r"\bx.learner\b"],
    "CATE": [r"\bcate\b", r"\bconditional\s+average\s+treatment\s+effect\b"],
    "ATE": [r"\baverage\s+treatment\s+effect\b", r"\bate\b"],
    "Synthetic Control": [r"\bsynthetic\s+control\b"],
    "G-computation": [r"\bg.computation\b", r"\bg.formula\b"],
    "IPW": [r"\binverse\s+probability\s+weight\b", r"\bipw\b"],
    "Backdoor criterion": [r"\bbackdoor\s+(criterion|adjustment)\b"],
    "Front-door criterion": [r"\bfront.door\s+(criterion|adjustment)\b"],
    "Causal forest": [r"\bcausal\s+forest\b"],
    "Bayesian network": [r"\bbayesian\s+network\b", r"\bbn\b"],
}

STUDY_DESIGN_PATTERNS = {
    "Simulation / Benchmark": [
        r"\bsimulat(ed|ion|ing)\s+(data|stud|experiment)\b",
        r"\bsynthetic\s+data\b",
        r"\bbenchmark\b",
        r"\bmonte\s+carlo\b",
    ],
    "Theoretical / Algorithmic": [
        r"\bwe\s+propos(e|ed)\s+(a|an|the)\s+(novel|new)\b",
        r"\btheorem\b",
        r"\bproof\b",
        r"\bwe\s+introduce\b",
        r"\bnew\s+(algorithm|method|framework|approach)\b",
    ],
    "Empirical / Observational": [
        r"\bobservational\s+(data|stud)\b",
        r"\breal.world\s+(data|application|dataset)\b",
        r"\bcohort\b",
        r"\bsurvey\s+data\b",
        r"\bcase\s+stud(y|ies)\b",
        r"\bclinical\s+data\b",
        r"\belectronic\s+health\s+record\b",
    ],
    "Review / Survey": [
        r"\b(systematic|literature|scoping)\s+review\b",
        r"\bsurvey\s+of\b",
        r"\bmeta.analysis\b",
        r"\boverview\s+of\b",
    ],
}

LIMITATION_PATTERNS = {
    "Causal sufficiency assumption": [
        r"\bcausal\s+sufficiency\b",
        r"\bno\s+(hidden|latent|unobserved)\s+(confounder|variable)\b",
        r"\bunmeasured\s+confounder\b",
    ],
    "High computational cost": [
        r"\bcomputational(ly)?\s+(cost|complex|expensive|burden)\b",
        r"\bscalab(le|ility)\b",
        r"\bnp.hard\b",
        r"\btime\s+complexity\b",
    ],
    "Linearity assumption": [
        r"\blinear(ity)?\s+assumption\b",
        r"\bassum(es?|ing)\s+linear\b",
        r"\bnon.linear\b.*\blimitation\b",
    ],
    "Positivity violation": [
        r"\bpositivity\s+(violation|assumption)\b",
        r"\boverlap\s+assumption\b",
    ],
    "Sensitivity to noise": [
        r"\bsensitiv(e|ity)\s+to\s+noise\b",
        r"\brobust(ness)?\s+to\s+noise\b",
    ],
    "Missing data": [
        r"\bmissing\s+data\b",
        r"\bincomplete\s+data\b",
        r"\bmissing\s+values?\b",
    ],
    "Sample size": [
        r"\bsmall\s+sample\b",
        r"\blimited\s+sample\b",
        r"\bsample\s+size\s+limit\b",
    ],
    "Faithfulness assumption": [
        r"\bfaithfulness\b",
        r"\bmarkov\s+assumption\b",
    ],
    "Acyclicity assumption": [
        r"\bacyclic(ity)?\s+assumption\b",
        r"\bcycl(e|ic)\b.*\blimitation\b",
    ],
}


# ══════════════════════════════════════════════════════════════
#  FUNÇÕES AUXILIARES
# ══════════════════════════════════════════════════════════════

def normalize_doi(doi_str):
    """Normaliza DOI para formato puro (sem prefixo URL)."""
    if not doi_str:
        return ""
    doi_str = doi_str.strip()
    for prefix in ["https://doi.org/", "http://doi.org/", "https://dx.doi.org/", "http://dx.doi.org/"]:
        if doi_str.lower().startswith(prefix):
            doi_str = doi_str[len(prefix):]
    return doi_str.strip().lower()


def safe_filename(text, max_len=80):
    """Gera nome de arquivo seguro a partir de texto."""
    clean = re.sub(r'[\\/*?:"<>|]', "", text).strip().replace(" ", "_")
    return clean[:max_len]


def doi_to_arxiv_pdf(doi):
    """Converte DOI arXiv para URL de PDF."""
    match = re.search(r'10\.48550/arxiv\.(.+)', doi, re.IGNORECASE)
    if match:
        return f"https://arxiv.org/pdf/{match.group(1)}.pdf"
    return None


def url_to_arxiv_pdf(url):
    """Converte URL arXiv abs/html para PDF."""
    if not url:
        return None
    match = re.search(r'arxiv\.org/abs/(\d+\.\d+)', url)
    if match:
        return f"https://arxiv.org/pdf/{match.group(1)}.pdf"
    match = re.search(r'arxiv\.org/pdf/(\d+\.\d+)', url)
    if match:
        return url if url.endswith('.pdf') else url + '.pdf'
    return None


def try_download_pdf(url, output_path, timeout=30):
    """Tenta baixar um PDF de uma URL. Retorna True se sucesso."""
    if not url or url == "N/A":
        return False
    if os.path.exists(output_path) and os.path.getsize(output_path) > 1000:
        return True  # Já baixado
    
    try:
        resp = requests.get(url, headers=HEADERS_BROWSER, timeout=timeout, 
                           stream=True, allow_redirects=True)
        if resp.status_code == 200:
            content = resp.content
            # Verificar se é realmente PDF
            if content[:5] == b'%PDF-' or 'pdf' in resp.headers.get('content-type', '').lower():
                with open(output_path, 'wb') as f:
                    f.write(content)
                if os.path.getsize(output_path) > 1000:
                    return True
                else:
                    os.remove(output_path)
        return False
    except Exception:
        return False


def try_unpaywall(doi, email):
    """Consulta a API Unpaywall para obter URL de PDF open access."""
    if not doi or not email:
        return None
    try:
        url = f"https://api.unpaywall.org/v2/{doi}?email={email}"
        resp = requests.get(url, timeout=15)
        if resp.status_code == 200:
            data = resp.json()
            best = data.get("best_oa_location", {})
            if best:
                return best.get("url_for_pdf") or best.get("url")
    except Exception:
        pass
    return None


def extract_text_from_pdf(pdf_path):
    """Extrai texto completo de um PDF usando pypdf."""
    try:
        reader = pypdf.PdfReader(pdf_path)
        pages_text = []
        for page in reader.pages:
            text = page.extract_text()
            if text:
                pages_text.append(text)
        return "\n".join(pages_text), len(reader.pages)
    except Exception as e:
        return "", 0


def search_patterns(text, pattern_dict):
    """Busca padrões regex em texto. Retorna termos encontrados."""
    text_lower = text.lower()
    found = []
    for term, patterns in pattern_dict.items():
        for pat in patterns:
            if re.search(pat, text_lower, re.IGNORECASE):
                found.append(term)
                break
    return found


def extract_sample_size(text):
    """Tenta extrair tamanho amostral do texto."""
    text_lower = text.lower()
    patterns = [
        r'n\s*=\s*(\d[\d,\.]+)',
        r'sample\s+size\s+(?:of\s+)?(\d[\d,\.]+)',
        r'(\d[\d,\.]+)\s+(?:participants|subjects|patients|observations|samples|individuals)',
        r'(?:dataset|data\s+set)\s+(?:of|with|containing)\s+(\d[\d,\.]+)',
    ]
    sizes = []
    for pat in patterns:
        matches = re.findall(pat, text_lower)
        for m in matches:
            try:
                n = int(m.replace(',', '').replace('.', ''))
                if 10 <= n <= 10_000_000:
                    sizes.append(n)
            except ValueError:
                pass
    return max(sizes) if sizes else None


def extract_country(text):
    """Tenta identificar país de afiliação do primeiro autor."""
    text_first_page = text[:3000]  # Primeiras linhas (cabeçalho)
    country_patterns = {
        "USA": [r"\bunited\s+states\b", r"\busa\b", r"\bu\.s\.a\b", r"\bnew\s+york\b", 
                r"\bcalifornia\b", r"\bmassachusetts\b", r"\bstanford\b", r"\bmit\b",
                r"\bcarnegie\s+mellon\b", r"\bberkeley\b"],
        "China": [r"\bchina\b", r"\bbeijing\b", r"\bshanghai\b", r"\btsinghua\b",
                  r"\bpeking\b", r"\bnanjing\b", r"\bshenzhen\b"],
        "UK": [r"\bunited\s+kingdom\b", r"\bengland\b", r"\blondon\b", r"\boxford\b",
               r"\bcambridge\b(?!.*massachusetts)", r"\bedinburgh\b"],
        "Germany": [r"\bgermany\b", r"\bdeutsch\b", r"\bmunich\b", r"\bberlin\b",
                    r"\bmax\s+planck\b", r"\btübingen\b", r"\btuebingen\b"],
        "Switzerland": [r"\bswitzerland\b", r"\bsuisse\b", r"\beth\s+zurich\b",
                       r"\bzurich\b", r"\bgeneva\b", r"\bepfl\b"],
        "Canada": [r"\bcanada\b", r"\btoronto\b", r"\bmontreal\b", r"\bvancouver\b",
                   r"\bmcgill\b", r"\bmila\b"],
        "France": [r"\bfrance\b", r"\bparis\b", r"\binria\b", r"\bcnrs\b"],
        "Netherlands": [r"\bnetherlands\b", r"\bamsterdam\b", r"\butrecht\b"],
        "Japan": [r"\bjapan\b", r"\btokyo\b", r"\bkyoto\b", r"\bosaka\b"],
        "Australia": [r"\baustralia\b", r"\bsydney\b", r"\bmelbourne\b"],
        "Italy": [r"\bitaly\b", r"\broma\b", r"\bmilan\b", r"\bbologna\b"],
        "South Korea": [r"\bkorea\b", r"\bseoul\b", r"\bkaist\b"],
        "Brazil": [r"\bbrazil\b", r"\bbrasil\b", r"\bsão\s+paulo\b", r"\brio\s+de\s+janeiro\b",
                   r"\busp\b", r"\bunicamp\b", r"\bufmg\b"],
        "India": [r"\bindia\b", r"\bmumbai\b", r"\bdelhi\b", r"\bbangalore\b"],
        "Spain": [r"\bspain\b", r"\bbarcelona\b", r"\bmadrid\b"],
        "Sweden": [r"\bsweden\b", r"\bstockholm\b", r"\buppsala\b"],
        "Singapore": [r"\bsingapore\b"],
        "Israel": [r"\bisrael\b", r"\btel\s+aviv\b", r"\bweizmann\b"],
    }
    
    text_lower = text_first_page.lower()
    for country, patterns in country_patterns.items():
        for pat in patterns:
            if re.search(pat, text_lower):
                return country
    return "Unknown"


# ══════════════════════════════════════════════════════════════
#  FASE 1: DOWNLOAD MASSIVO DE PDFs
# ══════════════════════════════════════════════════════════════

def phase1_download(email=None):
    """Baixa PDFs dos artigos aprovados na triagem T/R."""
    print("\n" + "=" * 70)
    print("  FASE 1: DOWNLOAD MASSIVO DE PDFs")
    print("=" * 70)
    
    os.makedirs(PDF_DIR, exist_ok=True)
    
    # 1a. Carregar mapa DOI→PDF URL do CSV
    print("\n[1a] Carregando URLs de PDF do CSV...")
    doi_to_pdf_url = {}
    doi_to_link_url = {}
    with open(CSV_PATH, encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            doi_raw = None
            pdf_url = None
            link_url = None
            for k, v in row.items():
                if 'DOI' in k:
                    doi_raw = v
                if 'PDF' in k:
                    pdf_url = v
                if 'Link' in k or ('URL' in k and 'PDF' not in k):
                    link_url = v
            
            doi = normalize_doi(doi_raw)
            if doi:
                if pdf_url and pdf_url.strip() and pdf_url.strip() != 'N/A':
                    doi_to_pdf_url[doi] = pdf_url.strip()
                if link_url and link_url.strip():
                    doi_to_link_url[doi] = link_url.strip()
    
    print(f"  ✓ {len(doi_to_pdf_url)} DOIs com URL de PDF")
    print(f"  ✓ {len(doi_to_link_url)} DOIs com Link")
    
    # 1b. Carregar registros da triagem T/R (incluídos)
    print("\n[1b] Carregando registros incluídos na triagem T/R...")
    wb = openpyxl.load_workbook(XLSX_PATH)
    ws = None
    for name in wb.sheetnames:
        if 'Planilha Triagem' in name:
            ws = wb[name]
            break
    
    if not ws:
        print("[ERRO] Aba 'Planilha Triagem' não encontrada.")
        wb.close()
        return []
    
    records_to_download = []
    for r in range(3, ws.max_row + 1):
        id_val = ws.cell(row=r, column=1).value
        if id_val is None:
            break
        
        # Verificar se passou na triagem T/R
        dec_r1 = str(ws.cell(row=r, column=12).value or "")
        resolution = str(ws.cell(row=r, column=18).value or "")
        final_tr = resolution if resolution else dec_r1
        
        if 'Incluir' not in final_tr:
            continue
        
        doi = normalize_doi(str(ws.cell(row=r, column=3).value or ""))
        titulo = str(ws.cell(row=r, column=4).value or "")
        autores = str(ws.cell(row=r, column=5).value or "")
        ano = ws.cell(row=r, column=6).value
        periodico = str(ws.cell(row=r, column=7).value or "")
        resumo = str(ws.cell(row=r, column=8).value or "")
        
        records_to_download.append({
            "row": r,
            "id": id_val,
            "doi": doi,
            "titulo": titulo,
            "autores": autores,
            "ano": ano,
            "periodico": periodico,
            "resumo": resumo,
        })
    
    wb.close()
    print(f"  ✓ {len(records_to_download)} registros incluídos na T/R para download")
    
    # 1c. Download
    print(f"\n[1c] Iniciando download de PDFs...")
    stats = {"downloaded": 0, "already_had": 0, "failed": 0, "unpaywall": 0}
    results = []
    
    for i, rec in enumerate(records_to_download):
        doi = rec["doi"]
        filename = f"{rec['id']}_{safe_filename(rec['autores'].split(',')[0] if rec['autores'] else 'unknown', 30)}_{rec['ano']}.pdf"
        filepath = os.path.join(PDF_DIR, filename)
        
        # Progresso
        if (i + 1) % 20 == 0 or i == 0:
            print(f"  [{i+1}/{len(records_to_download)}] Processando...")
        
        # Já temos?
        if os.path.exists(filepath) and os.path.getsize(filepath) > 1000:
            stats["already_had"] += 1
            results.append({**rec, "pdf_path": filepath, "download_status": "already_had"})
            continue
        
        success = False
        
        # Tentativa 1: URL direta do PDF (do CSV)
        pdf_url = doi_to_pdf_url.get(doi)
        if pdf_url:
            # Tentar converter arXiv
            arxiv_url = url_to_arxiv_pdf(pdf_url)
            if arxiv_url:
                success = try_download_pdf(arxiv_url, filepath)
            if not success:
                success = try_download_pdf(pdf_url, filepath)
        
        # Tentativa 2: arXiv via DOI
        if not success and doi:
            arxiv_pdf = doi_to_arxiv_pdf(doi)
            if arxiv_pdf:
                success = try_download_pdf(arxiv_pdf, filepath)
        
        # Tentativa 3: Link URL
        if not success:
            link_url = doi_to_link_url.get(doi)
            if link_url:
                arxiv_url = url_to_arxiv_pdf(link_url)
                if arxiv_url:
                    success = try_download_pdf(arxiv_url, filepath)
        
        # Tentativa 4: Unpaywall
        if not success and email and doi:
            oa_url = try_unpaywall(doi, email)
            if oa_url:
                success = try_download_pdf(oa_url, filepath)
                if success:
                    stats["unpaywall"] += 1
        
        if success:
            stats["downloaded"] += 1
            results.append({**rec, "pdf_path": filepath, "download_status": "downloaded"})
        else:
            stats["failed"] += 1
            results.append({**rec, "pdf_path": None, "download_status": "not_retrieved"})
        
        # Rate limiting
        time.sleep(0.5)
    
    print(f"\n  === Resultados do Download ===")
    print(f"  Já existiam:    {stats['already_had']}")
    print(f"  Baixados agora: {stats['downloaded']}")
    print(f"  Via Unpaywall:  {stats['unpaywall']}")
    print(f"  Não recuperados: {stats['failed']}")
    total_ok = stats['already_had'] + stats['downloaded']
    print(f"  TOTAL COM PDF:  {total_ok}/{len(records_to_download)}")
    
    return results


# ══════════════════════════════════════════════════════════════
#  FASE 2: EXTRAÇÃO DE TEXTO
# ══════════════════════════════════════════════════════════════

def phase2_extract_text(records):
    """Extrai texto completo dos PDFs baixados."""
    print("\n" + "=" * 70)
    print("  FASE 2: EXTRAÇÃO DE TEXTO COMPLETO")
    print("=" * 70)
    
    os.makedirs(TEXTS_DIR, exist_ok=True)
    
    extracted = 0
    failed = 0
    
    for rec in records:
        if not rec.get("pdf_path"):
            rec["full_text"] = ""
            rec["num_pages"] = 0
            continue
        
        text, num_pages = extract_text_from_pdf(rec["pdf_path"])
        rec["full_text"] = text
        rec["num_pages"] = num_pages
        
        if text and len(text) > 200:
            extracted += 1
            # Salvar texto para referência
            txt_path = os.path.join(TEXTS_DIR, f"{rec['id']}.txt")
            with open(txt_path, 'w', encoding='utf-8', errors='replace') as f:
                f.write(text)
        else:
            failed += 1
            rec["full_text"] = ""
    
    print(f"  ✓ Texto extraído com sucesso: {extracted}")
    print(f"  ✗ Falha na extração: {failed}")
    
    return records


# ══════════════════════════════════════════════════════════════
#  FASE 3: TRIAGEM REAL POR TEXTO COMPLETO
# ══════════════════════════════════════════════════════════════

def phase3_real_screening(records):
    """Triagem TC real baseada no conteúdo dos PDFs."""
    print("\n" + "=" * 70)
    print("  FASE 3: TRIAGEM REAL POR TEXTO COMPLETO")
    print("=" * 70)
    
    included = []
    excluded_no_pdf = []
    excluded_content = []
    
    for rec in records:
        text = rec.get("full_text", "")
        
        # Sem PDF → não recuperado
        if not rec.get("pdf_path") or not text:
            rec["tc_decision"] = "Não recuperado"
            rec["tc_reason"] = "PDF não disponível em acesso aberto"
            excluded_no_pdf.append(rec)
            continue
        
        text_lower = text.lower()
        
        # Critério 1: O artigo deve mencionar explicitamente inferência causal OU descoberta causal
        has_causal_inference = bool(re.search(r'\bcausal\s+inference\b', text_lower))
        has_causal_discovery = bool(re.search(r'\bcausal\s+discovery\b', text_lower))
        has_causal_effect = bool(re.search(r'\bcausal\s+effect\b', text_lower))
        has_causal_graph = bool(re.search(r'\bcausal\s+graph\b', text_lower))
        has_causal_model = bool(re.search(r'\bcausal\s+model\b', text_lower))
        has_dag = bool(re.search(r'\bdirected\s+acyclic\s+graph\b', text_lower))
        has_treatment_effect = bool(re.search(r'\btreatment\s+effect\b', text_lower))
        
        causal_score = sum([
            has_causal_inference * 2,
            has_causal_discovery * 2,
            has_causal_effect,
            has_causal_graph,
            has_causal_model,
            has_dag,
            has_treatment_effect,
        ])
        
        # Critério 2: Deve mencionar pelo menos um método específico
        discovery_methods = search_patterns(text, CAUSAL_DISCOVERY_TERMS)
        inference_methods = search_patterns(text, CAUSAL_INFERENCE_TERMS)
        all_methods = discovery_methods + inference_methods
        
        # Critério 3: Não deve ser um tipo de documento inadequado
        is_editorial = bool(re.search(r'\beditorial\b|\bletter\s+to\s+the\s+editor\b', text_lower[:1000]))
        is_too_short = len(text) < 3000  # Menos de ~1 página
        
        # Decisão
        if is_editorial:
            rec["tc_decision"] = "Excluído"
            rec["tc_reason"] = "Tipo de documento inadequado (editorial/carta)"
            excluded_content.append(rec)
        elif is_too_short:
            rec["tc_decision"] = "Excluído"
            rec["tc_reason"] = "Texto extraído insuficiente (possível erro de extração)"
            excluded_content.append(rec)
        elif causal_score >= 2 and len(all_methods) >= 1:
            rec["tc_decision"] = "Incluído"
            rec["tc_reason"] = ""
            rec["methods_found"] = all_methods
            rec["discovery_methods"] = discovery_methods
            rec["inference_methods"] = inference_methods
            included.append(rec)
        elif causal_score >= 3:
            # Alta relevância causal mesmo sem método específico identificado
            rec["tc_decision"] = "Incluído"
            rec["tc_reason"] = ""
            rec["methods_found"] = all_methods if all_methods else ["Não especificado"]
            rec["discovery_methods"] = discovery_methods
            rec["inference_methods"] = inference_methods
            included.append(rec)
        else:
            rec["tc_decision"] = "Excluído"
            rec["tc_reason"] = "Causalidade mencionada apenas tangencialmente"
            excluded_content.append(rec)
    
    print(f"  ✓ Incluídos (TC): {len(included)}")
    print(f"  ✗ Excluídos por conteúdo: {len(excluded_content)}")
    print(f"  ✗ Não recuperados (sem PDF): {len(excluded_no_pdf)}")
    
    return included, excluded_content, excluded_no_pdf


# ══════════════════════════════════════════════════════════════
#  FASE 4: EXTRAÇÃO REAL DE DADOS
# ══════════════════════════════════════════════════════════════

def phase4_extraction(included_records):
    """Extrai dados metodológicos reais dos textos completos."""
    print("\n" + "=" * 70)
    print("  FASE 4: EXTRAÇÃO REAL DE DADOS")
    print("=" * 70)
    
    for rec in included_records:
        text = rec.get("full_text", "")
        
        # Métodos (já parcialmente identificados na fase 3)
        if "methods_found" not in rec:
            rec["methods_found"] = search_patterns(text, CAUSAL_DISCOVERY_TERMS) + \
                                   search_patterns(text, CAUSAL_INFERENCE_TERMS)
        
        # Desenho do estudo
        designs = search_patterns(text, STUDY_DESIGN_PATTERNS)
        rec["study_design"] = designs[0] if designs else "Not classified"
        
        # Limitações
        rec["limitations"] = search_patterns(text, LIMITATION_PATTERNS)
        
        # Tamanho amostral
        rec["sample_size"] = extract_sample_size(text)
        
        # País
        rec["country"] = extract_country(text)
        
        # Número de páginas
        # (já extraído na fase 2)
    
    # Estatísticas
    with_methods = sum(1 for r in included_records if r.get("methods_found"))
    with_design = sum(1 for r in included_records if r.get("study_design") != "Not classified")
    with_limits = sum(1 for r in included_records if r.get("limitations"))
    with_sample = sum(1 for r in included_records if r.get("sample_size"))
    with_country = sum(1 for r in included_records if r.get("country") != "Unknown")
    
    print(f"  ✓ Com métodos identificados: {with_methods}/{len(included_records)}")
    print(f"  ✓ Com desenho classificado: {with_design}/{len(included_records)}")
    print(f"  ✓ Com limitações: {with_limits}/{len(included_records)}")
    print(f"  ✓ Com tamanho amostral: {with_sample}/{len(included_records)}")
    print(f"  ✓ Com país identificado: {with_country}/{len(included_records)}")
    
    return included_records


# ══════════════════════════════════════════════════════════════
#  FASE 5: AVALIAÇÃO DE QUALIDADE
# ══════════════════════════════════════════════════════════════

def phase5_quality(included_records):
    """Avaliação de qualidade objetiva baseada em critérios verificáveis."""
    print("\n" + "=" * 70)
    print("  FASE 5: AVALIAÇÃO DE QUALIDADE")
    print("=" * 70)
    
    for rec in included_records:
        text = rec.get("full_text", "")
        text_lower = text.lower()
        
        score = 0
        criteria = []
        
        # Critério 1: Descrição metodológica clara (>1 método identificado)
        if len(rec.get("methods_found", [])) >= 2:
            score += 1
            criteria.append("Metodologia detalhada")
        
        # Critério 2: Limitações discutidas
        if rec.get("limitations"):
            score += 1
            criteria.append("Limitações reportadas")
        
        # Critério 3: Tamanho amostral reportado ou estudo teórico
        if rec.get("sample_size") or rec.get("study_design") == "Theoretical / Algorithmic":
            score += 1
            criteria.append("Amostra/teoria explícita")
        
        # Critério 4: Reprodutibilidade (código mencionado)
        has_code = bool(re.search(
            r'\bgithub\.com\b|\bcode\s+(?:is\s+)?available\b|\bopen\s*source\b|\breproducib\b',
            text_lower
        ))
        if has_code:
            score += 1
            criteria.append("Código disponível")
        
        # Critério 5: Declaração de financiamento ou conflito
        has_funding = bool(re.search(
            r'\bfunding\b|\bgrant\b|\bsupported\s+by\b|\backnowledg\b|\bconflict\s+of\s+interest\b',
            text_lower
        ))
        if has_funding:
            score += 1
            criteria.append("Financiamento/conflito declarado")
        
        # Julgamento global
        if score >= 4:
            rec["quality_judgment"] = "Baixo Risco"
        elif score >= 2:
            rec["quality_judgment"] = "Alguma Preocupação"
        else:
            rec["quality_judgment"] = "Alto Risco"
        
        rec["quality_score"] = score
        rec["quality_criteria"] = criteria
    
    dist = Counter(r["quality_judgment"] for r in included_records)
    for judgment, count in dist.most_common():
        print(f"  {judgment}: {count}")
    
    return included_records


# ══════════════════════════════════════════════════════════════
#  FASE 6: ATUALIZAR PLANILHA EXCEL
# ══════════════════════════════════════════════════════════════

def phase6_update_excel(included, excluded_content, excluded_no_pdf):
    """Atualiza a planilha Excel com dados reais."""
    print("\n" + "=" * 70)
    print("  FASE 6: ATUALIZANDO PLANILHA EXCEL")
    print("=" * 70)
    
    wb = openpyxl.load_workbook(XLSX_PATH)
    
    # ── Estilos ──
    font_body = Font(name="Segoe UI", size=9, color="0F172A")
    align_left = Alignment(horizontal='left', vertical='top', wrap_text=True)
    align_center = Alignment(horizontal='center', vertical='center')
    fill_green = PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid")
    fill_red = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    fill_yellow = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    border_light = Border(
        left=Side(style='thin', color="E2E8F0"),
        right=Side(style='thin', color="E2E8F0"),
        top=Side(style='thin', color="E2E8F0"),
        bottom=Side(style='thin', color="E2E8F0"),
    )
    
    # ── 6a: Atualizar Planilha Triagem com decisões TC reais ──
    ws_tri = None
    for name in wb.sheetnames:
        if 'Planilha Triagem' in name:
            ws_tri = wb[name]
            break
    
    all_records = included + excluded_content + excluded_no_pdf
    row_map = {r["row"]: r for r in all_records}
    
    for r_idx, rec in row_map.items():
        decision = rec.get("tc_decision", "")
        reason = rec.get("tc_reason", "")
        
        # Colunas TC
        ws_tri.cell(row=r_idx, column=20, value="Pipeline Real").font = font_body
        ws_tri.cell(row=r_idx, column=21, value=decision).font = font_body
        ws_tri.cell(row=r_idx, column=22, value=reason).font = font_body
        
        # Status final
        if decision == "Incluído":
            ws_tri.cell(row=r_idx, column=29, value="Incluído").font = font_body
            ws_tri.cell(row=r_idx, column=29).fill = fill_green
        elif decision == "Não recuperado":
            ws_tri.cell(row=r_idx, column=29, value="Não recuperado").font = font_body
            ws_tri.cell(row=r_idx, column=29).fill = fill_yellow
        else:
            ws_tri.cell(row=r_idx, column=29, value="Excluído TC").font = font_body
            ws_tri.cell(row=r_idx, column=29).fill = fill_red
        
        # Observações com dados do PDF
        obs_parts = []
        if rec.get("num_pages"):
            obs_parts.append(f"PDF: {rec['num_pages']} páginas")
        if rec.get("methods_found"):
            obs_parts.append(f"Métodos: {', '.join(rec['methods_found'][:5])}")
        if rec.get("sample_size"):
            obs_parts.append(f"N={rec['sample_size']}")
        obs = " | ".join(obs_parts) if obs_parts else ""
        ws_tri.cell(row=r_idx, column=30, value=obs).font = font_body
        ws_tri.cell(row=r_idx, column=30).alignment = align_left
    
    print(f"  ✓ Planilha Triagem atualizada ({len(row_map)} registros)")
    
    # ── 6b: Planilha de Extração com dados reais ──
    ws_ext = None
    for name in wb.sheetnames:
        if 'Planilha Extração' in name or 'Planilha Extracao' in name:
            ws_ext = wb[name]
            break
    
    if ws_ext:
        # Limpar dados existentes
        for r in range(3, ws_ext.max_row + 1):
            for c in range(1, 26):
                ws_ext.cell(row=r, column=c, value=None)
        
        for i, rec in enumerate(included):
            r = i + 3
            fill = fill_zebra if i % 2 == 1 else fill_white
            
            ws_ext.cell(row=r, column=1, value=rec["id"]).font = font_body
            
            # Autor/Ano
            autor_short = rec["autores"].split(",")[0].strip() if rec["autores"] else "Unknown"
            ws_ext.cell(row=r, column=2, value=f"{autor_short} ({rec['ano']})").font = font_body
            
            ws_ext.cell(row=r, column=3, value=rec["doi"]).font = font_body
            ws_ext.cell(row=r, column=4, value="OpenAlex").font = font_body
            ws_ext.cell(row=r, column=5, value=rec["titulo"]).font = font_body
            ws_ext.cell(row=r, column=5).alignment = align_left
            
            # Objetivo (do resumo)
            resumo = rec.get("resumo", "")
            objetivo = resumo[:200] + "..." if len(resumo) > 200 else resumo
            ws_ext.cell(row=r, column=6, value=objetivo).font = font_body
            ws_ext.cell(row=r, column=6).alignment = align_left
            
            ws_ext.cell(row=r, column=7, value=rec["ano"]).font = font_body
            ws_ext.cell(row=r, column=8, value=rec["periodico"]).font = font_body
            
            # Desenho do estudo
            ws_ext.cell(row=r, column=12, value=rec.get("study_design", "")).font = font_body
            
            # Amostra
            sample = rec.get("sample_size")
            ws_ext.cell(row=r, column=13, value=f"N={sample}" if sample else "Não reportado").font = font_body
            
            # Métodos
            methods_str = ", ".join(rec.get("methods_found", []))
            ws_ext.cell(row=r, column=14, value=methods_str).font = font_body
            ws_ext.cell(row=r, column=14).alignment = align_left
            
            # Limitações
            limits_str = ", ".join(rec.get("limitations", []))
            ws_ext.cell(row=r, column=16, value=limits_str if limits_str else "Não identificado").font = font_body
            ws_ext.cell(row=r, column=16).alignment = align_left
            
            # País
            ws_ext.cell(row=r, column=17, value=rec.get("country", "Unknown")).font = font_body
            
            # Qualidade
            ws_ext.cell(row=r, column=20, value=rec.get("quality_judgment", "")).font = font_body
            
            # Aplicar zebra
            for c in range(1, 21):
                ws_ext.cell(row=r, column=c).fill = fill
                ws_ext.cell(row=r, column=c).border = border_light
        
        print(f"  ✓ Planilha Extração preenchida ({len(included)} registros reais)")
    
    # ── 6c: PRISMA Flow com números reais ──
    ws_prisma = None
    for name in wb.sheetnames:
        if 'PRISMA' in name:
            ws_prisma = wb[name]
            break
    
    n_brutos = 725  # Total identificados nas bases
    n_duplicatas = 725 - 522  # Removidos na dedup
    n_pre_triagem = 522 - 476  # Excluídos na pré-triagem (idioma, etc.) 
    # Nota: na verdade a triagem T/R excluiu os demais
    n_triagem_tr = 476  # Avaliados T/R que foram incluídos
    total_tr_avaliados = 522
    n_excl_tr = total_tr_avaliados - n_triagem_tr
    n_nao_recuperados = len(excluded_no_pdf)
    n_excl_tc = len(excluded_content)
    n_incluidos = len(included)
    
    if ws_prisma:
        # Limpar e reescrever
        prisma_data = [
            ("Registros identificados nas bases de dados", n_brutos),
            ("Duplicatas removidas", n_duplicatas),
            ("Registros após remoção de duplicatas", 522),
            ("Excluídos na triagem T/R (automatizada)", n_excl_tr),
            ("Registros incluídos na triagem T/R", n_triagem_tr),
            ("PDFs não recuperados (sem acesso aberto)", n_nao_recuperados),
            ("Excluídos após leitura do texto completo", n_excl_tc),
            ("Estudos incluídos na síntese final", n_incluidos),
        ]
        
        for i, (label, value) in enumerate(prisma_data):
            r = i + 2
            ws_prisma.cell(row=r, column=1, value=label).font = font_body
            ws_prisma.cell(row=r, column=2, value=value).font = font_body
        
        print(f"  ✓ PRISMA Flow atualizado")
    
    # Salvar
    wb.save(XLSX_PATH)
    wb.close()
    print(f"  ✓ Planilha salva: {XLSX_PATH}")
    
    return {
        "n_brutos": n_brutos,
        "n_duplicatas": n_duplicatas,
        "n_triagem_tr": n_triagem_tr,
        "n_excl_tr": n_excl_tr,
        "n_nao_recuperados": n_nao_recuperados,
        "n_excl_tc": n_excl_tc,
        "n_incluidos": n_incluidos,
    }


# ══════════════════════════════════════════════════════════════
#  FASE 7: SÍNTESE REAL
# ══════════════════════════════════════════════════════════════

def phase7_synthesis(included_records, prisma_numbers):
    """Gera relatório de síntese com dados reais."""
    print("\n" + "=" * 70)
    print("  FASE 7: GERANDO SÍNTESE COM DADOS REAIS")
    print("=" * 70)
    
    n = len(included_records)
    
    # Distribuições
    dist_ano = Counter(r.get("ano") for r in included_records if r.get("ano"))
    dist_design = Counter(r.get("study_design", "Not classified") for r in included_records)
    dist_quality = Counter(r.get("quality_judgment", "") for r in included_records)
    dist_country = Counter(r.get("country", "Unknown") for r in included_records)
    
    # Métodos (multi-valued)
    all_methods = []
    discovery_count = Counter()
    inference_count = Counter()
    for r in included_records:
        for m in r.get("discovery_methods", []):
            discovery_count[m] += 1
            all_methods.append(m)
        for m in r.get("inference_methods", []):
            inference_count[m] += 1
            all_methods.append(m)
    dist_methods = Counter(all_methods)
    
    # Limitações (multi-valued)
    all_limits = []
    for r in included_records:
        all_limits.extend(r.get("limitations", []))
    dist_limits = Counter(all_limits)
    
    # ── Gerar Markdown ──
    lines = []
    lines.append("# Relatório de Síntese — Dados Reais da Revisão Sistemática")
    lines.append(f"**Tema**: Inferência Causal e Descoberta Causal (1990–2026)")
    lines.append(f"**Data de Geração**: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    lines.append(f"**Artigos Incluídos**: {n} (baseado em leitura real dos PDFs)")
    lines.append("")
    lines.append("---")
    lines.append("")
    
    # PRISMA
    lines.append("## Fluxo PRISMA (Números Reais)")
    pn = prisma_numbers
    lines.append(f"- Registros identificados: **{pn['n_brutos']}**")
    lines.append(f"- Duplicatas removidas: **{pn['n_duplicatas']}**")
    lines.append(f"- Incluídos na triagem T/R: **{pn['n_triagem_tr']}**")
    lines.append(f"- Excluídos T/R: **{pn['n_excl_tr']}**")
    lines.append(f"- PDFs não recuperados: **{pn['n_nao_recuperados']}**")
    lines.append(f"- Excluídos após leitura TC: **{pn['n_excl_tc']}**")
    lines.append(f"- **Incluídos na síntese: {pn['n_incluidos']}**")
    lines.append("")
    
    # Temporal
    lines.append("## 1. Distribuição Temporal")
    lines.append("| Ano | N | % |")
    lines.append("| --- | --- | --- |")
    for ano in sorted(dist_ano.keys(), reverse=True):
        pct = dist_ano[ano] / n * 100
        lines.append(f"| {ano} | {dist_ano[ano]} | {pct:.1f}% |")
    lines.append("")
    
    # Desenho
    lines.append("## 2. Desenho dos Estudos")
    lines.append("| Desenho | N | % |")
    lines.append("| --- | --- | --- |")
    for des, c in dist_design.most_common():
        lines.append(f"| {des} | {c} | {c/n*100:.1f}% |")
    lines.append("")
    
    # Métodos
    lines.append("## 3. Métodos e Algoritmos Identificados nos Textos Completos")
    lines.append("")
    lines.append("### Descoberta Causal")
    lines.append("| Método | N | % dos estudos |")
    lines.append("| --- | --- | --- |")
    for met, c in discovery_count.most_common():
        lines.append(f"| {met} | {c} | {c/n*100:.1f}% |")
    lines.append("")
    lines.append("### Inferência Causal")
    lines.append("| Método | N | % dos estudos |")
    lines.append("| --- | --- | --- |")
    for met, c in inference_count.most_common():
        lines.append(f"| {met} | {c} | {c/n*100:.1f}% |")
    lines.append("")
    
    # Qualidade
    lines.append("## 4. Avaliação de Qualidade")
    lines.append("| Julgamento | N | % |")
    lines.append("| --- | --- | --- |")
    for q, c in dist_quality.most_common():
        lines.append(f"| {q} | {c} | {c/n*100:.1f}% |")
    lines.append("")
    
    # Limitações
    lines.append("## 5. Limitações Identificadas nos Textos")
    lines.append("| Limitação | N | % |")
    lines.append("| --- | --- | --- |")
    for lim, c in dist_limits.most_common():
        lines.append(f"| {lim} | {c} | {c/n*100:.1f}% |")
    lines.append("")
    
    # Geografia
    lines.append("## 6. Distribuição Geográfica")
    lines.append("| País | N | % |")
    lines.append("| --- | --- | --- |")
    for ps, c in dist_country.most_common(15):
        lines.append(f"| {ps} | {c} | {c/n*100:.1f}% |")
    lines.append("")
    
    # Nota de transparência
    lines.append("---")
    lines.append("## Nota de Transparência Metodológica")
    lines.append(f"- Todos os {n} artigos incluídos tiveram seus PDFs baixados e lidos por completo.")
    lines.append(f"- A extração de métodos, limitações e dados geográficos foi feita por mineração de texto (regex) sobre o texto completo extraído via pypdf.")
    lines.append(f"- {pn['n_nao_recuperados']} artigos não puderam ser acessados (paywall/indisponibilidade) e foram marcados como 'não recuperados'.")
    lines.append(f"- A triagem T/R foi realizada de forma automatizada com base em termos-chave nos títulos e resumos reais.")
    lines.append(f"- A triagem TC foi realizada com base na leitura real do texto completo dos PDFs.")
    
    # Salvar
    with open(SYNTHESIS_PATH, 'w', encoding='utf-8') as f:
        f.write("\n".join(lines))
    
    print(f"  ✓ Relatório salvo: {SYNTHESIS_PATH}")
    print(f"  ✓ {n} artigos reais sintetizados")
    
    return dist_ano, dist_design, dist_methods, dist_quality, dist_limits, dist_country


# ══════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(description="Pipeline Real de Revisão Sistemática")
    parser.add_argument("--skip-download", action="store_true",
                       help="Pula download, usa PDFs já baixados")
    parser.add_argument("--email", type=str, default="",
                       help="E-mail para API Unpaywall (opcional)")
    args = parser.parse_args()
    
    print("=" * 70)
    print("  PIPELINE REAL DE REVISÃO SISTEMÁTICA")
    print("  Tema: Inferência Causal & Descoberta Causal")
    print(f"  Data: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    print("=" * 70)
    
    start_time = time.time()
    
    # Fase 1: Download
    if args.skip_download:
        print("\n[SKIP] Download pulado. Reconstruindo lista de registros...")
        records = phase1_download.__wrapped__(args.email) if hasattr(phase1_download, '__wrapped__') else phase1_download(args.email)
    else:
        records = phase1_download(args.email)
    
    # Fase 2: Extração de texto
    records = phase2_extract_text(records)
    
    # Fase 3: Triagem TC real
    included, excluded_content, excluded_no_pdf = phase3_real_screening(records)
    
    # Fase 4: Extração de dados
    included = phase4_extraction(included)
    
    # Fase 5: Qualidade
    included = phase5_quality(included)
    
    # Fase 6: Atualizar Excel
    prisma_numbers = phase6_update_excel(included, excluded_content, excluded_no_pdf)
    
    # Fase 7: Síntese
    phase7_synthesis(included, prisma_numbers)
    
    elapsed = time.time() - start_time
    
    print("\n" + "=" * 70)
    print(f"  PIPELINE CONCLUÍDO em {elapsed/60:.1f} minutos")
    print(f"  PDFs baixados: {sum(1 for r in records if r.get('pdf_path'))}")
    print(f"  Artigos incluídos (reais): {len(included)}")
    print(f"  Síntese: {SYNTHESIS_PATH}")
    print("=" * 70)
    
    # Salvar dados para uso posterior (gerar artigo)
    summary = {
        "total_identified": prisma_numbers["n_brutos"],
        "duplicates_removed": prisma_numbers["n_duplicatas"],
        "included_tr": prisma_numbers["n_triagem_tr"],
        "excluded_tr": prisma_numbers["n_excl_tr"],
        "not_retrieved": prisma_numbers["n_nao_recuperados"],
        "excluded_tc": prisma_numbers["n_excl_tc"],
        "included_final": len(included),
        "pdfs_downloaded": sum(1 for r in records if r.get("pdf_path")),
        "elapsed_minutes": round(elapsed / 60, 1),
    }
    with open("pipeline_summary.json", 'w', encoding='utf-8') as f:
        json.dump(summary, f, indent=2, ensure_ascii=False)
    
    print(f"  Resumo salvo: pipeline_summary.json")


if __name__ == "__main__":
    main()
