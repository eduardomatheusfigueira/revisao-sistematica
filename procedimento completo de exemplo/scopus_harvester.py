import os
import re
import csv
import json
import time
import logging
import argparse
from dataclasses import dataclass, asdict
from typing import List, Dict, Any, Optional
from collections import Counter
import requests
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

# Configure logging
def setup_logging(log_path: str):
    logger = logging.getLogger("ScopusHarvester")
    logger.setLevel(logging.DEBUG)
    
    if not logger.handlers:
        # File handler
        file_handler = logging.FileHandler(log_path, encoding="utf-8")
        file_handler.setLevel(logging.DEBUG)
        file_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(name)s - %(message)s')
        file_handler.setFormatter(file_formatter)
        logger.addHandler(file_handler)
        
        # Console handler
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        console_formatter = logging.Formatter('%(levelname)s: %(message)s')
        console_handler.setFormatter(console_formatter)
        logger.addHandler(console_handler)
        
    return logger

@dataclass
class ScopusRecord:
    id: str
    doi: str
    title: str
    authors: str
    year: int
    publication_date: str
    journal_name: str
    journal_issns: str
    language: str
    url: str
    pdf_url: str
    citations: int
    concepts: str
    abstract: str

    @classmethod
    def from_scopus_raw(cls, entry: Dict[str, Any]) -> "ScopusRecord":
        # Extract unique identifier
        raw_id = entry.get("dc:identifier", "")
        clean_id = raw_id.split(":")[-1] if raw_id else entry.get("eid", "N/A")
        
        # Extract DOI
        doi = entry.get("prism:doi") or "N/A"
        
        # Extract Title
        title = entry.get("dc:title") or "N/A"
        
        # Extract Authors (falls back to creator, will be enriched later if possible)
        authors = entry.get("dc:creator") or "N/A"
        
        # Extract Year & Date
        cover_date = entry.get("prism:coverDate") or "N/A"
        year = 0
        if cover_date != "N/A":
            try:
                year = int(cover_date.split("-")[0])
            except ValueError:
                year = 0
        
        # Extract Journal Information
        journal_name = entry.get("prism:publicationName") or "N/A"
        
        issn = entry.get("prism:issn")
        eissn = entry.get("prism:eIssn")
        issns_list = []
        if issn: issns_list.append(issn)
        if eissn: issns_list.append(eissn)
        journal_issns = "; ".join(issns_list) if issns_list else "N/A"
        
        # URL link (Scopus record page)
        url_link = "N/A"
        links = entry.get("link", [])
        for link in links:
            if link.get("@ref") == "scopus":
                url_link = link.get("@href")
                break
        if url_link == "N/A":
            url_link = entry.get("prism:url", "N/A")
            
        # Citations
        try:
            citations = int(entry.get("citedby-count") or 0)
        except ValueError:
            citations = 0
            
        return cls(
            id=clean_id,
            doi=doi,
            title=title,
            authors=authors,
            year=year,
            publication_date=cover_date,
            journal_name=journal_name,
            journal_issns=journal_issns,
            language="N/A",
            url=url_link,
            pdf_url="N/A",
            citations=citations,
            concepts="N/A",
            abstract="N/A"
        )

class ScopusClient:
    def __init__(self, api_key: str, base_url: str, user_agent: str, politeness_delay: float, max_retries: int, backoff_factor: float):
        self.api_key = api_key
        self.base_url = base_url
        self.user_agent = user_agent
        self.politeness_delay = politeness_delay
        self.max_retries = max_retries
        self.backoff_factor = backoff_factor
        self.session = requests.Session()
        
    def search_page(self, query: str, start: int, count: int) -> Optional[Dict[str, Any]]:
        headers = {
            "X-ELS-APIKey": self.api_key,
            "Accept": "application/json",
            "User-Agent": self.user_agent
        }
        
        params = {
            "query": query,
            "start": start,
            "count": count
        }
        
        retries = 0
        while retries < self.max_retries:
            try:
                response = self.session.get(self.base_url, params=params, headers=headers, timeout=20)
                if response.status_code == 200:
                    return response.json()
                elif response.status_code in [429, 500, 502, 503, 504]:
                    sleep_time = self.backoff_factor ** retries
                    print(f"Warning: Scopus API status {response.status_code}. Retrying in {sleep_time:.2f}s...")
                    time.sleep(sleep_time)
                    retries += 1
                else:
                    print(f"Error: Scopus API returned HTTP status {response.status_code}: {response.text}")
                    return None
            except requests.RequestException as e:
                sleep_time = self.backoff_factor ** retries
                print(f"Connection warning: {e}. Retrying in {sleep_time:.2f}s...")
                time.sleep(sleep_time)
                retries += 1
                
        return None

class MetadataEnricher:
    def __init__(self, api_key: str, logger: logging.Logger):
        self.api_key = api_key
        self.logger = logger
        self.session = requests.Session()
        
    def enrich_record(self, record: ScopusRecord) -> ScopusRecord:
        # Step 1: Try native Scopus Abstract Retrieval API (if institutional network allows)
        self.logger.debug(f"Attempting native Scopus abstract retrieval for EID {record.id}...")
        url = f"https://api.elsevier.com/content/abstract/scopus_id/{record.id}"
        headers = {
            "X-ELS-APIKey": self.api_key,
            "Accept": "application/json"
        }
        
        try:
            res = self.session.get(url, headers=headers, timeout=10)
            if res.status_code == 200:
                data = res.json()
                doc = data.get("abstracts-retrieval-response", {})
                coredata = doc.get("coredata", {})
                
                # Check if abstract is in coredata
                abstract_text = coredata.get("dc:description")
                if abstract_text:
                    record.abstract = abstract_text
                    self.logger.debug(f"  Successfully retrieved abstract natively from Scopus for EID {record.id}.")
                    
                    # Extract full authors list from native response
                    authors_obj = coredata.get("dc:creator", {})
                    author_list = authors_obj.get("author", []) if isinstance(authors_obj, dict) else []
                    if author_list:
                        names = []
                        for auth in author_list:
                            pref = auth.get("preferred-name", {})
                            surname = pref.get("ce:surname")
                            given = pref.get("ce:given-name")
                            if surname and given:
                                names.append(f"{surname}, {given}")
                            elif auth.get("ce:indexed-name"):
                                names.append(auth.get("ce:indexed-name"))
                        if names:
                            record.authors = "; ".join(names)
                    return record
            else:
                self.logger.debug(f"  Native abstract retrieval returned HTTP status {res.status_code}. Falling back to OpenAlex.")
        except Exception as e:
            self.logger.debug(f"  Native abstract retrieval failed: {e}. Falling back to OpenAlex.")
            
        # Step 2: Fallback to OpenAlex lookup by DOI
        if record.doi and record.doi != "N/A":
            clean_doi = record.doi.replace("https://doi.org/", "").strip()
            openalex_url = f"https://api.openalex.org/works/https://doi.org/{clean_doi}"
            headers_oa = {"User-Agent": "ScopusHarvester/1.0 (contact: user@mail.com)"}
            try:
                res_oa = self.session.get(openalex_url, headers=headers_oa, timeout=10)
                if res_oa.status_code == 200:
                    oa_data = res_oa.json()
                    self.logger.debug(f"  Successfully matched DOI {record.doi} in OpenAlex.")
                    self._fill_from_openalex(record, oa_data)
                    return record
            except Exception as e:
                self.logger.debug(f"  OpenAlex DOI lookup failed: {e}")
                
        # Step 3: Fallback to OpenAlex lookup by Title
        if record.title and record.title != "N/A":
            # Clean title
            clean_title = record.title.split("As condicionantes")[0].strip()
            openalex_search_url = "https://api.openalex.org/works"
            headers_oa = {"User-Agent": "ScopusHarvester/1.0 (contact: user@mail.com)"}
            params = {"filter": f"title.search:{clean_title}"}
            try:
                res_oa = self.session.get(openalex_search_url, params=params, headers=headers_oa, timeout=10)
                if res_oa.status_code == 200:
                    results = res_oa.json().get("results", [])
                    if results:
                        # Match first result
                        self.logger.debug(f"  Successfully matched Title '{record.title[:30]}' in OpenAlex.")
                        self._fill_from_openalex(record, results[0])
                        return record
            except Exception as e:
                self.logger.debug(f"  OpenAlex Title lookup failed: {e}")
                
        return record

    def _fill_from_openalex(self, record: ScopusRecord, oa_data: Dict[str, Any]):
        # Reconstruct abstract
        abstract_inverted = oa_data.get("abstract_inverted_index") or {}
        if abstract_inverted:
            try:
                word_positions = []
                for word, positions in abstract_inverted.items():
                    for pos in positions:
                        word_positions.append((pos, word))
                word_positions.sort()
                record.abstract = " ".join([word for _, word in word_positions])
            except Exception:
                pass
                
        # Extract full authors
        authorships = oa_data.get("authorships") or []
        authors_list = []
        for auth in authorships:
            author_obj = auth.get("author") or {}
            name = author_obj.get("display_name")
            if name:
                authors_list.append(name)
        if authors_list:
            record.authors = "; ".join(authors_list)
            
        # Extract concepts
        concepts_list = oa_data.get("concepts") or []
        concept_names = []
        for c in concepts_list:
            c_name = c.get("display_name")
            if c_name:
                concept_names.append(c_name)
        if concept_names:
            record.concepts = "; ".join(concept_names)
            
        # Extract language
        record.language = oa_data.get("language") or "N/A"
        
        # Extract open access PDF link if available
        open_access = oa_data.get("open_access") or {}
        pdf_url = open_access.get("oa_url")
        if pdf_url:
            record.pdf_url = pdf_url

class DataCleaner:
    @staticmethod
    def clean_text(text: Optional[str]) -> str:
        if not text or text == "N/A":
            return "N/A"
        text = re.sub(r'<[^>]+>', ' ', text)
        text = re.sub(r'\s+', ' ', text)
        return text.strip()

    @classmethod
    def clean_record(cls, record: ScopusRecord) -> ScopusRecord:
        record.title = cls.clean_text(record.title)
        record.authors = cls.clean_text(record.authors)
        record.journal_name = cls.clean_text(record.journal_name)
        record.concepts = cls.clean_text(record.concepts)
        record.abstract = cls.clean_text(record.abstract)
        
        if not record.doi: record.doi = "N/A"
        if not record.language: record.language = "N/A"
        if not record.url: record.url = "N/A"
        if not record.pdf_url: record.pdf_url = "N/A"
        
        return record

class ScopusExporter:
    def __init__(self, output_dir: str, logger: logging.Logger):
        self.output_dir = output_dir
        self.logger = logger
        
    def export_to_json(self, records: List[ScopusRecord], filename: str):
        json_path = os.path.join(self.output_dir, filename)
        self.logger.info(f"Exporting JSON raw backup: {json_path}")
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump([asdict(r) for r in records], f, ensure_ascii=False, indent=2)
            
    def export_to_csv(self, records: List[ScopusRecord], filename: str):
        csv_path = os.path.join(self.output_dir, filename)
        self.logger.info(f"Exporting CSV: {csv_path}")
        
        headers = [
            "ID Scopus (EID)", "DOI", "Título", "Autores", "Ano de Publicação", 
            "Data de Publicação", "Periódico (Nome)", "Periódico (ISSN)", 
            "Idioma", "Link / URL", "PDF URL", "Citações", "Palavras-chave / Conceitos", "Resumo"
        ]
        
        with open(csv_path, "w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(headers)
            for r in records:
                writer.writerow([
                    r.id, r.doi, r.title, r.authors, r.year,
                    r.publication_date, r.journal_name, r.journal_issns,
                    r.language, r.url, r.pdf_url, r.citations, r.concepts, r.abstract
                ])
                
    def export_to_excel(self, records: List[ScopusRecord], filename: str):
        excel_path = os.path.join(self.output_dir, filename)
        self.logger.info(f"Creating a brand-new Excel workbook: {excel_path}")
        
        wb = openpyxl.Workbook()
        
        font_family = "Segoe UI"
        font_header = Font(name=font_family, size=10, bold=True, color="FFFFFF")
        font_regular = Font(name=font_family, size=10)
        
        # Scopus Steel Blue branding palette
        fill_header = PatternFill(start_color="3C5A80", end_color="3C5A80", fill_type="solid") # Steel Blue
        fill_zebra = PatternFill(start_color="F1F4F8", end_color="F1F4F8", fill_type="solid")  # Soft Steel White
        
        border_all = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )
        
        align_left = Alignment(horizontal='left', vertical='center', wrap_text=False)
        align_center = Alignment(horizontal='center', vertical='center')
        align_wrap_left = Alignment(horizontal='left', vertical='top', wrap_text=True)
        
        # 1. Data Sheet
        ws_data = wb.active
        ws_data.title = "Scopus Cleaned Data"
        
        headers = [
            "ID Scopus (EID)", "DOI", "Título", "Autores", "Ano de Publicação", 
            "Data de Publicação", "Periódico (Nome)", "Periódico (ISSN)", 
            "Idioma", "Link / URL", "PDF URL", "Citações", "Palavras-chave / Conceitos", "Resumo"
        ]
        
        ws_data.row_dimensions[1].height = 28
        for col_idx, h in enumerate(headers, 1):
            cell = ws_data.cell(row=1, column=col_idx, value=h)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = align_center
            cell.border = border_all
            
        for row_idx, r in enumerate(records, 2):
            ws_data.row_dimensions[row_idx].height = 22
            row_data = [
                r.id, r.doi, r.title, r.authors, r.year,
                r.publication_date, r.journal_name, r.journal_issns,
                r.language, r.url, r.pdf_url, r.citations, r.concepts, r.abstract
            ]
            
            is_zebra = (row_idx % 2 == 0)
            
            for col_idx, val in enumerate(row_data, 1):
                cell = ws_data.cell(row=row_idx, column=col_idx, value=val)
                cell.font = font_regular
                cell.border = border_all
                if is_zebra:
                    cell.fill = fill_zebra
                    
                if col_idx in [1, 5, 12]:  # ID, Year, Citations
                    cell.alignment = align_center
                    if col_idx in [5, 12] and isinstance(val, int):
                        cell.number_format = '#,##0'
                elif col_idx in [2, 6, 8, 9, 10, 11]:  # DOI, Date, ISSN, Lang, Links
                    cell.alignment = align_center
                elif col_idx in [3, 4, 7, 13]:  # Title, Authors, Journal, Concepts
                    cell.alignment = align_left
                elif col_idx == 14:  # Abstract (wrapped)
                    cell.alignment = align_wrap_left
                    ws_data.row_dimensions[row_idx].height = 50
                    
        for col in ws_data.columns:
            max_len = 0
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            for cell in col:
                val = str(cell.value or '')
                if cell.row == 1:
                    max_len = max(max_len, len(val) + 4)
                else:
                    if col[0].column == 14:  # Abstract
                        max_len = max(max_len, 45)
                    else:
                        max_len = max(max_len, min(len(val), 35))
            ws_data.column_dimensions[col_letter].width = max(max_len, 10)
            
        # 2. Summary Sheet
        ws_summary = wb.create_sheet("Extraction Summary")
        ws_summary.row_dimensions[2].height = 26
        ws_summary.cell(row=2, column=2, value="Scopus Extraction Summary").font = Font(name=font_family, size=16, bold=True, color="3C5A80")
        
        ws_summary.cell(row=4, column=2, value="Execution Date:").font = Font(name=font_family, size=10, bold=True)
        ws_summary.cell(row=4, column=3, value=time.strftime('%Y-%m-%d %H:%M:%S')).font = font_regular
        ws_summary.cell(row=5, column=2, value="Total Unique Works:").font = Font(name=font_family, size=10, bold=True)
        ws_summary.cell(row=5, column=3, value=len(records)).font = font_regular
        
        journals_count = Counter([r.journal_name for r in records if r.journal_name != "N/A"])
        
        ws_summary.cell(row=7, column=2, value="Top Journals Summary").font = Font(name=font_family, size=12, bold=True, color="3C5A80")
        ws_summary.cell(row=8, column=2, value="Journal Name").font = font_header
        ws_summary.cell(row=8, column=2).fill = fill_header
        ws_summary.cell(row=8, column=3, value="Count").font = font_header
        ws_summary.cell(row=8, column=3).fill = fill_header
        
        cur_row = 9
        for j, count in journals_count.most_common(15):
            ws_summary.row_dimensions[cur_row].height = 20
            c_name = ws_summary.cell(row=cur_row, column=2, value=j)
            c_count = ws_summary.cell(row=cur_row, column=3, value=count)
            c_name.font = font_regular
            c_count.font = font_regular
            c_name.border = border_all
            c_count.border = border_all
            c_name.alignment = align_left
            c_count.alignment = align_center
            cur_row += 1
            
        wb.save(excel_path)
        self.logger.info(f"Excel workbook created and saved successfully at: {excel_path}")

    def generate_analysis_report(self, records: List[ScopusRecord], filename: str, query: str):
        report_path = os.path.join(self.output_dir, filename)
        self.logger.info(f"Generating markdown summary report: {report_path}")
        
        total = len(records)
        years = [r.year for r in records if r.year > 0]
        journals = [r.journal_name for r in records if r.journal_name != "N/A"]
        
        languages_flat = []
        for r in records:
            if r.language and r.language != "N/A":
                languages_flat.extend([lang.strip() for lang in r.language.split(";") if lang.strip()])
                
        years_count = Counter(years)
        journal_count = Counter(journals)
        lang_count = Counter(languages_flat)
        
        authors_list = []
        for r in records:
            if r.authors and r.authors != "N/A":
                authors_list.extend([a.strip() for a in r.authors.split(";") if a.strip()])
        auth_count = Counter(authors_list)

        report = []
        report.append("# Scopus Scientific Data Extraction Report")
        report.append(f"**Coleta realizada em**: {time.strftime('%d/%m/%Y %H:%M:%S')}")
        report.append(f"**Expressão de Busca**: `{query}`")
        report.append(f"**Total de artigos exclusivos recuperados**: {total}\n")
        
        report.append("## 1. Cronologia de Publicações")
        report.append("| Ano | Quantidade | Percentual |")
        report.append("| --- | --- | --- |")
        for yr, count in sorted(years_count.items(), reverse=True):
            pct = (count / total) * 100
            report.append(f"| {yr} | {count} | {pct:.1f}% |")
        if not years:
            report.append("| N/A | 0 | 0.0% |")
        report.append("")
        
        report.append("## 2. Periódicos mais Frequentes (Top 10)")
        report.append("| Periódico | Artigos | Percentual |")
        report.append("| --- | --- | --- |")
        for j, count in journal_count.most_common(10):
            pct = (count / total) * 100
            report.append(f"| {j} | {count} | {pct:.1f}% |")
        report.append("")
        
        report.append("## 3. Distribuição por Idiomas de Resumos (Enriquecidos)")
        report.append("| Idioma | Ocorrências | Percentual |")
        report.append("| --- | --- | --- |")
        total_lang = sum(lang_count.values()) if lang_count else 1
        for lang, count in lang_count.most_common():
            pct = (count / total_lang) * 100
            report.append(f"| {lang} | {count} | {pct:.1f}% |")
        report.append("")
        
        report.append("## 4. Autores mais Frequentes (Top 10)")
        report.append("| Autor | Artigos |")
        report.append("| --- | --- |")
        for auth, count in auth_count.most_common(10):
            report.append(f"| {auth} | {count} |")
        report.append("")

        report.append("## 5. Artigos mais Citados")
        report.append("| Título | Periódico | Ano | Citações | DOI |")
        report.append("| --- | --- | --- | --- | --- |")
        sorted_records = sorted(records, key=lambda x: x.citations, reverse=True)
        for r in sorted_records[:10]:
            report.append(f"| {r.title} | {r.journal_name} | {r.year} | {r.citations} | {r.doi} |")
        report.append("")
        
        with open(report_path, "w", encoding="utf-8") as f:
            f.write("\n".join(report))

def main():
    # Setup argparse CLI
    parser = argparse.ArgumentParser(description="Scopus Scientific Data Harvester - CLI Tool")
    parser.add_argument("--query", type=str, help="Search query expression for Scopus (e.g. 'TITLE-ABS-KEY(\"inferencia causal\")')")
    parser.add_argument("--start-year", type=int, help="Start year constraint (e.g. 2000)")
    parser.add_argument("--end-year", type=int, help="End year constraint (e.g. 2026)")
    parser.add_argument("--output-dir", type=str, help="Directory to save output files")
    parser.add_argument("--config", type=str, default="config_scopus.json", help="Path to config JSON file")
    args = parser.parse_args()

    # Load configuration
    config = {}
    if os.path.exists(args.config):
        try:
            with open(args.config, "r", encoding="utf-8") as f:
                config = json.load(f)
        except Exception as e:
            print(f"Warning: Failed to load config file '{args.config}': {e}. Using CLI arguments/defaults.")

    # Resolve settings
    search_cfg = config.get("search", {})
    query = args.query or search_cfg.get("query")
    
    if not query:
        print("\n[ERROR] No search query provided! Please configure it in config_scopus.json or use the --query argument.")
        parser.print_help()
        return

    start_year = args.start_year or search_cfg.get("start_year") or 2000
    end_year = args.end_year or search_cfg.get("end_year") or 2026

    paths = config.get("paths", {})
    output_dir = args.output_dir or paths.get("output_dir") or "scopus_outputs"
    log_name = paths.get("log_name") or "scopus_harvester.log"
    csv_name = paths.get("csv_name") or "scopus_clean_data.csv"
    json_name = paths.get("json_name") or "scopus_raw_backup.json"
    excel_name = paths.get("excel_name") or "Scopus_Data_Export.xlsx"
    report_name = paths.get("report_name") or "scopus_summary_report.md"

    # Setup output folder and logger
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    logger = setup_logging(os.path.join(output_dir, log_name))
    logger.info("=== SCOPUS SYSTEM DATA HARVESTER STARTED ===")
    logger.info(f"Target query: '{query}'")
    logger.info(f"Temporal window: {start_year} - {end_year}")
    logger.info(f"Saving outputs to: {output_dir}")
    
    api_cfg = config.get("api", {})
    api_key = api_cfg.get("api_key")
    if not api_key:
        logger.error("Scopus API Key is missing! Please configure 'api_key' in the config_scopus.json file.")
        return
        
    base_url = api_cfg.get("base_url") or "https://api.elsevier.com/content/search/scopus"
    limit = api_cfg.get("limit") or 25
    user_agent = api_cfg.get("user_agent") or "ScopusHarvester/1.0"
    politeness_delay = api_cfg.get("politeness_delay_seconds") or 1.0
    max_retries = api_cfg.get("max_retries") or 5
    backoff_factor = api_cfg.get("backoff_factor") or 1.5

    client = ScopusClient(
        api_key=api_key,
        base_url=base_url,
        user_agent=user_agent,
        politeness_delay=politeness_delay,
        max_retries=max_retries,
        backoff_factor=backoff_factor
    )
    
    raw_entries: List[Dict[str, Any]] = []
    start = 0
    
    logger.info("Initiating search queries on Scopus API...")
    while True:
        logger.info(f"Requesting results starting from offset {start}...")
        response_data = client.search_page(query, start, limit)
        
        if not response_data:
            logger.error(f"Error fetching Scopus API page starting from offset {start}. Breaking loop.")
            break
            
        results = response_data.get("search-results", {})
        total_results = int(results.get("opensearch:totalResults", "0"))
        
        if start == 0:
            logger.info(f"Total matching records in Scopus index: {total_results}")
            
        entries = results.get("entry", [])
        if not entries:
            logger.info("No more records found in Scopus search. Harvesting complete.")
            break
            
        logger.info(f"Fetched {len(entries)} records starting from offset {start}.")
        raw_entries.extend(entries)
        
        # Break condition
        if len(entries) < limit or start + len(entries) >= total_results:
            logger.info("Finished harvesting all matches from Scopus.")
            break
            
        start += len(entries)
        time.sleep(client.politeness_delay)
        
    logger.info(f"Total unique records harvested: {len(raw_entries)}")
    
    # Process, validate and clean data
    enricher = MetadataEnricher(api_key, logger)
    cleaned_records: List[ScopusRecord] = []
    
    logger.info("Starting metadata enrichment and validation (abstracts lookup)...")
    for idx, entry in enumerate(raw_entries, 1):
        try:
            logger.info(f"Enriching record {idx}/{len(raw_entries)}: EID {entry.get('eid', 'N/A')}")
            record = ScopusRecord.from_scopus_raw(entry)
            
            # Enrich with abstract and other details (VPN direct Scopus or OpenAlex fallback)
            enriched_record = enricher.enrich_record(record)
            cleaned_record = DataCleaner.clean_record(enriched_record)
            
            # Year constraint check
            if start_year <= cleaned_record.year <= end_year:
                cleaned_records.append(cleaned_record)
            else:
                logger.debug(f"Record {cleaned_record.id} filtered out by year ({cleaned_record.year}).")
            
            # Politeness delay for metadata enrichment
            time.sleep(politeness_delay)
        except Exception as e:
            logger.error(f"Error processing record at index {idx}: {e}", exc_info=True)
            
    logger.info(f"Processing finished. {len(cleaned_records)} records passed validation filters.")
    
    # Export Data
    exporter = ScopusExporter(output_dir, logger)
    
    # JSON backup
    exporter.export_to_json(cleaned_records, json_name)
    
    # CSV clean dataset
    exporter.export_to_csv(cleaned_records, csv_name)
    
    # Excel new workbook
    exporter.export_to_excel(cleaned_records, excel_name)
    
    # Markdown analysis
    exporter.generate_analysis_report(cleaned_records, report_name, query)
    
    logger.info("=== SCOPUS SYSTEM DATA HARVESTER PROCESS COMPLETED SUCCESSFULLY ===")

if __name__ == "__main__":
    main()
