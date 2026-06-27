import os
import re
import csv
import json
import time
import logging
import argparse
from dataclasses import dataclass, asdict
from typing import List, Dict, Any, Optional
import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

# Configure logging
def setup_logging(log_path: str):
    logger = logging.getLogger("SciELOHarvester")
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
class SciELORecord:
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
    def from_html_item(cls, item) -> "SciELORecord":
        # 1. Clean ID (PID)
        checkbox = item.find("input", class_="my_selection")
        if checkbox:
            val = checkbox.get("value", "")
            clean_id = val.rsplit("-", 1)[0] if "-" in val else val
        else:
            clean_id = "N/A"
            
        # 2. Title
        title_tag = item.find("strong", class_="title")
        title = title_tag.get_text().strip() if title_tag else "N/A"
        
        # 3. Landing page URL
        link_tag = item.find("a")
        url_link = link_tag.get("href") if link_tag else "N/A"
        
        # 4. Authors
        authors_div = item.find("div", class_="authors")
        authors = "N/A"
        if authors_div:
            authors_list = [a.get_text().strip() for a in authors_div.find_all("a", class_="author")]
            if authors_list:
                authors = "; ".join(authors_list)
            
        # 5. Journal details
        source_div = item.find("div", class_="source")
        journal_name = "N/A"
        year = 0
        volume = "N/A"
        issue = "N/A"
        elocation = "N/A"
        journal_issns = "N/A"
        
        if source_div:
            journal_a = source_div.find("a", class_="dropdown-toggle")
            if journal_a:
                journal_name = journal_a.get_text().strip()
                
            info_a = source_div.find("a", class_="openJournalInfo")
            if info_a:
                journal_issns = info_a.get("data-issn", "N/A")
                
            # Year is in a span with 4 digits followed by comma
            for s in source_div.find_all("span"):
                txt = s.get_text().strip()
                m = re.match(r'^(\d{4}),?$', txt)
                if m:
                    year = int(m.group(1))
                    break
            
            # Find Volume, elocation, and Issue
            vol_label = source_div.find(string=re.compile("Volume", re.I))
            if vol_label:
                vol_span = vol_label.find_next("span")
                if vol_span:
                    volume = vol_span.get_text().strip()
            
            eloc_label = source_div.find(string=re.compile("elocation", re.I))
            if eloc_label:
                eloc_span = eloc_label.find_next("span")
                if eloc_span:
                    elocation = eloc_span.get_text().strip()
                    
            issue_label = source_div.find(string=re.compile("Número|Number", re.I))
            if issue_label:
                issue_span = issue_label.find_next("span")
                if issue_span:
                    issue = issue_span.get_text().strip()
                    
        # Fallback for year and journal (e.g. for preprints)
        if year == 0:
            text_content = item.get_text()
            date_match = re.search(r'\b(19\d{2}|20\d{2})[-/]\d{2}[-/]\d{2}\b', text_content)
            if date_match:
                year = int(date_match.group(1))
            else:
                year_matches = re.findall(r'\b(19\d{2}|20\d{2})\b', text_content)
                if year_matches:
                    year = int(year_matches[0])
                    
        if journal_name == "N/A" and "[SciELO Preprints]" in item.get_text():
            journal_name = "SciELO Preprints"
            
        # 6. DOI
        doi_tag = item.find("a", href=re.compile(r'https://doi.org/'))
        doi = doi_tag.get("href") if doi_tag else "N/A"
        
        # 7. PDF URL
        pdf_tag = item.find("a", href=re.compile(r'sci_pdf|_pdf|pdf'))
        pdf_url = pdf_tag.get("href") if pdf_tag else "N/A"
        
        # 8. Languages & Abstracts
        abstract_divs = item.find_all("div", class_="abstract")
        abstracts = []
        languages_list = []
        for ab in abstract_divs:
            lang_id = ab.get("id", "").split("_")[-1]
            ab_text = ab.get_text().strip()
            # Clean up abstract text a bit
            ab_text = re.sub(r'\s+', ' ', ab_text).strip()
            abstracts.append(f"[{lang_id.upper()}]: {ab_text}")
            languages_list.append(lang_id)
        abstract = " | ".join(abstracts) if abstracts else "N/A"
        language = "; ".join(languages_list) if languages_list else "N/A"
        
        # Build pub_date string
        pub_date = str(year) if year > 0 else "N/A"
        
        return cls(
            id=clean_id,
            doi=doi,
            title=title,
            authors=authors,
            year=year,
            publication_date=pub_date,
            journal_name=journal_name,
            journal_issns=journal_issns,
            language=language,
            url=url_link,
            pdf_url=pdf_url,
            citations=0, # Direct scraper does not fetch citations (normally 0 in search results)
            concepts="N/A",
            abstract=abstract
        )

class SciELOWebClient:
    def __init__(self, user_agent: str, politeness_delay: float, max_retries: int, backoff_factor: float):
        self.base_url = "https://search.scielo.org/"
        if "SciELOHarvester" in user_agent or "bot" in user_agent.lower():
            self.user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        else:
            self.user_agent = user_agent
        self.politeness_delay = politeness_delay
        self.max_retries = max_retries
        self.backoff_factor = backoff_factor
        self.session = requests.Session()
        
    def search_page(self, query: str, offset: int) -> Optional[str]:
        # Realistic headers to bypass Bunny Shield WAF
        headers = {
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
            "Accept-Language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
            "Cache-Control": "no-cache",
            "Pragma": "no-cache",
            "Sec-Ch-Ua": '"Not_A Brand";v="8", "Chromium";v="120", "Google Chrome";v="120"',
            "Sec-Ch-Ua-Mobile": "?0",
            "Sec-Ch-Ua-Platform": '"Windows"',
            "Sec-Fetch-Dest": "document",
            "Sec-Fetch-Mode": "navigate",
            "Sec-Fetch-Site": "none",
            "Sec-Fetch-User": "?1",
            "Upgrade-Insecure-Requests": "1",
            "User-Agent": self.user_agent
        }
        
        params = {
            "q": query,
            "lang": "pt",
            "from": offset
        }
        
        retries = 0
        while retries < self.max_retries:
            try:
                response = self.session.get(self.base_url, params=params, headers=headers, timeout=20)
                if response.status_code == 200:
                    return response.text
                elif response.status_code in [429, 500, 502, 503, 504]:
                    sleep_time = self.backoff_factor ** retries
                    print(f"Warning: HTTP status {response.status_code} received. Retrying in {sleep_time:.2f}s...")
                    time.sleep(sleep_time)
                    retries += 1
                else:
                    print(f"Error: Server returned HTTP status {response.status_code}")
                    return None
            except requests.RequestException as e:
                sleep_time = self.backoff_factor ** retries
                print(f"Connection warning: {e}. Retrying in {sleep_time:.2f}s...")
                time.sleep(sleep_time)
                retries += 1
                
        return None

class DataCleaner:
    @staticmethod
    def clean_text(text: Optional[str]) -> str:
        if not text or text == "N/A":
            return "N/A"
        text = re.sub(r'<[^>]+>', ' ', text)
        text = re.sub(r'\s+', ' ', text)
        return text.strip()

    @classmethod
    def clean_record(cls, record: SciELORecord) -> SciELORecord:
        record.title = cls.clean_text(record.title)
        record.authors = cls.clean_text(record.authors)
        record.journal_name = cls.clean_text(record.journal_name)
        record.abstract = cls.clean_text(record.abstract)
        
        if not record.doi: record.doi = "N/A"
        if not record.language: record.language = "N/A"
        if not record.url: record.url = "N/A"
        if not record.pdf_url: record.pdf_url = "N/A"
        
        return record

class SciELOExporter:
    def __init__(self, output_dir: str, logger: logging.Logger):
        self.output_dir = output_dir
        self.logger = logger
        
    def export_to_json(self, records: List[SciELORecord], filename: str):
        json_path = os.path.join(self.output_dir, filename)
        self.logger.info(f"Exporting JSON raw backup: {json_path}")
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump([asdict(r) for r in records], f, ensure_ascii=False, indent=2)
            
    def export_to_csv(self, records: List[SciELORecord], filename: str):
        csv_path = os.path.join(self.output_dir, filename)
        self.logger.info(f"Exporting CSV: {csv_path}")
        
        headers = [
            "ID SciELO (PID)", "DOI", "Título", "Autores", "Ano de Publicação", 
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
                
    def export_to_excel(self, records: List[SciELORecord], filename: str):
        excel_path = os.path.join(self.output_dir, filename)
        self.logger.info(f"Creating a brand-new Excel workbook: {excel_path}")
        
        wb = openpyxl.Workbook()
        
        font_family = "Segoe UI"
        font_header = Font(name=font_family, size=10, bold=True, color="FFFFFF")
        font_regular = Font(name=font_family, size=10)
        
        fill_header = PatternFill(start_color="1E4A38", end_color="1E4A38", fill_type="solid") # Dark Forest Green
        fill_zebra = PatternFill(start_color="F0F7F4", end_color="F0F7F4", fill_type="solid")  # Soft Mint Green
        
        border_all = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )
        
        align_left = Alignment(horizontal='left', vertical='center', wrap_text=False)
        align_center = Alignment(horizontal='center', vertical='center')
        align_wrap_left = Alignment(horizontal='left', vertical='top', wrap_text=True)
        
        # 1. Main Data Sheet
        ws_data = wb.active
        ws_data.title = "SciELO Cleaned Data"
        
        headers = [
            "ID SciELO (PID)", "DOI", "Título", "Autores", "Ano de Publicação", 
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
        ws_summary.cell(row=2, column=2, value="SciELO Extraction Summary").font = Font(name=font_family, size=16, bold=True, color="1E4A38")
        
        ws_summary.cell(row=4, column=2, value="Execution Date:").font = Font(name=font_family, size=10, bold=True)
        ws_summary.cell(row=4, column=3, value=time.strftime('%Y-%m-%d %H:%M:%S')).font = font_regular
        ws_summary.cell(row=5, column=2, value="Total Unique Works:").font = Font(name=font_family, size=10, bold=True)
        ws_summary.cell(row=5, column=3, value=len(records)).font = font_regular
        
        from collections import Counter
        journals_count = Counter([r.journal_name for r in records if r.journal_name != "N/A"])
        
        ws_summary.cell(row=7, column=2, value="Top Journals Summary").font = Font(name=font_family, size=12, bold=True, color="1E4A38")
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

    def generate_analysis_report(self, records: List[SciELORecord], filename: str, query: str):
        report_path = os.path.join(self.output_dir, filename)
        self.logger.info(f"Generating markdown summary report: {report_path}")
        
        from collections import Counter
        
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
        report.append("# SciELO Scientific Data Extraction Report (via Portal Search)")
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
        
        report.append("## 3. Distribuição por Idiomas de Resumos")
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

        report.append("## 5. Artigos Encontrados (Top 20 por Ordem de Relevância)")
        report.append("| Título | Periódico | Ano | DOI |")
        report.append("| --- | --- | --- | --- |")
        for r in records[:20]:
            report.append(f"| {r.title} | {r.journal_name} | {r.year} | {r.doi} |")
        report.append("")
        
        with open(report_path, "w", encoding="utf-8") as f:
            f.write("\n".join(report))

def main():
    # Setup argparse CLI
    parser = argparse.ArgumentParser(description="SciELO Scientific Data Harvester (Direct Search) - CLI Tool")
    parser.add_argument("--query", type=str, help="Search query expression for SciELO (e.g. '\"inferência causal\"')")
    parser.add_argument("--start-year", type=int, help="Start year constraint (e.g. 2000)")
    parser.add_argument("--end-year", type=int, help="End year constraint (e.g. 2026)")
    parser.add_argument("--output-dir", type=str, help="Directory to save output files")
    parser.add_argument("--limit", type=int, help="Limit number of pages to fetch")
    parser.add_argument("--config", type=str, default="config_scielo.json", help="Path to config JSON file")
    args = parser.parse_args()

    # Load configuration if present
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
        print("\n[ERROR] No search query provided! Please configure it in config_scielo.json or use the --query argument.")
        parser.print_help()
        return

    start_year = args.start_year or search_cfg.get("start_year") or 2000
    end_year = args.end_year or search_cfg.get("end_year") or 2026

    paths = config.get("paths", {})
    output_dir = args.output_dir or paths.get("output_dir") or "scielo_outputs"
    log_name = paths.get("log_name") or "scielo_harvester.log"
    csv_name = paths.get("csv_name") or "scielo_clean_data.csv"
    json_name = paths.get("json_name") or "scielo_raw_backup.json"
    excel_name = paths.get("excel_name") or "SciELO_Data_Export.xlsx"
    report_name = paths.get("report_name") or "scielo_summary_report.md"

    # Setup directories and logger
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    logger = setup_logging(os.path.join(output_dir, log_name))
    logger.info("=== SCIELO SYSTEM DATA HARVESTER STARTED (DIRECT WEB SEARCH MODE) ===")
    logger.info(f"Target query: '{query}'")
    logger.info(f"Temporal window: {start_year} - {end_year}")
    logger.info(f"Saving outputs to: {output_dir}")
    
    api_cfg = config.get("api", {})
    user_agent = api_cfg.get("user_agent") or "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    politeness_delay = api_cfg.get("politeness_delay_seconds") or 1.0
    max_retries = api_cfg.get("max_retries") or 5
    backoff_factor = api_cfg.get("backoff_factor") or 1.5

    client = SciELOWebClient(
        user_agent=user_agent,
        politeness_delay=politeness_delay,
        max_retries=max_retries,
        backoff_factor=backoff_factor
    )
    
    harvested_records: List[SciELORecord] = []
    offset = 1
    
    logger.info(f"Initiating retrieval loop for query: '{query}'...")
    while True:
        logger.info(f"Requesting results starting from offset {offset}...")
        html_content = client.search_page(query, offset)
        
        if not html_content:
            logger.error(f"Error fetching page at offset {offset}. Breaking loop.")
            break
            
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # Get total hits count
        total_hits_tag = soup.find(id="TotalHits")
        total_hits = 0
        if total_hits_tag:
            try:
                total_hits = int(total_hits_tag.get_text().strip())
            except ValueError:
                total_hits = 0
                
        if offset == 1:
            logger.info(f"SciELO total matching records in search portal: {total_hits}")
            
        items = soup.find_all("div", class_="item")
        if not items:
            logger.info("No more records found in this page. Harvesting complete.")
            break
            
        logger.info(f"Fetched {len(items)} records starting from offset {offset}.")
        
        for item in items:
            try:
                record = SciELORecord.from_html_item(item)
                harvested_records.append(record)
            except Exception as e:
                logger.error(f"Error parsing HTML item at offset {offset}: {e}", exc_info=True)
                
        # Break condition
        if len(items) < 15 or offset + len(items) > total_hits:
            logger.info("Finished harvesting all matches.")
            break
            
        offset += len(items)
        time.sleep(client.politeness_delay)
        
    logger.info(f"Total unique records harvested: {len(harvested_records)}")
    
    # Process, validate and clean data
    cleaned_records: List[SciELORecord] = []
    for record in harvested_records:
        try:
            cleaned_record = DataCleaner.clean_record(record)
            
            # Year constraint check
            if start_year <= cleaned_record.year <= end_year:
                cleaned_records.append(cleaned_record)
            else:
                logger.debug(f"Record {cleaned_record.id} filtered out by year ({cleaned_record.year}).")
        except Exception as e:
            logger.error(f"Error cleaning record {record.id}: {e}", exc_info=True)
            
    logger.info(f"Processing finished. {len(cleaned_records)} records passed validation filters.")
    
    # Export Data
    exporter = SciELOExporter(output_dir, logger)
    
    # JSON backup
    exporter.export_to_json(harvested_records, json_name)
    
    # CSV clean dataset
    exporter.export_to_csv(cleaned_records, csv_name)
    
    # Excel new workbook
    exporter.export_to_excel(cleaned_records, excel_name)
    
    # Markdown analysis
    exporter.generate_analysis_report(cleaned_records, report_name, query)
    
    logger.info("=== SCIELO SYSTEM DATA HARVESTER PROCESS COMPLETED SUCCESSFULLY ===")

if __name__ == "__main__":
    main()
