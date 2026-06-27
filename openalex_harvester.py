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
    logger = logging.getLogger("OpenAlexHarvester")
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
class OpenAlexRecord:
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
    original_sources: str
    abstract: str

    @classmethod
    def from_openalex_raw(cls, item: Dict[str, Any]) -> "OpenAlexRecord":
        # Extract ID (extracting only the ID part from URL if present)
        raw_id = item.get("id") or "N/A"
        clean_id = raw_id.split("/")[-1] if "/" in raw_id else raw_id
        
        # Extract DOI
        doi = item.get("doi") or "N/A"
        
        # Extract Title
        title = item.get("title") or "N/A"
        
        # Extract Authors
        authorships = item.get("authorships") or []
        authors_list = []
        for auth in authorships:
            author_obj = auth.get("author") or {}
            name = author_obj.get("display_name")
            if name:
                authors_list.append(name)
        authors = "; ".join(authors_list) if authors_list else "N/A"
        
        # Extract Year & Date
        year = item.get("publication_year") or 0
        pub_date = item.get("publication_date") or "N/A"
        
        # Extract Journal / Host source
        primary_loc = item.get("primary_location") or {}
        source_obj = primary_loc.get("source") or {}
        journal_name = source_obj.get("display_name") or "N/A"
        
        issns = source_obj.get("issn") or []
        journal_issns = "; ".join(issns) if issns else "N/A"
        
        # Extract Language
        language = item.get("language") or "N/A"
        
        # URL (primary landing page URL)
        url = primary_loc.get("landing_page_url") or "N/A"
        if url == "N/A":
            url = raw_id
            
        # PDF URL
        open_access = item.get("open_access") or {}
        pdf_url = open_access.get("oa_url") or "N/A"
        
        # Citations
        citations = item.get("cited_by_count") or 0
        
        # Extract UNIQUE display names of all indexing sources/repositories
        locations = item.get("locations") or []
        sources_seen = set()
        for loc in locations:
            src = loc.get("source") or {}
            src_name = src.get("display_name")
            if src_name:
                sources_seen.add(src_name)
        original_sources = "; ".join(sorted(sources_seen)) if sources_seen else "N/A"
        
        # Reconstruct abstract from inverted index
        abstract = "N/A"
        abstract_inverted = item.get("abstract_inverted_index") or {}
        if abstract_inverted:
            try:
                word_positions = []
                for word, positions in abstract_inverted.items():
                    for pos in positions:
                        word_positions.append((pos, word))
                word_positions.sort()
                abstract = " ".join([word for _, word in word_positions])
            except Exception:
                abstract = "N/A"
                
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
            url=url,
            pdf_url=pdf_url,
            citations=citations,
            original_sources=original_sources,
            abstract=abstract
        )

class OpenAlexClient:
    def __init__(self, base_url: str, user_agent: str, politeness_delay: float, max_retries: int, backoff_factor: float):
        self.base_url = base_url
        self.user_agent = user_agent
        self.politeness_delay = politeness_delay
        self.max_retries = max_retries
        self.backoff_factor = backoff_factor
        self.session = requests.Session()

    def build_filter_string(self, query: str, start_year: int, end_year: int, filters_config: Dict[str, Any]) -> str:
        filters_list = []
        
        # Text query filter specifically targeting Title and Abstract
        filters_list.append(f"title_and_abstract.search:{query}")
        
        # Year filter
        filters_list.append(f"publication_year:{start_year}-{end_year}")
        
        # Custom repository IDs filter
        rep_ids = filters_config.get("repository_ids") or []
        if rep_ids:
            # Format: locations.source.id:S123|S456
            formatted_ids = "|".join(rep_ids)
            filters_list.append(f"locations.source.id:{formatted_ids}")
            
        # Custom publisher IDs filter
        pub_ids = filters_config.get("publisher_ids") or []
        if pub_ids:
            formatted_ids = "|".join(pub_ids)
            filters_list.append(f"primary_location.source.publisher_lineage:{formatted_ids}")
            
        # Open Access filter
        if filters_config.get("only_open_access"):
            filters_list.append("is_oa:true")
            
        # Source types filter (e.g. journal, repository)
        src_types = filters_config.get("source_types") or []
        if src_types:
            formatted_types = "|".join(src_types)
            filters_list.append(f"locations.source.type:{formatted_types}")
            
        return ",".join(filters_list)

    def fetch_page(self, filter_str: str, page: int, per_page: int) -> Optional[Dict[str, Any]]:
        headers = {
            "User-Agent": self.user_agent
        }
        
        params = {
            "filter": filter_str,
            "page": page,
            "per_page": per_page
        }
        
        retries = 0
        while retries < self.max_retries:
            try:
                response = self.session.get(self.base_url, params=params, headers=headers, timeout=20)
                if response.status_code == 200:
                    return response.json()
                elif response.status_code in [429, 500, 502, 503, 504]:
                    sleep_time = self.backoff_factor ** retries
                    print(f"Warning: OpenAlex API status {response.status_code}. Retrying in {sleep_time:.2f}s...")
                    time.sleep(sleep_time)
                    retries += 1
                else:
                    print(f"Error: OpenAlex API returned HTTP status {response.status_code}: {response.text}")
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
        # Remove HTML tags and normalize spacing
        text = re.sub(r'<[^>]+>', ' ', text)
        text = re.sub(r'\s+', ' ', text)
        return text.strip()

    @classmethod
    def clean_record(cls, record: OpenAlexRecord) -> OpenAlexRecord:
        record.title = cls.clean_text(record.title)
        record.authors = cls.clean_text(record.authors)
        record.journal_name = cls.clean_text(record.journal_name)
        record.original_sources = cls.clean_text(record.original_sources)
        record.abstract = cls.clean_text(record.abstract)
        
        if not record.doi: record.doi = "N/A"
        if not record.language: record.language = "N/A"
        if not record.url: record.url = "N/A"
        if not record.pdf_url: record.pdf_url = "N/A"
        
        return record

class OpenAlexExporter:
    def __init__(self, output_dir: str, logger: logging.Logger):
        self.output_dir = output_dir
        self.logger = logger
        
    def export_to_json(self, records: List[OpenAlexRecord], filename: str):
        json_path = os.path.join(self.output_dir, filename)
        self.logger.info(f"Exporting JSON raw backup: {json_path}")
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump([asdict(r) for r in records], f, ensure_ascii=False, indent=2)
            
    def export_to_csv(self, records: List[OpenAlexRecord], filename: str):
        csv_path = os.path.join(self.output_dir, filename)
        self.logger.info(f"Exporting CSV: {csv_path}")
        
        headers = [
            "ID OpenAlex", "DOI", "Título", "Autores", "Ano de Publicação", 
            "Data de Publicação", "Periódico (Nome)", "Periódico (ISSN)", 
            "Idioma", "Link / URL", "PDF URL", "Citações", "Fontes / Bases Originais", "Resumo"
        ]
        
        with open(csv_path, "w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(headers)
            for r in records:
                writer.writerow([
                    r.id, r.doi, r.title, r.authors, r.year,
                    r.publication_date, r.journal_name, r.journal_issns,
                    r.language, r.url, r.pdf_url, r.citations, r.original_sources, r.abstract
                ])
                
    def export_to_excel(self, records: List[OpenAlexRecord], filename: str):
        excel_path = os.path.join(self.output_dir, filename)
        self.logger.info(f"Creating a brand-new Excel workbook: {excel_path}")
        
        wb = openpyxl.Workbook()
        
        font_family = "Segoe UI"
        font_header = Font(name=font_family, size=10, bold=True, color="FFFFFF")
        font_regular = Font(name=font_family, size=10)
        
        # Royal Purple branding color palette
        fill_header = PatternFill(start_color="6C3082", end_color="6C3082", fill_type="solid") # Royal Purple
        fill_zebra = PatternFill(start_color="F9F6FF", end_color="F9F6FF", fill_type="solid")  # Soft Purple White
        
        border_all = Border(
            left=Side(style='thin', color='E5E0EB'),
            right=Side(style='thin', color='E5E0EB'),
            top=Side(style='thin', color='E5E0EB'),
            bottom=Side(style='thin', color='E5E0EB')
        )
        
        align_left = Alignment(horizontal='left', vertical='center', wrap_text=False)
        align_center = Alignment(horizontal='center', vertical='center')
        align_wrap_left = Alignment(horizontal='left', vertical='top', wrap_text=True)
        
        # 1. Data Sheet
        ws_data = wb.active
        ws_data.title = "OpenAlex Cleaned Data"
        
        headers = [
            "ID OpenAlex", "DOI", "Título", "Autores", "Ano de Publicação", 
            "Data de Publicação", "Periódico (Nome)", "Periódico (ISSN)", 
            "Idioma", "Link / URL", "PDF URL", "Citações", "Fontes / Bases Originais", "Resumo"
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
                r.language, r.url, r.pdf_url, r.citations, r.original_sources, r.abstract
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
                elif col_idx in [3, 4, 7, 13]:  # Title, Authors, Journal, Original Sources
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
        ws_summary.cell(row=2, column=2, value="OpenAlex Extraction Summary").font = Font(name=font_family, size=16, bold=True, color="6C3082")
        
        ws_summary.cell(row=4, column=2, value="Execution Date:").font = Font(name=font_family, size=10, bold=True)
        ws_summary.cell(row=4, column=3, value=time.strftime('%Y-%m-%d %H:%M:%S')).font = font_regular
        ws_summary.cell(row=5, column=2, value="Total Unique Works:").font = Font(name=font_family, size=10, bold=True)
        ws_summary.cell(row=5, column=3, value=len(records)).font = font_regular
        
        # Compile original databases count
        sources_list = []
        for r in records:
            if r.original_sources and r.original_sources != "N/A":
                sources_list.extend([s.strip() for s in r.original_sources.split(";") if s.strip()])
        sources_count = Counter(sources_list)
        
        ws_summary.cell(row=7, column=2, value="Top Indexing Databases / Repositories").font = Font(name=font_family, size=12, bold=True, color="6C3082")
        ws_summary.cell(row=8, column=2, value="Source / Database Name").font = font_header
        ws_summary.cell(row=8, column=2).fill = fill_header
        ws_summary.cell(row=8, column=3, value="Count").font = font_header
        ws_summary.cell(row=8, column=3).fill = fill_header
        
        cur_row = 9
        for src_name, count in sources_count.most_common(20):
            ws_summary.row_dimensions[cur_row].height = 20
            c_name = ws_summary.cell(row=cur_row, column=2, value=src_name)
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

    def generate_analysis_report(self, records: List[OpenAlexRecord], filename: str, query: str):
        report_path = os.path.join(self.output_dir, filename)
        self.logger.info(f"Generating markdown summary report: {report_path}")
        
        total = len(records)
        years = [r.year for r in records if r.year > 0]
        journals = [r.journal_name for r in records if r.journal_name != "N/A"]
        
        languages_flat = []
        for r in records:
            if r.language and r.language != "N/A":
                languages_flat.append(r.language.strip())
                
        years_count = Counter(years)
        journal_count = Counter(journals)
        lang_count = Counter(languages_flat)
        
        authors_list = []
        for r in records:
            if r.authors and r.authors != "N/A":
                authors_list.extend([a.strip() for a in r.authors.split(";") if a.strip()])
        auth_count = Counter(authors_list)

        # Count original databases
        sources_list = []
        for r in records:
            if r.original_sources and r.original_sources != "N/A":
                sources_list.extend([s.strip() for s in r.original_sources.split(";") if s.strip()])
        sources_count = Counter(sources_list)

        report = []
        report.append("# OpenAlex Scientific Data Extraction Report")
        report.append(f"**Coleta realizada em**: {time.strftime('%d/%m/%Y %H:%M:%S')}")
        report.append(f"**Expressão de Busca**: `{query}`")
        report.append(f"**Total de artigos exclusivos recuperados**: {total}\n")
        
        report.append("## 1. Origem Bibliométrica: Fontes e Bases de Origem")
        report.append("| Base / Repositório Original | Artigos Indexados | Percentual de Cobertura |")
        report.append("| --- | --- | --- |")
        for src_name, count in sources_count.most_common(20):
            pct = (count / total) * 100
            report.append(f"| {src_name} | {count} | {pct:.1f}% |")
        report.append("")
        
        report.append("## 2. Cronologia de Publicações")
        report.append("| Ano | Quantidade | Percentual |")
        report.append("| --- | --- | --- |")
        for yr, count in sorted(years_count.items(), reverse=True):
            pct = (count / total) * 100
            report.append(f"| {yr} | {count} | {pct:.1f}% |")
        if not years:
            report.append("| N/A | 0 | 0.0% |")
        report.append("")
        
        report.append("## 3. Periódicos/Revistas mais Frequentes (Top 10)")
        report.append("| Periódico | Artigos | Percentual |")
        report.append("| --- | --- | --- |")
        for j, count in journal_count.most_common(10):
            pct = (count / total) * 100
            report.append(f"| {j} | {count} | {pct:.1f}% |")
        report.append("")
        
        report.append("## 4. Distribuição por Idiomas de Resumos")
        report.append("| Idioma | Ocorrências | Percentual |")
        report.append("| --- | --- | --- |")
        total_lang = sum(lang_count.values()) if lang_count else 1
        for lang, count in lang_count.most_common():
            pct = (count / total_lang) * 100
            report.append(f"| {lang} | {count} | {pct:.1f}% |")
        report.append("")
        
        report.append("## 5. Autores mais Frequentes (Top 10)")
        report.append("| Autor | Artigos |")
        report.append("| --- | --- |")
        for auth, count in auth_count.most_common(10):
            report.append(f"| {auth} | {count} |")
        report.append("")

        report.append("## 6. Artigos mais Citados")
        report.append("| Título | Periódico | Ano | Citações | DOI |")
        report.append("| --- | --- | --- | --- | --- |")
        sorted_records = sorted(records, key=lambda x: x.citations, reverse=True)
        for r in sorted_records[:15]:
            report.append(f"| {r.title} | {r.journal_name} | {r.year} | {r.citations} | {r.doi} |")
        report.append("")
        
        with open(report_path, "w", encoding="utf-8") as f:
            f.write("\n".join(report))

def main():
    # Setup argparse CLI
    parser = argparse.ArgumentParser(description="OpenAlex Scientific Data Harvester - CLI Tool")
    parser.add_argument("--query", type=str, help="Search query expression (e.g. '\"inferencia causal\"')")
    parser.add_argument("--start-year", type=int, help="Start year constraint (e.g. 2000)")
    parser.add_argument("--end-year", type=int, help="End year constraint (e.g. 2026)")
    parser.add_argument("--output-dir", type=str, help="Directory to save output files")
    parser.add_argument("--config", type=str, default="config_openalex.json", help="Path to config JSON file")
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
        print("\n[ERROR] No search query provided! Please configure it in config_openalex.json or use the --query argument.")
        parser.print_help()
        return

    start_year = args.start_year or search_cfg.get("start_year") or 2000
    end_year = args.end_year or search_cfg.get("end_year") or 2026

    paths = config.get("paths", {})
    output_dir = args.output_dir or paths.get("output_dir") or "openalex_outputs"
    log_name = paths.get("log_name") or "openalex_harvester.log"
    csv_name = paths.get("csv_name") or "openalex_clean_data.csv"
    json_name = paths.get("json_name") or "openalex_raw_backup.json"
    excel_name = paths.get("excel_name") or "OpenAlex_Data_Export.xlsx"
    report_name = paths.get("report_name") or "openalex_summary_report.md"

    # Setup output folder and logger
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    logger = setup_logging(os.path.join(output_dir, log_name))
    logger.info("=== OPENALEX SYSTEM DATA HARVESTER STARTED ===")
    logger.info(f"Target query: '{query}'")
    logger.info(f"Temporal window: {start_year} - {end_year}")
    logger.info(f"Saving outputs to: {output_dir}")
    
    api_cfg = config.get("api", {})
    base_url = api_cfg.get("base_url") or "https://api.openalex.org/works"
    limit = api_cfg.get("limit") or 50
    user_agent = api_cfg.get("user_agent") or "OpenAlexHarvester/1.0 (contact: user@mail.com)"
    politeness_delay = api_cfg.get("politeness_delay_seconds") or 1.0
    max_retries = api_cfg.get("max_retries") or 5
    backoff_factor = api_cfg.get("backoff_factor") or 1.5

    client = OpenAlexClient(
        base_url=base_url,
        user_agent=user_agent,
        politeness_delay=politeness_delay,
        max_retries=max_retries,
        backoff_factor=backoff_factor
    )
    
    filters_cfg = search_cfg.get("filters") or {}
    filter_str = client.build_filter_string(query, start_year, end_year, filters_cfg)
    logger.info(f"Query filters: {filter_str}")

    raw_items: List[Dict[str, Any]] = []
    page = 1
    
    logger.info("Initiating search queries on OpenAlex API...")
    while True:
        logger.info(f"Requesting results starting page {page}...")
        response_data = client.fetch_page(filter_str, page, limit)
        
        if not response_data:
            logger.error(f"Error fetching OpenAlex API page {page}. Breaking loop.")
            break
            
        results = response_data.get("results", [])
        meta = response_data.get("meta", {})
        total_results = int(meta.get("count", "0"))
        
        if page == 1:
            logger.info(f"Total matching records in OpenAlex catalog: {total_results}")
            
        if not results:
            logger.info("No more records found in OpenAlex search. Harvesting complete.")
            break
            
        logger.info(f"Fetched {len(results)} records from page {page}.")
        raw_items.extend(results)
        
        # Break condition
        if len(results) < limit or len(raw_items) >= total_results:
            logger.info("Finished harvesting all matches from OpenAlex.")
            break
            
        page += 1
        time.sleep(client.politeness_delay)
        
    logger.info(f"Total unique records harvested: {len(raw_items)}")
    
    # Process, validate and clean data
    cleaned_records: List[OpenAlexRecord] = []
    
    logger.info("Starting data mapping and cleaning...")
    for idx, item in enumerate(raw_items, 1):
        try:
            record = OpenAlexRecord.from_openalex_raw(item)
            cleaned_record = DataCleaner.clean_record(record)
            
            # double check year window constraint
            if start_year <= cleaned_record.year <= end_year:
                cleaned_records.append(cleaned_record)
            else:
                logger.debug(f"Record {cleaned_record.id} filtered out by year ({cleaned_record.year}).")
        except Exception as e:
            logger.error(f"Error processing record at index {idx}: {e}", exc_info=True)
            
    logger.info(f"Processing finished. {len(cleaned_records)} records passed validation filters.")
    
    # Export Data
    exporter = OpenAlexExporter(output_dir, logger)
    
    # JSON backup
    exporter.export_to_json(cleaned_records, json_name)
    
    # CSV clean dataset
    exporter.export_to_csv(cleaned_records, csv_name)
    
    # Excel new workbook
    exporter.export_to_excel(cleaned_records, excel_name)
    
    # Markdown analysis
    exporter.generate_analysis_report(cleaned_records, report_name, query)
    
    logger.info("=== OPENALEX SYSTEM DATA HARVESTER PROCESS COMPLETED SUCCESSFULLY ===")

if __name__ == "__main__":
    main()
