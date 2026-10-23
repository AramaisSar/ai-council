import streamlit as st
from openai import OpenAI
import PyPDF2
import docx
import json
import os
import tempfile
import base64
import requests
from datetime import datetime, timedelta
import uuid
from io import BytesIO
from pathlib import Path
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import pandas as pd
import yfinance as yf

# ================= НАСТРОЙКИ =================
KEY_FILE = "key.txt"
HISTORY_FILE = "history.json"
PROJECTS_FILE = "projects.json"
LIBRARY_INDEX_FILE = "library_index.json"
BASE_URL = "https://api.vsegpt.ru/v1"
MODEL_CHATGPT = "openai/gpt-4o-mini"
MODEL_CLAUDE  = "anthropic/claude-3-5-haiku"
MODEL_QWEN    = "qwen/qwen-plus"
MAX_FILE_SIZE = 50000 
MAX_ARTICLES = 5
MAX_BOOK_PAGES = 30
MAX_EXCEL_ROWS = 100
# =============================================

# --- КАСТОМНЫЙ CSS ---
st.markdown("""
<style>
    .stApp { background-color: #FAF8F5; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    .stApp p, .stApp div, .stApp label { color: #4A403A !important; }
    h1, h2, h3 { color: #8B5E3C !important; font-weight: 600 !important; }
    section[data-testid="stSidebar"] { background-color: #FDFBF9 !important; border-right: 1px solid #F0EBE3 !important; }
    
    .stButton>button {
        background-color: #D4A373 !important; color: #FFFFFF !important; border: none !important;
        border-radius: 12px !important; padding: 10px 20px !important; font-weight: 600 !important;
        transition: all 0.3s ease !important; box-shadow: 0 2px 10px rgba(212, 163, 115, 0.2) !important;
    }
    .stButton>button:hover { background-color: #BC8A5F !important; transform: translateY(-1px) !important; }
    
    .ai-card {
        background: #FFFFFF; border-radius: 16px; padding: 24px; margin: 20px 0;
        box-shadow: 0 4px 24px rgba(139, 94, 60, 0.06); border: 1px solid #F0EBE3;
    }
    .card-chatgpt { border-left: 5px solid #E9C46A; }
    .card-claude { border-left: 5px solid #E07A5F; }
    .card-qwen { border-left: 5px solid #8B5E3C; background-color: #FAF6F1; }
    .card-external { border-left: 5px solid #6B9080; background-color: #F4F9F7; }
    .card-search { border-left: 5px solid #5B8DB8; background-color: #F0F7FB; }
    .card-library { border-left: 5px solid #9B59B6; background-color: #F8F0FB; }
    .card-excel { border-left: 5px solid #27AE60; background-color: #F0FBF4; }
    .card-finance { border-left: 5px solid #E74C3C; background-color: #FDF0EF; }
    .card-api { border-left: 5px solid #3498DB; background-color: #EBF5FB; }

    .stTextArea textarea { border-radius: 12px !important; border: 1px solid #E0D8CC !important; background: #FFFFFF !important; color: #4A403A !important; padding: 15px !important; }
    
    .stSuccess { background-color: #F1F8E9 !important; color: #558B2F !important; border-radius: 10px !important; border: 1px solid #DCEDC8 !important; }
    .stError { background-color: #FBE9E7 !important; color: #D84315 !important; border-radius: 10px !important; border: 1px solid #FFCCBC !important; }
    .stWarning { background-color: #FFF8E1 !important; color: #F57F17 !important; border-radius: 10px !important; border: 1px solid #FFE082 !important; }
    .stInfo { background-color: #E3F2FD !important; color: #1565C0 !important; border-radius: 10px !important; border: 1px solid #BBDEFB !important; }
    
    hr { border: none !important; height: 1px !important; background: #E0D8CC !important; margin: 30px 0 !important; }
    
    .article-item, .book-item {
        background: #FFFFFF; border-radius: 12px; padding: 16px; margin: 12px 0;
        border: 1px solid #E0D8CC; border-left: 4px solid #5B8DB8;
    }
    .book-item { border-left-color: #9B59B6; }
    
    .branch-card {
        background: #FFFFFF; border-radius: 20px; padding: 30px; margin: 15px 0;
        box-shadow: 0 6px 30px rgba(0,0,0,0.08); border: 2px solid #F0EBE3;
        cursor: pointer; transition: all 0.3s ease;
    }
    .branch-card:hover {
        transform: translateY(-3px);
        box-shadow: 0 10px 40px rgba(0,0,0,0.12);
    }
    .branch-science { border-left: 8px solid #9B59B6; }
    .branch-legal { border-left: 8px solid #2C3E50; }
    .branch-business { border-left: 8px solid #27AE60; }
    
    /* Кнопки истории */
    .history-btn {
        width: 100%;
        text-align: left;
        background: #FFFFFF !important;
        color: #4A403A !important;
        border: 1px solid #E0D8CC !important;
        border-radius: 10px !important;
        padding: 12px 15px !important;
        margin: 8px 0 !important;
        font-size: 14px !important;
        font-weight: 500 !important;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04) !important;
    }
    .history-btn:hover {
        background: #F8F5F0 !important;
        border-color: #D4A373 !important;
    }
    .history-btn-date {
        font-size: 11px !important;
        color: #A89F91 !important;
        font-weight: 400 !important;
    }
    
    /* Кнопка удаления */
    .delete-btn {
        background-color: #E74C3C !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 6px 12px !important;
        font-size: 12px !important;
        margin-top: 8px !important;
    }
    .delete-btn:hover {
        background-color: #C0392B !important;
    }
</style>
""", unsafe_allow_html=True)

st.set_page_config(page_title="Совет ИИ", page_icon="🤖", layout="wide")

# --- ФУНКЦИИ ДЛЯ ФИНАНСОВЫХ API ---
def get_cbr_data():
    try:
        url = "https://www.cbr-xml-daily.ru/daily_json.js"
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        data = response.json()
        result = {'date': data.get('Date', ''), 'key_rate': data.get('KeyRate', ''), 'currencies': {}}
        for currency_code in ['USD', 'EUR', 'CNY', 'GBP', 'JPY']:
            if currency_code in data['Valute']:
                curr = data['Valute'][currency_code]
                result['currencies'][currency_code] = {
                    'name': curr.get('Name', ''), 'value': curr.get('Value', 0),
                    'previous': curr.get('Previous', 0), 'change': curr.get('Value', 0) - curr.get('Previous', 0)
                }
        return result
    except Exception as e:
        return {'error': f'Ошибка ЦБ РФ: {str(e)}'}

def get_yahoo_finance_data(ticker, period='1mo'):
    try:
        stock = yf.Ticker(ticker)
        hist = stock.history(period=period)
        if hist.empty:
            return {'error': f'Нет данных для {ticker}'}
        info = stock.info
        return {
            'ticker': ticker, 'period': period,
            'current_price': hist['Close'].iloc[-1] if len(hist) > 0 else None,
            'price_change_percent': ((hist['Close'].iloc[-1] - hist['Close'].iloc[0]) / hist['Close'].iloc[0] * 100) if len(hist) > 1 else None,
            'high': hist['High'].max(), 'low': hist['Low'].min(),
            'company_name': info.get('shortName', ticker), 'currency': info.get('currency', 'USD'),
            'market_cap': info.get('marketCap', None), 'pe_ratio': info.get('trailingPE', None),
        }
    except Exception as e:
        return {'error': f'Ошибка Yahoo Finance: {str(e)}'}

# --- ФУНКЦИИ ДЛЯ ФИНАНСОВЫХ МОДЕЛЕЙ ---
def create_financial_model(initial_investment, annual_revenue, annual_costs, growth_rate, discount_rate, years=5):
    wb = openpyxl.Workbook()
    ws_inputs = wb.active
    ws_inputs.title = "Исходные данные"
    ws_inputs['A1'] = "ИСХОДНЫЕ ДАННЫЕ"
    ws_inputs['A1'].font = Font(size=14, bold=True, color="FFFFFF")
    ws_inputs['A1'].fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
    ws_inputs.merge_cells('A1:C1')
    
    params = [("Начальные инвестиции (руб):", initial_investment), ("Годовая выручка год 1 (руб):", annual_revenue),
              ("Годовые расходы (руб):", annual_costs), ("Темп роста (% в год):", growth_rate),
              ("Ставка дисконтирования (%):", discount_rate), ("Горизонт (лет):", years)]
    for i, (label, value) in enumerate(params, 3):
        ws_inputs[f'A{i}'] = label; ws_inputs[f'A{i}'].font = Font(bold=True); ws_inputs[f'B{i}'] = value
    
    ws_cf = wb.create_sheet("Денежные потоки")
    ws_cf['A1'] = "ПРОГНОЗ ДЕНЕЖНЫХ ПОТОКОВ"
    ws_cf['A1'].font = Font(size=14, bold=True, color="FFFFFF")
    ws_cf['A1'].fill = PatternFill(start_color="27AE60", end_color="27AE60", fill_type="solid")
    ws_cf.merge_cells('A1:G1')
    
    headers = ["Показатель", "Год 0", "Год 1", "Год 2", "Год 3", "Год 4", "Год 5"]
    for col, h in enumerate(headers, 1):
        c = ws_cf.cell(row=3, column=col); c.value = h; c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color="34495E", end_color="34495E", fill_type="solid")
    
    rows_data = [("Инвестиции", -initial_investment, None), ("Выручка", None, annual_revenue),
                 ("Расходы", None, annual_costs), ("EBITDA", None, None), ("Налог 20%", None, None),
                 ("Чистый поток", None, None), ("Коэф. дисконт.", None, None),
                 ("Дисконт. поток", None, None), ("Накопл. NPV", None, None)]
    
    for row_idx, (label, v0, v1) in enumerate(rows_data, 4):
        ws_cf.cell(row=row_idx, column=1).value = label
        ws_cf.cell(row=row_idx, column=1).font = Font(bold=True)
        if v0 is not None: ws_cf.cell(row=row_idx, column=2).value = v0
        if v1 is not None: ws_cf.cell(row=row_idx, column=3).value = v1
        
        if row_idx == 2:
            for y in range(4, 8):
                p = get_column_letter(y-1); ws_cf.cell(row=row_idx, column=y).value = f"={p}{row_idx}*(1+B5/100)"
        elif row_idx == 3:
            for y in range(4, 8):
                p = get_column_letter(y-1); ws_cf.cell(row=row_idx, column=y).value = f"={p}{row_idx}*(1+B4/100)"
        elif row_idx == 4:
            for y in range(2, 8):
                cl = get_column_letter(y); ws_cf.cell(row=row_idx, column=y).value = f"={cl}2-{cl}3"
        elif row_idx == 5:
            for y in range(2, 8):
                cl = get_column_letter(y); ws_cf.cell(row=row_idx, column=y).value = f"={cl}4*0.2"
        elif row_idx == 6:
            ws_cf.cell(row=row_idx, column=2).value = "=B2"
            for y in range(3, 8):
                cl = get_column_letter(y); ws_cf.cell(row=row_idx, column=y).value = f"={cl}4-{cl}5"
        elif row_idx == 7:
            ws_cf.cell(row=row_idx, column=2).value = 1
            for y in range(3, 8):
                ws_cf.cell(row=row_idx, column=y).value = f"=1/(1+B6/100)^{y-2}"
        elif row_idx == 8:
            for y in range(2, 8):
                cl = get_column_letter(y); ws_cf.cell(row=row_idx, column=y).value = f"={cl}6*{cl}7"
        elif row_idx == 9:
            ws_cf.cell(row=row_idx, column=2).value = "=B8"
            for y in range(3, 8):
                cl = get_column_letter(y); p = get_column_letter(y-1)
                ws_cf.cell(row=row_idx, column=y).value = f"={p}9+{cl}8"
    
    for r in range(4, 13):
        for c in range(2, 8):
            ws_cf.cell(row=r, column=c).number_format = '#,##0.00'
    
    ws_m = wb.create_sheet("Ключевые показатели")
    ws_m['A1'] = "КЛЮЧЕВЫЕ ПОКАЗАТЕЛИ"
    ws_m['A1'].font = Font(size=14, bold=True, color="FFFFFF")
    ws_m['A1'].fill = PatternFill(start_color="E74C3C", end_color="E74C3C", fill_type="solid")
    ws_m.merge_cells('A1:C1')
    
    metrics = [("NPV", f"='Денежные потоки'!F9", '#,##0.00'),
               ("IRR", f"=IRR('Денежные потоки'!B6:G6)", '0.00%'),
               ("ROI", f"=('Денежные потоки'!F6-B2)/ABS(B2)*100", '0.00'),
               ("Срок окупаемости", "5", '0.0'),
               ("Индекс рентабельности", f"='Денежные потоки'!F9/ABS(B2)", '0.00')]
    for i, (l, f, fmt) in enumerate(metrics, 3):
        ws_m[f'A{i}'] = l; ws_m[f'A{i}'].font = Font(bold=True)
        ws_m[f'B{i}'] = f; ws_m[f'B{i}'].number_format = fmt
    
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        wb.save(tmp.name); return tmp.name

# --- ФУНКЦИИ ДЛЯ НАУЧНОГО ПОИСКА ---
def search_openalex(query, max_results=15):
    try:
        url = f"https://api.openalex.org/works?search={query}&per-page={max_results}&sort=relevance_score:desc&filter=open_access.is_oa:true"
        response = requests.get(url, timeout=15); response.raise_for_status(); data = response.json()
        articles = []
        for work in data.get('results', []):
            title = work.get('title', 'Без названия')
            abstract = work.get('abstract_inverted_index'); abstract_text = ""
            if abstract:
                words = {}
                for word, positions in abstract.items():
                    for pos in positions: words[pos] = word
                abstract_text = ' '.join([words[i] for i in sorted(words.keys())])
            open_access = work.get('open_access', {}); pdf_url = open_access.get('oa_url', '')
            authors = [a.get('author', {}).get('display_name', '') for a in work.get('authorships', [])[:3]]
            if pdf_url:
                articles.append({'title': title, 'abstract': abstract_text[:500] if abstract_text else 'Нет аннотации',
                    'authors': ', '.join(authors) if authors else 'Неизвестны', 'year': work.get('publication_year', ''),
                    'pdf_url': pdf_url, 'source': 'OpenAlex', 'has_full_text': True})
        return articles
    except Exception as e:
        return [{'error': f'Ошибка OpenAlex: {str(e)}'}]

def search_cyberleninka(query, max_results=15):
    try:
        url = f"https://cyberleninka.ru/api/n/search?q={query}&per_page={max_results}"
        response = requests.get(url, timeout=15); response.raise_for_status(); data = response.json()
        articles = []
        for article in data.get('results', []):
            pdf_url = article.get('pdf', '')
            if pdf_url:
                articles.append({'title': article.get('title', ''), 'abstract': article.get('abstract', '')[:500],
                    'authors': article.get('authors', ''), 'year': article.get('year', ''),
                    'pdf_url': pdf_url, 'source': 'КиберЛенинка', 'has_full_text': True})
        return articles
    except Exception as e:
        return [{'error': f'Ошибка КиберЛенинки: {str(e)}'}]

def search_semantic_scholar(query, max_results=15):
    try:
        url = "https://api.semanticscholar.org/graph/v1/paper/search"
        params = {'query': query, 'limit': max_results, 'fields': 'title,abstract,authors,year,externalIds,citationCount'}
        response = requests.get(url, params=params, timeout=15); response.raise_for_status(); data = response.json()
        articles = []
        for paper in data.get('data', []):
            title = paper.get('title', ''); abstract = paper.get('abstract', '')[:500]
            authors = [a.get('name', '') for a in paper.get('authors', [])[:3]]
            year = paper.get('year', ''); citations = paper.get('citationCount', 0)
            ext = paper.get('externalIds', {}); doi = ext.get('DOI', '')
            pdf_url = f"https://doi.org/{doi}" if doi else ''
            articles.append({'title': title, 'abstract': abstract if abstract else 'Нет аннотации',
                'authors': ', '.join(authors) if authors else 'Неизвестны', 'year': year,
                'pdf_url': pdf_url, 'source': 'Semantic Scholar', 'citations': citations, 'has_full_text': bool(doi)})
        articles.sort(key=lambda x: x.get('citations', 0), reverse=True)
        return articles
    except Exception as e:
        return [{'error': f'Ошибка Semantic Scholar: {str(e)}'}]

def download_and_extract_pdf(pdf_url, max_chars=15000):
    try:
        response = requests.get(pdf_url, timeout=30); response.raise_for_status()
        pdf_file = BytesIO(response.content); pdf_reader = PyPDF2.PdfReader(pdf_file)
        text = "".join([page.extract_text() + "\n" for page in pdf_reader.pages])
        return text[:max_chars] if len(text) > max_chars else text
    except Exception as e:
        return f"Ошибка PDF: {str(e)}"

# --- ФУНКЦИИ ДЛЯ EXCEL ---
def extract_excel_data(file, max_rows=MAX_EXCEL_ROWS):
    try:
        excel_file = pd.ExcelFile(file); all_sheets = {}
        for sheet in excel_file.sheet_names:
            df = pd.read_excel(excel_file, sheet_name=sheet, nrows=max_rows)
            text = f"=== ЛИСТ: {sheet} ===\nСтрок: {len(df)}, Столбцов: {len(df.columns)}\n"
            text += f"Столбцы: {', '.join(df.columns.tolist())}\n\n"
            text += "ПРЕВЬЮ:\n" + df.head().to_string(index=False) + "\n\n"
            numeric = df.select_dtypes(include=['number']).columns
            if len(numeric) > 0:
                text += "СТАТИСТИКА:\n"
                for col in numeric:
                    text += f"{col}: мин={df[col].min():.2f}, макс={df[col].max():.2f}, среднее={df[col].mean():.2f}\n"
            all_sheets[sheet] = text
        combined = "\n\n".join(all_sheets.values())
        return (combined[:MAX_FILE_SIZE] if len(combined) > MAX_FILE_SIZE else combined), len(excel_file.sheet_names)
    except Exception as e:
        return f"Ошибка Excel: {str(e)}", 0

def analyze_excel_structure(file):
    try:
        df = pd.read_excel(file, nrows=10)
        analysis = {'has_financial_data': False, 'has_dates': False, 'numeric_columns': [], 'text_columns': []}
        for col in df.columns:
            if pd.api.types.is_numeric_dtype(df[col]):
                analysis['numeric_columns'].append(col)
                if any(k in col.lower() for k in ['выруч','прибыл','доход','расход','revenue','profit','income','expense']):
                    analysis['has_financial_data'] = True
            elif pd.api.types.is_datetime64_any_dtype(df[col]):
                analysis['has_dates'] = True
            else:
                analysis['text_columns'].append(col)
        return analysis
    except:
        return {'has_financial_data': False, 'has_dates': False}

# --- ФУНКЦИИ ДЛЯ БИБЛИОТЕКИ ---
def scan_library_folder(folder_path):
    books = []; folder = Path(folder_path)
    if not folder.exists(): return {'error': f'Папка не найдена: {folder_path}'}
    for fp in folder.iterdir():
        if fp.is_file() and fp.suffix.lower() in ['.pdf', '.docx', '.txt']:
            try:
                metadata = {}
                if fp.suffix.lower() == '.pdf':
                    with open(fp, 'rb') as f:
                        pr = PyPDF2.PdfReader(f); info = pr.metadata
                        if info:
                            metadata['title'] = info.get('/Title', ''); metadata['author'] = info.get('/Author', '')
                            metadata['year'] = info.get('/CreationDate', '')[:4] if info.get('/CreationDate') else ''
                        metadata['pages'] = len(pr.pages)
                elif fp.suffix.lower() == '.docx':
                    doc = docx.Document(fp); metadata['pages'] = len(doc.paragraphs) // 50
                books.append({'path': str(fp), 'filename': fp.name, 'title': metadata.get('title', fp.stem),
                    'author': metadata.get('author', 'Неизвестен'), 'year': metadata.get('year', ''),
                    'size_pages': metadata.get('pages', 0), 'extension': fp.suffix.lower()})
            except Exception as e:
                books.append({'path': str(fp), 'filename': fp.name, 'title': fp.stem, 'author': 'Неизвестен',
                    'year': '', 'size_pages': 0, 'extension': fp.suffix.lower(), 'error': str(e)})
    return books

def search_relevant_books(query, books, top_n=3):
    if not books: return []
    qw = set(query.lower().split()); scored = []
    for book in books:
        score = len(qw & set(book['title'].lower().split())) * 3
        score += len(qw & set(book['author'].lower().split())) * 2
        score += len(qw & set(book['filename'].lower().split()))
        if score > 0: scored.append((score, book))
    scored.sort(reverse=True, key=lambda x: x[0])
    return [b for s, b in scored[:top_n]]

def load_book_content(file_path, max_pages=MAX_BOOK_PAGES):
    fp = Path(file_path)
    if not fp.exists(): return f"Файл не найден: {fp}"
    try:
        if fp.suffix.lower() == '.pdf':
            with open(fp, 'rb') as f:
                pr = PyPDF2.PdfReader(f); return "".join([p.extract_text() + "\n" for p in pr.pages[:max_pages]])
        elif fp.suffix.lower() == '.docx':
            doc = docx.Document(fp); return "\n".join([p.text for p in doc.paragraphs[:max_pages*3]])
        elif fp.suffix.lower() == '.txt':
            with open(fp, 'r', encoding='utf-8') as f: return "".join(f.readlines()[:max_pages*50])
        return "Неподдерживаемый формат"
    except Exception as e:
        return f"Ошибка: {str(e)}"

# --- ОСНОВНЫЕ ФУНКЦИИ ---
def load_projects():
    if os.path.exists(PROJECTS_FILE):
        with open(PROJECTS_FILE, "r", encoding="utf-8") as f: return json.load(f)
    return []

def save_projects(projects):
    with open(PROJECTS_FILE, "w", encoding="utf-8") as f: json.dump(projects, f, ensure_ascii=False, indent=2)

def create_project(name, description=""):
    projects = load_projects()
    new_project = {"id": str(uuid.uuid4())[:8], "name": name, "description": description, "created": datetime.now().strftime("%d.%m.%Y %H:%M")}
    projects.append(new_project); save_projects(projects); return new_project

def delete_project(project_id):
    projects = [p for p in load_projects() if p["id"] != project_id]; save_projects(projects)
    history = [h for h in load_history() if h.get("project_id") != project_id]; save_history(history)

def load_history():
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, "r", encoding="utf-8") as f: return json.load(f)
    return []

def save_history(history):
    with open(HISTORY_FILE, "w", encoding="utf-8") as f: json.dump(history, f, ensure_ascii=False, indent=2)

def save_to_history(entry):
    history = load_history(); history.append(entry); save_history(history)

def delete_history_entry(entry_id):
    history = load_history()
    history = [h for h in history if h.get("id") != entry_id]
    save_history(history)

def get_history_entry(entry_id):
    """Получает конкретную запись истории по ID"""
    history = load_history()
    for h in history:
        if h.get("id") == entry_id:
            return h
    return None

def load_key():
    try: return st.secrets["API_KEY"]
    except: pass
    try:
        with open(KEY_FILE, "r", encoding="utf-8") as f: return f.read().strip()
    except FileNotFoundError: return None

def extract_text(file):
    fn = file.name.lower()
    if fn.endswith('.txt'): return file.read().decode('utf-8')
    elif fn.endswith('.pdf'):
        pr = PyPDF2.PdfReader(file); return "".join([p.extract_text() + "\n" for p in pr.pages])
    elif fn.endswith('.docx'):
        doc = docx.Document(file); return "".join([p.text + "\n" for p in doc.paragraphs])
    else: raise Exception("Неподдерживаемый формат.")

def ask(model, prompt):
    key = load_key()
    if not key: raise Exception("Не найден key.txt или секрет API_KEY")
    client = OpenAI(api_key=key, base_url=BASE_URL)
    return client.chat.completions.create(model=model, messages=[{"role": "user", "content": prompt}], temperature=0.3).choices[0].message.content

def analyze_image(file):
    key = load_key()
    if not key: raise Exception("Не найден key.txt")
    image_bytes = file.read(); base64_image = base64.b64encode(image_bytes).decode('utf-8')
    fn = file.name.lower(); mime = "image/jpeg" if fn.endswith(('.jpg', '.jpeg')) else "image/png"
    client = OpenAI(api_key=key, base_url=BASE_URL)
    response = client.chat.completions.create(model=MODEL_CHATGPT, messages=[{
        "role": "user", "content": [
            {"type": "text", "text": "Распознай весь текст на изображении. Опиши таблицы, схемы, подписи. Верни структурированный текст."},
            {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{base64_image}"}}
        ]}], max_tokens=2000)
    return response.choices[0].message.content

# --- ПРОМПТЫ ДЛЯ РАЗНЫХ ВЕТВЕЙ ---
def get_branch_prompts(branch, task, full_context, draft=None, critique=None):
    if branch == 'science':
        chatgpt_prompt = f"""Ты — научный ассистент и исследователь. Твоя задача: подготовить ЧЕРНОВИК НАУЧНОЙ РАБОТЫ.

{full_context}

## ЗАДАЧА:
{task}

## СТРУКТУРА НАУЧНОЙ РАБОТЫ:
1. Введение (актуальность, цель, задачи)
2. Обзор литературы / Теоретическая часть
3. Методология / Основная часть
4. Результаты и анализ
5. Выводы и рекомендации
6. СПИСОК ЛИТЕРАТУРЫ

## ТРЕБОВАНИЯ:
- Строго опирайся на источники
- Цитируй: книги как [К1], статьи как [1], данные как [Данные]
- Научный стиль изложения
- В конце обязательно добавь СПИСОК ЛИТЕРАТУРЫ

Напиши полный черновик научной работы."""

        claude_prompt = f"""Ты — строгий научный рецензент (рецензент научного журнала). Проверь черновик научной работы.

{full_context}

## ЗАДАЧА:
{task}

## ЧЕРНОВИК:
{draft}

## ПРОВЕРЬ:
1. ✅ Все ли утверждения подкреплены источниками?
2. ✅ Корректны ли цитаты и ссылки?
3. ✅ Есть ли список литературы и соответствует ли он цитированиям?
4. ✅ Соблюдена ли структура научной работы?
5. ⚠️ Какие утверждения требуют дополнительных доказательств?
6. ❌ Что упущено из источников?

Будь МАКСИМАЛЬНО строгим как рецензент научного журнала."""

        qwen_prompt = f"""Ты — главный редактор научного журнала. Создай ФИНАЛЬНУЮ ВЕРСИЮ научной работы.

{full_context}

## ЗАДАЧА:
{task}

## ЧЕРНОВИК:
{draft}

## КРИТИКА РЕЦЕНЗЕНТА:
{critique}

## ТРЕБОВАНИЯ:
1. Исправь ВСЕ замечания рецензента
2. Добавь недостающие цитаты [К1], [1], [Данные]
3. Соблюдай структуру: Введение → Обзор → Методология → Результаты → Выводы → Список литературы
4. Научный стиль, академическая точность
5. В конце — полный СПИСОК ЛИТЕРАТУРЫ

Создай финальную версию научной работы, готовую к публикации."""

    elif branch == 'legal':
        chatgpt_prompt = f"""Ты — практикующий юрист с 15-летним стажем. Твоя задача:
