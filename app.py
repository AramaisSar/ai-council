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
from openpyxl.chart import BarChart, LineChart, Reference
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
    
    /* Кнопка удаления беседы */
    .delete-btn {
        background-color: #E74C3C !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 5px 10px !important;
        font-size: 12px !important;
        margin-top: 10px !important;
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
    """Удаляет конкретную запись из истории по ID"""
    history = load_history()
    history = [h for h in history if h.get("id") != entry_id]
    save_history(history)

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
5. ️ Какие утверждения требуют дополнительных доказательств?
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
        chatgpt_prompt = f"""Ты — практикующий юрист с 15-летним стажем. Твоя задача: подготовить ПРОЕКТ ЮРИДИЧЕСКОГО ДОКУМЕНТА.

{full_context}

## ЗАДАЧА:
{task}

## ТРЕБОВАНИЯ К ЮРИДИЧЕСКОМУ ДОКУМЕНТУ:
- Строгий юридический стиль (не научный!)
- Ссылки на конкретные статьи законов (ГК РФ, УК РФ, АПК РФ и т.д.)
- Чёткая структура документа (преамбула, разделы, пункты, подписи)
- Практическая применимость — документ должен быть готов к использованию
- Если это договор — все существенные условия
- Если это иск/жалоба — ссылки на нормы права и судебную практику
- НЕ пиши научную статью — пиши именно документ!

## ФОРМАТ:
Начни с названия документа, затем преамбула/вводная часть, затем основное содержание по разделам, в конце — подписи/реквизиты.

Подготовь полный проект документа."""

        claude_prompt = f"""Ты — опытный юрист-рецензент, старший партнёр юридической фирмы. Проверь проект юридического документа.

{full_context}

## ЗАДАЧА:
{task}

## ПРОЕКТ ДОКУМЕНТА:
{draft}

## ПРОВЕРЬ:
1. ✅ Все ли существенные условия включены?
2. ✅ Корректны ли ссылки на нормы права (статьи законов)?
3. ✅ Нет ли юридических рисков и двусмысленностей?
4. ✅ Соответствует ли документ актуальному законодательству РФ?
5. ️ Какие пункты могут быть оспорены в суде?
6. ❌ Что упущено и может создать проблемы?
7. ️ Если есть мнение внешнего ИИ (Cognevia/ПравоТех) — учтено ли оно?

Будь МАКСИМАЛЬНО строгим. Найди все юридические риски."""

        qwen_prompt = f"""Ты — старший партнёр юридической фирмы. Создай ФИНАЛЬНУЮ ВЕРСИЮ юридического документа.

{full_context}

## ЗАДАЧА:
{task}

## ПРОЕКТ:
{draft}

## ЗАМЕЧАНИЯ РЕЦЕНЗЕНТА:
{critique}

## ТРЕБОВАНИЯ:
1. Исправь ВСЕ юридические риски, указанные рецензентом
2. Добавь недостающие ссылки на нормы права
3. Устрани все двусмысленности
4. Соблюдай строгий юридический стиль
5. Документ должен быть готов к подписанию/подаче в суд
6. НЕ пиши научную статью — это должен быть именно юридический документ!

Создай финальную версию документа, готовую к использованию."""

    else:  # business
        chatgpt_prompt = f"""Ты — бизнес-аналитик и финансовый консультант уровня McKinsey/BCG. Твоя задача: подготовить БИЗНЕС-ОТЧЁТ / АНАЛИТИЧЕСКИЙ ДОКУМЕНТ.

{full_context}

## ЗАДАЧА:
{task}

## СТРУКТУРА БИЗНЕС-ОТЧЁТА:
1. Резюме (Executive Summary) — ключевые выводы на 1 странице
2. Анализ ситуации / рынка
3. Финансовый анализ (если есть данные)
4. Оценка рисков
5. Рекомендации и план действий
6. Приложения (расчёты, таблицы)

## ТРЕБОВАНИЯ:
- Деловой стиль, конкретика, цифры
- Используй данные из финансовых API (курсы валют, ключевая ставка, котировки)
- Если есть Excel — проведи полноценный финансовый анализ
- Цитируй источники: [Данные], [Финансы], [1]
- Практические рекомендации, а не теория

Подготовь полный бизнес-отчёт."""

        claude_prompt = f"""Ты — финансовый аудитор и бизнес-консультант. Проверь бизнес-отчёт.

{full_context}

## ЗАДАЧА:
{task}

## ОТЧЁТ:
{draft}

## ПРОВЕРЬ:
1. ✅ Все ли расчёты корректны?
2. ✅ Использованы ли актуальные финансовые данные (курсы, ставки)?
3. ✅ Реалистичны ли прогнозы и допущения?
4. ✅ Учтены ли все ключевые риски?
5. ⚠️ Какие выводы недостаточно обоснованы?
6.  Что упущено в анализе?

Будь строгим как финансовый аудитор."""

        qwen_prompt = f"""Ты — инвестиционный консультант и стратег. Создай ФИНАЛЬНУЮ ВЕРСИЮ бизнес-отчёта.

{full_context}

## ЗАДАЧА:
{task}

## ОТЧЁТ:
{draft}

## ЗАМЕЧАНИЯ АУДИТОРА:
{critique}

## ТРЕБОВАНИЯ:
1. Исправь все замечания аудитора
2. Добавь недостающие финансовые расчёты
3. Используй актуальные данные (курсы, ставки, котировки)
4. Деловой стиль, конкретные цифры и рекомендации
5. Структура: Резюме → Анализ → Финансы → Риски → Рекомендации

Создай финальный бизнес-отчёт, готовый для презентации руководству или инвесторам."""

    return chatgpt_prompt, claude_prompt, qwen_prompt

# --- ИНТЕРФЕЙС ---
st.sidebar.header("📁 Проекты")
projects = load_projects()

with st.sidebar.expander("➕ Новый проект"):
    new_project_name = st.text_input("Название проекта:", placeholder="Например: Юридическая практика")
    new_project_desc = st.text_area("Описание:", height=80)
    if st.button("Создать проект"):
        if new_project_name.strip():
            create_project(new_project_name.strip(), new_project_desc.strip())
            st.success(f"Проект '{new_project_name}' создан!"); st.rerun()
        else: st.warning("Введите название")

if projects:
    project_names = [f"{p['name']} ({p['created']})" for p in projects]
    selected_name = st.sidebar.selectbox("Выбери проект:", project_names)
    selected_project = projects[project_names.index(selected_name)]
    if st.sidebar.button("🗑 Удалить проект"):
        if st.sidebar.checkbox("Да, удалить"):
            delete_project(selected_project["id"]); st.sidebar.success("Удалён"); st.rerun()
else:
    st.sidebar.info("Нет проектов. Создай первый!"); selected_project = None

st.sidebar.markdown("---")

if selected_project:
    st.sidebar.header(f" История: {selected_project['name']}")
    history = load_history()
    project_history = [h for h in history if h.get("project_id") == selected_project["id"]]
    
    if project_history:
        st.sidebar.markdown("---")
        for i, entry in enumerate(reversed(project_history)):
            short = entry["task"][:40] + "..." if len(entry["task"]) > 40 else entry["task"]
            with st.sidebar.expander(f"📝 {short}\n🕒 {entry['date']}"):
                st.write("**Задача:**"); st.write(entry["task"])
                if entry.get("file"): st.write(f"**Файл:** {entry['file']}")
                st.write("**Результат:**"); st.success(entry["final"])
                
                # Кнопка удаления беседы
                if st.button("🗑 Удалить эту беседу", key=f"delete_{entry.get('id', i)}", use_container_width=True):
                    delete_history_entry(entry.get("id"))
                    st.sidebar.success("✅ Беседа удалена!")
                    st.rerun()
    else:
        st.sidebar.info("Пока нет задач")
else:
    st.sidebar.header("📂 История"); st.sidebar.info("Выбери проект")

st.sidebar.markdown("---")
if load_key(): st.sidebar.success("🔑 Ключ активен")
else: st.sidebar.error("❌ Ключ не найден")

# --- КНОПКА ВЫХОДА ИЗ ВЕТВИ ---
if 'branch' in st.session_state and st.session_state['branch']:
    if st.sidebar.button("🏠 Вернуться к выбору ветви"):
        st.session_state['branch'] = None
        st.rerun()

# === ГЛАВНАЯ ЛОГИКА: ВЫБОР ВЕТВИ ИЛИ РАБОТА ===
if 'branch' not in st.session_state:
    st.session_state['branch'] = None

if not st.session_state['branch']:
    st.markdown(f"""<div class="ai-card"><h3 style="margin: 0; color: #8B5E3C;">📁 Проект: {selected_project['name'] if selected_project else 'Не выбран'}</h3>
    <p style="margin: 10px 0 0 0; color: #A89F91;">Выбери направление работы. В каждой ветви "аспиранты" заточены под конкретный результат.</p></div>""", unsafe_allow_html=True)
    
    st.markdown("---")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("""
        <div class="branch-card branch-science">
            <h3 style="color: #9B59B6; margin: 0 0 10px 0;">🎓 Научная ветвь</h3>
            <p style="color: #666; margin: 0 0 15px 0;">Научные статьи, диссертации, обзоры литературы</p>
            <p style="color: #888; font-size: 0.9em; margin: 0;">
            ✅ Научный поиск (КиберЛенинка, OpenAlex, Semantic Scholar)<br>
            ✅ Моя научная библиотека<br>
            ✅ Структура: Введение → Обзор → Методология → Выводы → Список литературы
            </p>
        </div>
        """, unsafe_allow_html=True)
        if st.button(" Научная ветвь", key="btn_science", use_container_width=True):
            st.session_state['branch'] = 'science'; st.rerun()
    
    with col2:
        st.markdown("""
        <div class="branch-card branch-legal">
            <h3 style="color: #2C3E50; margin: 0 0 10px 0;">⚖️ Юридическая ветвь</h3>
            <p style="color: #666; margin: 0 0 15px 0;">Договоры, иски, претензии, жалобы, письма</p>
            <p style="color: #888; font-size: 0.9em; margin: 0;">
            ✅ Мнение внешнего ИИ (Cognevia, ПравоТех)<br>
            ✅ Строгий юридический стиль<br>
            ✅ Ссылки на нормы права и судебную практику
            </p>
        </div>
        """, unsafe_allow_html=True)
        if st.button("⚖️ Юридическая ветвь", key="btn_legal", use_container_width=True):
            st.session_state['branch'] = 'legal'; st.rerun()
    
    with col3:
        st.markdown("""
        <div class="branch-card branch-business">
            <h3 style="color: #27AE60; margin: 0 0 10px 0;"> Бизнес-финансовая ветвь</h3>
            <p style="color: #666; margin: 0 0 15px 0;">Бизнес-планы, финансовые отчёты, анализ</p>
            <p style="color: #888; font-size: 0.9em; margin: 0;">
            ✅ Генератор финансовых моделей (NPV, IRR, ROI)<br>
            ✅ Актуальные данные (ЦБ РФ, Yahoo Finance)<br>
            ✅ Деловой стиль с цифрами и рекомендациями
            </p>
        </div>
        """, unsafe_allow_html=True)
        if st.button("💼 Бизнес-финансовая ветвь", key="btn_business", use_container_width=True):
            st.session_state['branch'] = 'business'; st.rerun()
    
    st.stop()

# === РАБОТА ВНУТРИ ВЕТВИ ===
branch = st.session_state['branch']
branch_names = {'science': ' Научная ветвь', 'legal': '⚖️ Юридическая ветвь', 'business': '💼 Бизнес-финансовая ветвь'}
branch_colors = {'science': '#9B59B6', 'legal': '#2C3E50', 'business': '#27AE60'}

st.markdown(f"""<div class="ai-card" style="border-left: 8px solid {branch_colors[branch]};">
    <h3 style="margin: 0; color: {branch_colors[branch]};">{branch_names[branch]}</h3>
    <p style="margin: 10px 0 0 0; color: #A89F91;">Проект: {selected_project['name'] if selected_project else 'Не выбран'}</p>
</div>""", unsafe_allow_html=True)

st.markdown("---")

# === ИНСТРУМЕНТЫ ВЕТВИ ===
if branch == 'science':
    st.markdown("""<div class="ai-card card-library"><h3 style="margin: 0 0 10px 0; color: #9B59B6;">📚 Моя научная библиотека</h3>
    <p style="margin: 0; color: #A89F91;">Укажи путь к папке с учебниками и монографиями.</p></div>""", unsafe_allow_html=True)
    
    library_path = st.text_input("Путь к папке:", placeholder="C:\\Users\\123\\Documents\\Научная_библиотека")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("📂 Сканировать библиотеку", use_container_width=True):
            if library_path.strip():
                with st.spinner("Сканирую..."):
                    books = scan_library_folder(library_path)
                    if isinstance(books, dict) and 'error' in books: st.error(books['error'])
                    else: st.success(f"✅ Найдено {len(books)} книг!"); st.session_state['library_books'] = books
            else: st.warning("Введи путь")
    with col2:
        use_library = st.checkbox("📖 Использовать библиотеку", value=False)
    
    if 'library_books' in st.session_state:
        st.markdown(f"### 📚 Найдено: {len(st.session_state['library_books'])} книг")
        for i, book in enumerate(st.session_state['library_books'][:5], 1):
            st.markdown(f"**{i}.** {book['title']} — {book['author']} ({book['year']})")
    
    st.markdown("---")
    
    st.markdown("""<div class="ai-card card-search"><h3 style="margin: 0 0 10px 0; color: #5B8DB8;">🔬 Научный поиск</h3>
    <p style="margin: 0; color: #A89F91;">Поиск по научным базам данных.</p></div>""", unsafe_allow_html=True)
    
    search_query = st.text_input("Тема поиска:", placeholder="artificial intelligence in civil law")
    col1, col2, col3 = st.columns(3)
    with col1: use_cyberleninka = st.checkbox("🇷 КиберЛенинка", value=True)
    with col2: use_openalex = st.checkbox("🌍 OpenAlex", value=True)
    with col3: use_semantic = st.checkbox(" Semantic Scholar", value=True)
    deep_analysis = st.checkbox("🎓 Глубокое исследование (скачать полные тексты)", value=True)
    
    if st.button("🔍 Найти статьи", use_container_width=True):
        if search_query.strip():
            all_articles = []
            if use_cyberleninka:
                with st.spinner("КиберЛенинка..."): all_articles.extend(search_cyberleninka(search_query))
            if use_openalex:
                with st.spinner("OpenAlex..."): all_articles.extend(search_openalex(search_query))
            if use_semantic:
                with st.spinner("Semantic Scholar..."): all_articles.extend(search_semantic_scholar(search_query))
            
            full_text = [a for a in all_articles if a.get('has_full_text', False)]
            if deep_analysis and full_text:
                with st.spinner(f"Скачиваю {min(len(full_text), MAX_ARTICLES)} статей..."):
                    articles_with_text = []
                    for article in full_text[:MAX_ARTICLES]:
                        text = download_and_extract_pdf(article['pdf_url'])
                        article['full_text'] = text; articles_with_text.append(article)
                st.session_state['deep_analysis_articles'] = articles_with_text
                st.success(f"✅ Скачано {len(articles_with_text)} статей!")
            else:
                st.session_state['search_results'] = all_articles
                st.success(f"Найдено: {len(all_articles)}")
        else: st.warning("Введи тему")

elif branch == 'legal':
    st.markdown("""<div class="ai-card card-external"><h3 style="margin: 0 0 10px 0; color: #6B9080;">️ Мнение внешнего юридического ИИ</h3>
    <p style="margin: 0; color: #A89F91;">Вставь ответ от Cognevia, ПравоТех, Doczilla (необязательно).</p></div>""", unsafe_allow_html=True)
    
    external_ai_opinion = st.text_area("Ответ от внешнего ИИ:", height=120, 
        placeholder="Скопируй сюда анализ от Cognevia, ПравоТех или другого юридического ИИ...")
    
    st.markdown("---")
    
    st.markdown("""<div class="ai-card"><h3 style="margin: 0 0 10px 0; color: #2C3E50;"> Тип юридического документа</h3>
    <p style="margin: 0; color: #A89F91;">Выбери тип — это поможет ИИ понять структуру документа.</p></div>""", unsafe_allow_html=True)
    
    doc_type = st.selectbox("Тип документа:", [
        "Договор (купли-продажи, аренды, оказания услуг и т.д.)",
        "Исковое заявление",
        "Претензия",
        "Жалоба (в суд, прокуратуру, Роспотребнадзор)",
        "Юридическое письмо",
        "Правовое заключение (memo)",
        "Другое (укажу в задаче)"
    ], index=0)

elif branch == 'business':
    st.markdown("""<div class="ai-card card-finance"><h3 style="margin: 0 0 10px 0; color: #E74C3C;">💼 Генератор финансовых моделей</h3>
    <p style="margin: 0; color: #A89F91;">Создай модель с NPV, IRR, ROI за 1 минуту.</p></div>""", unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        initial_investment = st.number_input("Инвестиции (руб):", min_value=0, value=1000000, step=100000)
        annual_revenue = st.number_input("Выручка год 1 (руб):", min_value=0, value=500000, step=50000)
        annual_costs = st.number_input("Расходы (руб):", min_value=0, value=300000, step=50000)
    with col2:
        growth_rate = st.number_input("Рост (%/год):", min_value=0, max_value=100, value=10, step=5)
        discount_rate = st.number_input("Ставка дисконтирования (%):", min_value=0, max_value=100, value=12, step=1)
        years = st.number_input("Горизонт (лет):", min_value=1, max_value=10, value=5, step=1)
    
    if st.button("📊 Создать финансовую модель", use_container_width=True):
        with st.spinner("Генерирую..."):
            try:
                excel_path = create_financial_model(initial_investment, annual_revenue, annual_costs, growth_rate, discount_rate, years)
                with open(excel_path, "rb") as f:
                    st.download_button(label="📥 Скачать модель Excel", data=f,
                        file_name=f"Финмодель_{datetime.now().strftime('%d%m%Y_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                st.success("✅ Модель создана! Все формулы автоматические.")
                os.remove(excel_path)
            except Exception as e: st.error(f"Ошибка: {str(e)}")
    
    st.markdown("---")
    
    st.markdown("""<div class="ai-card card-api"><h3 style="margin: 0 0 10px 0; color: #3498DB;">📈 Актуальные финансовые данные</h3>
    <p style="margin: 0; color: #A89F91;">ЦБ РФ и Yahoo Finance.</p></div>""", unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        use_cbr = st.checkbox("🇷 ЦБ РФ (курсы, ставка)", value=True)
    with col2:
        use_yahoo = st.checkbox("🌍 Yahoo Finance", value=False)
        yahoo_tickers = st.text_input("Тикеры:", placeholder="AAPL, MSFT, ^GSPC", value="AAPL")
    
    if st.button("📊 Получить данные", use_container_width=True):
        financial_parts = []
        if use_cbr:
            with st.spinner("ЦБ РФ..."):
                cbr = get_cbr_data()
                if 'error' not in cbr:
                    text = f"## 🇷🇺 ЦБ РФ ({cbr.get('date', '')})\n**Ключевая ставка:** {cbr.get('key_rate', 'N/A')}%\n**Курсы:**\n"
                    for code, d in cbr.get('currencies', {}).items():
                        s = "" if d['change'] > 0 else "📉" if d['change'] < 0 else "➡️"
                        text += f"- **{code}**: {d['value']:.2f} руб {s} ({d['change']:+.2f})\n"
                    financial_parts.append(text); st.success(f"✅ ЦБ РФ: ставка {cbr.get('key_rate', 'N/A')}%")
                else: st.error(cbr.get('error'))
        if use_yahoo and yahoo_tickers.strip():
            with st.spinner("Yahoo Finance..."):
                tickers = [t.strip() for t in yahoo_tickers.split(',')]
                yahoo_data = []
                for ticker in tickers:
                    data = get_yahoo_finance_data(ticker, '1mo'); yahoo_data.append(data)
                    if 'error' not in data:
                        st.success(f"✅ {ticker}: {data.get('current_price', 0):.2f} {data.get('currency', '')}")
                    else: st.error(f"❌ {ticker}: {data.get('error')}")
                st.session_state['yahoo_data'] = yahoo_data
        
        if financial_parts or 'yahoo_data' in st.session_state:
            full_fin = "# 📊 АКТУАЛЬНЫЕ ФИНАНСОВЫЕ ДАННЫЕ\n\n"
            for p in financial_parts: full_fin += p + "\n"
            if 'yahoo_data' in st.session_state:
                full_fin += "## 🌍 Yahoo Finance\n"
                for d in st.session_state['yahoo_data']:
                    if 'error' not in d:
                        full_fin += f"### {d.get('ticker', '')} ({d.get('company_name', '')})\n"
                        full_fin += f"- Цена: {d.get('current_price', 0):.2f} {d.get('currency', '')}\n"
                        if d.get('price_change_percent'): full_fin += f"- Изменение: {d['price_change_percent']:+.2f}%\n"
                        full_fin += "\n"
            st.session_state['financial_api_data'] = full_fin
            st.success("✅ Данные добавлены в контекст!")

# === ОБЩИЕ ЭЛЕМЕНТЫ ДЛЯ ВСЕХ ВЕТВЕЙ ===
st.markdown("---")

st.markdown("""<div class="ai-card card-excel"><h3 style="margin: 0 0 10px 0; color: #27AE60;">📎 Загрузи материалы</h3>
<p style="margin: 0; color: #A89F91;">PDF, DOCX, TXT, Excel, фото (JPG, PNG)</p></div>""", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Выбери файл:", type=['pdf', 'docx', 'txt', 'jpg', 'jpeg', 'png', 'xlsx', 'xls'])
file_text = ""; file_name = ""; file_type = ""
excel_analysis = None

if uploaded_file is not None:
    try:
        file_name = uploaded_file.name; fn = file_name.lower()
        if fn.endswith(('.jpg', '.jpeg', '.png')):
            with st.spinner("Анализирую изображение..."): file_text = analyze_image(uploaded_file)
            st.success(f"✅ Изображение: {file_name}"); file_type = "image"
        elif fn.endswith(('.xlsx', '.xls')):
            with st.spinner("Анализирую Excel..."):
                file_text, num = extract_excel_data(uploaded_file)
                excel_analysis = analyze_excel_structure(uploaded_file)
                st.success(f"✅ Excel: {file_name} ({num} листов)")
                if excel_analysis and excel_analysis['has_financial_data']: st.success("💰 Финансовые данные обнаружены")
            file_type = "excel"
        else:
            file_text = extract_text(uploaded_file)
            if len(file_text) > MAX_FILE_SIZE: file_text = file_text[:MAX_FILE_SIZE]
            st.success(f"✅ Файл: {file_name}"); file_type = "document"
    except Exception as e: st.error(f"❌ Ошибка: {e}")

st.markdown("---")

task = st.text_area("Опиши задачу:", height=150, 
    placeholder={
        'science': "Например: 'Напиши научную статью о влиянии ИИ на гражданское право'",
        'legal': "Например: 'Составь договор купли-продажи автомобиля между физическими лицами'",
        'business': "Например: 'Проанализируй инвестиционную привлекательность проекта с учётом текущей ключевой ставки'"
    }.get(branch, ""))

col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    if st.button("🚀 Запустить обсуждение", use_container_width=True):
        if not selected_project: st.warning("⚠️ Выбери проект")
        elif not task.strip(): st.warning("️ Введи задачу")
        else:
            try:
                context_parts = []
                
                if 'financial_api_data' in st.session_state and st.session_state['financial_api_data']:
                    context_parts.append(st.session_state['financial_api_data'])
                
                if branch == 'science' and use_library and 'library_books' in st.session_state:
                    with st.spinner("Ищу книги..."):
                        relevant = search_relevant_books(task, st.session_state['library_books'], top_n=3)
                        if relevant:
                            bt = "# 📚 БИБЛИОТЕКА\nИНСТРУКЦИЯ: Цитируй как [К1], [К2]\n\n"
                            for i, book in enumerate(relevant, 1):
                                with st.status(f"Загружаю: {book['title'][:50]}..."):
                                    bt += f"## [К{i}] {book['title']}\n**Автор:** {book['author']} | **Год:** {book['year']}\n\n{load_book_content(book['path'])[:10000]}\n\n---\n\n"
                            context_parts.append(bt); st.success(f"✅ {len(relevant)} книг")
                
                if branch == 'legal' and 'external_ai_opinion' in dir() and external_ai_opinion.strip():
                    context_parts.append(f"# ⚖️ МНЕНИЕ ВНЕШНЕГО ИИ\n\n{external_ai_opinion}\n\n")
                
                if file_type == "excel" and file_text:
                    ec = "# 📊 ДАННЫЕ EXCEL\n\n"
                    if excel_analysis and excel_analysis['has_financial_data']:
                        ec += "️ Финансовые данные. Выяви тренды, рассчитай метрики, найди аномалии.\n\n"
                    ec += file_text + "\n\n"; context_parts.append(ec)
                
                if 'deep_analysis_articles' in st.session_state and st.session_state['deep_analysis_articles']:
                    at = "#  СТАТЬИ (ПОЛНЫЕ ТЕКСТЫ)\nИНСТРУКЦИЯ: Цитируй как [1], [2]\n\n"
                    for i, a in enumerate(st.session_state['deep_analysis_articles'], 1):
                        at += f"## [{i}] {a['title']}\n**Авторы:** {a['authors']} | **Год:** {a['year']}\n\n{a.get('full_text', '')[:10000]}\n\n---\n\n"
                    context_parts.append(at)
                
                if file_type in ['document', 'image'] and file_text:
                    context_parts.append(f"# 📄 ФАЙЛ\n\n{file_text}\n\n")
                
                full_context = "\n".join(context_parts) if context_parts else ""
                
                instructions = """# ️ ИНСТРУКЦИИ:
1. Используй ТОЛЬКО информацию из источников.
2. НЕ выдумывай факты.
3. Цитируй источники.
4. Соблюдай стиль выбранной ветви.

"""
                full_context_with_instructions = full_context + instructions if full_context else instructions
                
                chatgpt_prompt, claude_prompt, qwen_prompt = get_branch_prompts(
                    branch, task, full_context_with_instructions)
                
                with st.spinner("📝 ChatGPT пишет..."):
                    draft = ask(MODEL_CHATGPT, chatgpt_prompt)
                st.markdown(f"""<div class="ai-card card-chatgpt"><h3 style="margin: 0 0 15px 0; color: #D4A373;">1️⃣ Черновик от ChatGPT</h3></div>""", unsafe_allow_html=True)
                st.write(draft)
                
                with st.spinner("🔍 Claude проверяет..."):
                    critique = ask(MODEL_CLAUDE, claude_prompt.replace('{draft}', draft))
                st.markdown(f"""<div class="ai-card card-claude"><h3 style="margin: 0 0 15px 0; color: #E07A5F;">2️ Критика от Claude</h3></div>""", unsafe_allow_html=True)
                st.write(critique)
                
                with st.spinner(" Qwen готовит финал..."):
                    final = ask(MODEL_QWEN, qwen_prompt.replace('{draft}', draft).replace('{critique}', critique))
                st.markdown(f"""<div class="ai-card card-qwen"><h3 style="margin: 0 0 15px 0; color: #8B5E3C;">3️⃣ ✅ Финал от Qwen</h3></div>""", unsafe_allow_html=True)
                st.write(final)
                
                doc = docx.Document()
                doc.add_heading(f'Задача ({branch_names[branch]}):', level=1)
                doc.add_paragraph(task)
                if file_name: doc.add_heading('Файл:', level=2); doc.add_paragraph(file_name)
                doc.add_heading('1. Черновик', level=2); doc.add_paragraph(draft)
                doc.add_heading('2. Критика', level=2); doc.add_paragraph(critique)
                doc.add_heading('3. Финал', level=2); doc.add_paragraph(final)
                
                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                    doc.save(tmp.name); tmp_path = tmp.name
                with open(tmp_path, "rb") as f:
                    st.download_button(label="📥 Скачать в Word", data=f,
                        file_name=f"Результат_{branch}_{datetime.now().strftime('%d%m%Y_%H%M')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                os.remove(tmp_path)
                
                # Сохраняем с уникальным ID
                entry_id = str(uuid.uuid4())[:8]
                save_to_history({"id": entry_id, "project_id": selected_project["id"], 
                    "date": datetime.now().strftime("%d.%m.%Y %H:%M"),
                    "task": task, "file": file_name, "branch": branch, 
                    "draft": draft, "critique": critique, "final": final})
                st.sidebar.success(" Сохранено!")
                
            except Exception as e: st.error(f"❌ Ошибка: {str(e)}")
