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

# --- КАСТОМНЫЙ CSS: ХАЙТЕК СТИЛЬ ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');
    
    .stApp {
    background: linear-gradient(135deg, #0f0c29 0%, #302b63 50%, #24243e 100%) !important;
    font-family: 'Inter', 'Segoe UI', sans-serif !important;
    max-width: 100% !important;
    width: 100% !important;
    padding: 0 !important;
}

.stApp .block-container {
    max-width: 100% !important;
    width: 100% !important;
    padding: 20px 40px !important;
}
    .stApp p, .stApp div, .stApp label, .stApp h1, .stApp h2, .stApp h3 {
        color: #E0E0E0 !important;
    }
    
    h1, h2, h3 {
        color: #00D9FF !important;
        font-weight: 700 !important;
        text-shadow: 0 0 20px rgba(0, 217, 255, 0.5);
    }
    
    section[data-testid="stSidebar"] {
        background: rgba(15, 12, 41, 0.95) !important;
        border-right: 2px solid rgba(0, 217, 255, 0.3) !important;
        backdrop-filter: blur(10px);
    }
    
   .stButton>button {
    background: linear-gradient(135deg, rgba(102, 126, 234, 0.8) 0%, rgba(118, 75, 162, 0.8) 100%) !important;
    color: #FFFFFF !important;
    border: 2px solid rgba(0, 217, 255, 0.4) !important;
    border-radius: 14px !important;
    padding: 14px 28px !important;
    font-weight: 600 !important;
    font-size: 15px !important;
    transition: all 0.3s ease !important;
    box-shadow: 0 4px 20px rgba(102, 126, 234, 0.3) !important;
    letter-spacing: 1px;
    width: 100% !important;
}

.stButton>button:hover {
    background: linear-gradient(135deg, rgba(118, 75, 162, 0.9) 0%, rgba(102, 126, 234, 0.9) 100%) !important;
    transform: translateY(-2px) !important;
    box-shadow: 0 8px 30px rgba(102, 126, 234, 0.5) !important;
    border-color: #00D9FF !important;
}
    
    .ai-card {
        background: rgba(255, 255, 255, 0.05) !important;
        border-radius: 20px !important;
        padding: 28px !important;
        margin: 20px 0 !important;
        box-shadow: 0 8px 32px rgba(0, 0, 0, 0.3) !important;
        backdrop-filter: blur(10px) !important;
        border: 1px solid rgba(0, 217, 255, 0.2) !important;
        transition: all 0.3s ease !important;
    }
    
    .ai-card:hover {
        border-color: rgba(0, 217, 255, 0.5) !important;
        box-shadow: 0 12px 40px rgba(0, 217, 255, 0.2) !important;
    }
    
    .card-chatgpt { border-left: 5px solid #FFD700 !important; }
    .card-claude { border-left: 5px solid #FF6B6B !important; }
    .card-qwen { border-left: 5px solid #00D9FF !important; background: rgba(0, 217, 255, 0.05) !important; }
    .card-external { border-left: 5px solid #00FF88 !important; background: rgba(0, 255, 136, 0.05) !important; }
    .card-search { border-left: 5px solid #00D9FF !important; background: rgba(0, 217, 255, 0.05) !important; }
    .card-library { border-left: 5px solid #B388FF !important; background: rgba(179, 136, 255, 0.05) !important; }
    .card-excel { border-left: 5px solid #00FF88 !important; background: rgba(0, 255, 136, 0.05) !important; }
    .card-finance { border-left: 5px solid #FF6B6B !important; background: rgba(255, 107, 107, 0.05) !important; }
    .card-api { border-left: 5px solid #00D9FF !important; background: rgba(0, 217, 255, 0.05) !important; }

    .stTextArea textarea {
    border-radius: 12px !important;
    border: 2px solid rgba(0, 217, 255, 0.3) !important;
    background: rgba(255, 255, 255, 0.95) !important;
    color: #1a1a1a !important;
    padding: 15px !important;
    font-size: 16px !important;
}

.stTextArea textarea:focus {
    border-color: #00D9FF !important;
    box-shadow: 0 0 15px rgba(0, 217, 255, 0.3) !important;
}

/* Обычные текстовые поля (однострочные) */
.stTextInput input {
    border-radius: 12px !important;
    border: 2px solid rgba(0, 217, 255, 0.3) !important;
    background: rgba(255, 255, 255, 0.95) !important;
    color: #1a1a1a !important;
    padding: 12px 15px !important;
    font-size: 15px !important;
}

.stTextInput input:focus {
    border-color: #00D9FF !important;
    box-shadow: 0 0 15px rgba(0, 217, 255, 0.3) !important;
}
    
    .stTextArea textarea:focus {
        border-color: #00D9FF !important;
        box-shadow: 0 0 15px rgba(0, 217, 255, 0.3) !important;
    }
    
    .stSuccess {
        background: rgba(0, 255, 136, 0.15) !important;
        color: #00FF88 !important;
        border-radius: 12px !important;
        border: 1px solid rgba(0, 255, 136, 0.3) !important;
    }
    
    .stError {
        background: rgba(255, 107, 107, 0.15) !important;
        color: #FF6B6B !important;
        border-radius: 12px !important;
        border: 1px solid rgba(255, 107, 107, 0.3) !important;
    }
    
    .stWarning {
        background: rgba(255, 215, 0, 0.15) !important;
        color: #FFD700 !important;
        border-radius: 12px !important;
        border: 1px solid rgba(255, 215, 0, 0.3) !important;
    }
    
    .stInfo {
        background: rgba(0, 217, 255, 0.15) !important;
        color: #00D9FF !important;
        border-radius: 12px !important;
        border: 1px solid rgba(0, 217, 255, 0.3) !important;
    }
    
    hr {
        border: none !important;
        height: 2px !important;
        background: linear-gradient(90deg, transparent, rgba(0, 217, 255, 0.5), transparent) !important;
        margin: 30px 0 !important;
    }
    
    .article-item, .book-item {
        background: rgba(255, 255, 255, 0.05) !important;
        border-radius: 16px !important;
        padding: 20px !important;
        margin: 15px 0 !important;
        border: 1px solid rgba(0, 217, 255, 0.2) !important;
        border-left: 4px solid #00D9FF !important;
        backdrop-filter: blur(5px) !important;
    }
    
    .book-item { border-left-color: #B388FF !important; }
    
  .branch-card {
    background: rgba(255, 255, 255, 0.06) !important;
    border-radius: 20px !important;
    padding: 30px 35px !important;
    margin: 0 15px !important;
    box-shadow: 0 8px 40px rgba(0, 0, 0, 0.4) !important;
    backdrop-filter: blur(10px) !important;
    border: 2px solid rgba(0, 217, 255, 0.2) !important;
    cursor: pointer !important;
    transition: all 0.4s ease !important;
    position: relative;
    overflow: hidden;
    height: 280px !important;
    display: flex !important;
    flex-direction: column !important;
    align-items: center !important;
    justify-content: center !important;
    text-align: center !important;
}

.branch-card::before {
    content: '';
    position: absolute;
    top: 0;
    left: -100%;
    width: 100%;
    height: 100%;
    background: linear-gradient(90deg, transparent, rgba(0, 217, 255, 0.08), transparent);
    transition: left 0.6s ease;
}

.branch-card:hover::before {
    left: 100%;
}

.branch-card:hover {
    transform: translateY(-5px) scale(1.02) !important;
    box-shadow: 0 15px 50px rgba(0, 217, 255, 0.25) !important;
    border-color: rgba(0, 217, 255, 0.5) !important;
}

.branch-science { 
    border-left: 6px solid #B388FF !important;
    background: linear-gradient(135deg, rgba(179, 136, 255, 0.1) 0%, rgba(255, 255, 255, 0.03) 100%) !important;
}

.branch-legal { 
    border-left: 6px solid #00D9FF !important;
    background: linear-gradient(135deg, rgba(0, 217, 255, 0.1) 0%, rgba(255, 255, 255, 0.03) 100%) !important;
}

.branch-business { 
    border-left: 6px solid #00FF88 !important;
    background: linear-gradient(135deg, rgba(0, 255, 136, 0.1) 0%, rgba(255, 255, 255, 0.03) 100%) !important;
}

.branch-science h3 { color: #B388FF !important; text-shadow: 0 0 25px rgba(179, 136, 255, 0.6) !important; }
.branch-legal h3 { color: #00D9FF !important; text-shadow: 0 0 25px rgba(0, 217, 255, 0.6) !important; }
.branch-business h3 { color: #00FF88 !important; text-shadow: 0 0 25px rgba(0, 255, 136, 0.6) !important; }
    
    .branch-card::before {
        content: '';
        position: absolute;
        top: 0;
        left: -100%;
        width: 100%;
        height: 100%;
        background: linear-gradient(90deg, transparent, rgba(0, 217, 255, 0.1), transparent);
        transition: left 0.5s ease;
    }
    
    .branch-card:hover::before {
        left: 100%;
    }
    
    .branch-card:hover {
        transform: translateY(-5px) scale(1.02) !important;
        box-shadow: 0 15px 50px rgba(0, 217, 255, 0.3) !important;
        border-color: #00D9FF !important;
    }
    
    .branch-science { 
        border-left: 8px solid #B388FF !important;
        background: linear-gradient(135deg, rgba(179, 136, 255, 0.1) 0%, rgba(255, 255, 255, 0.05) 100%) !important;
    }
    
    .branch-legal { 
        border-left: 8px solid #00D9FF !important;
        background: linear-gradient(135deg, rgba(0, 217, 255, 0.1) 0%, rgba(255, 255, 255, 0.05) 100%) !important;
    }
    
    .branch-business { 
        border-left: 8px solid #00FF88 !important;
        background: linear-gradient(135deg, rgba(0, 255, 136, 0.1) 0%, rgba(255, 255, 255, 0.05) 100%) !important;
    }
    
    .branch-science h3 { color: #B388FF !important; text-shadow: 0 0 20px rgba(179, 136, 255, 0.5) !important; }
    .branch-legal h3 { color: #00D9FF !important; text-shadow: 0 0 20px rgba(0, 217, 255, 0.5) !important; }
    .branch-business h3 { color: #00FF88 !important; text-shadow: 0 0 20px rgba(0, 255, 136, 0.5) !important; }
    
    .history-btn-date {
        font-size: 11px !important;
        color: #888 !important;
        font-weight: 400 !important;
    }
    .branch-grain { 
    border-left: 6px solid #FFD700 !important;
    background: linear-gradient(135deg, rgba(255, 215, 0, 0.1) 0%, rgba(255, 255, 255, 0.03) 100%) !important;
}
.branch-grain h3 { color: #FFD700 !important; text-shadow: 0 0 25px rgba(255, 215, 0, 0.6) !important; }
    
    /* Стили для фонового изображения на главной */
   .main-header {
    background: linear-gradient(90deg, rgba(15, 12, 41, 0.9) 0%, rgba(48, 43, 99, 0.8) 50%, rgba(36, 36, 62, 0.9) 100%), 
                url('https://images.unsplash.com/photo-1480714378408-67cf0d13bc1b?w=1920&q=80') center/cover;
    border-radius: 16px;
    padding: 30px 50px;
    margin: 0 0 30px 0;
    box-shadow: 0 8px 40px rgba(0, 0, 0, 0.5);
    border: 1px solid rgba(0, 217, 255, 0.3);
    display: flex;
    align-items: center;
    justify-content: space-between;
    width: 100%;
    box-sizing: border-box;
}

.main-header h1 {
    font-size: 2.5rem !important;
    margin: 0 !important;
    background: linear-gradient(135deg, #00D9FF 0%, #B388FF 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    white-space: nowrap;
    letter-spacing: 2px;
}

.main-header p {
    font-size: 1.1rem !important;
    color: rgba(255, 255, 255, 0.85) !important;
    margin: 0 !important;
    text-shadow: 0 1px 5px rgba(0, 0, 0, 0.5);
}
    
    /* Неоновые эффекты для кнопок */
    .neon-button {
        position: relative;
        overflow: hidden;
    }
    
    .neon-button::after {
        content: '';
        position: absolute;
        top: 50%;
        left: 50%;
        width: 0;
        height: 0;
        border-radius: 50%;
        background: rgba(0, 217, 255, 0.3);
        transform: translate(-50%, -50%);
        transition: width 0.6s ease, height 0.6s ease;
    }
    
    .neon-button:hover::after {
        width: 300px;
        height: 300px;
    }
    
    /* Анимация для карточек */
    @keyframes glow {
        0%, 100% { box-shadow: 0 0 20px rgba(0, 217, 255, 0.3); }
        50% { box-shadow: 0 0 40px rgba(0, 217, 255, 0.6); }
    }
    
    .glow-effect {
        animation: glow 2s ease-in-out infinite;
    }
/* Выравнивание колонок Streamlit */
[data-testid="column"] {
    display: flex !important;
    flex-direction: column !important;
}

[data-testid="column"] > div {
    flex: 1 !important;
    display: flex !important;
    flex-direction: column !important;
}

/* Кнопки прижаты к низу */
[data-testid="column"] .stButton {
    margin-top: auto !important;
}

[data-testid="column"] .stButton button {
    width: 100% !important;
}
</style>
""", unsafe_allow_html=True)

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
# --- ФУНКЦИИ ДЛЯ ЗЕРНОВОГО БИЗНЕСА ---
def normalize_address(address):
    """Заменяет сокращения на полные названия для лучшего поиска"""
    replacements = [
        ('х.', 'хутор'), ('ст-ца.', 'станица'), ('ст.', 'станица'), ('обл.', 'область'),
        ('г.', 'город'), ('с.', 'село'), ('п.', 'посёлок'), ('пгт.', 'посёлок городского типа'),
        ('р-н.', 'район'), ('р-он.', 'район'), ('ул.', 'улица'), ('д.', 'дом'),
        ('мкр.', 'микрорайон')
    ]
    result = address
    for short, full in replacements:
        result = result.replace(short, full)
    return result

def load_yandex_key():
    try: return st.secrets["YANDEX_API_KEY"]
    except Exception: pass
    try:
        with open("yandex_key.txt", "r", encoding="utf-8") as f: return f.read().strip()
    except FileNotFoundError: return None

def geocode_yandex(address, key):
    try:
        url = "https://geocode-maps.yandex.ru/1.x/"
        params = {'apikey': key, 'geocode': address, 'format': 'json', 'results': 1}
        response = requests.get(url, params=params, timeout=15)
        response.raise_for_status()
        data = response.json()
        members = data['response']['GeoObjectCollection']['featureMember']
        if members:
            pos = members[0]['GeoObject']['Point']['pos']
            lon, lat = pos.split(' ')
            return float(lat), float(lon), "Яндекс.Карты"
    except Exception:
        pass
    return None

def geocode_photon(address):
    try:
        url = "https://photon.komoot.io/api/"
        params = {'q': address, 'limit': 1}
        response = requests.get(url, params=params, timeout=15)
        response.raise_for_status()
        data = response.json()
        features = data.get('features', [])
        if features:
            coords = features[0].get('geometry', {}).get('coordinates')
            if coords:
                return float(coords[1]), float(coords[0]), "Photon"
    except Exception:
        pass
    return None

def geocode_nominatim(address):
    try:
        url = "https://nominatim.openstreetmap.org/search"
        params = {'q': address, 'format': 'json', 'limit': 1}
        headers = {'User-Agent': 'ai-council-grain/1.0'}
        response = requests.get(url, params=params, headers=headers, timeout=15)
        response.raise_for_status()
        data = response.json()
        if data:
            return float(data[0]['lat']), float(data[0]['lon']), "OpenStreetMap"
    except Exception:
        pass
    return None

def geocode_address(address):
    """Пробует геокодеры по очереди: Яндекс - Photon - OpenStreetMap"""
    yandex_key = load_yandex_key()
    variants = [address, normalize_address(address), address + ", Россия"]
    
    if yandex_key:
        for variant in variants:
            result = geocode_yandex(variant, yandex_key)
            if result: return result
    
    for variant in variants:
        result = geocode_photon(variant)
        if result: return result
    
    for variant in variants:
        result = geocode_nominatim(variant)
        if result: return result
    
    return None, None, None

def get_route_distance_km(addr1, addr2):
    """Считает расстояние по дороге между адресами"""
    lat1, lon1, src1 = geocode_address(addr1)
    lat2, lon2, src2 = geocode_address(addr2)
    if not lat1 or not lat2:
        return None, "Не удалось найти один из адресов. Проверь написание или введи км вручную.", None
    try:
        url = f"https://router.project-osrm.org/route/v1/driving/{lon1},{lat1};{lon2},{lat2}"
        response = requests.get(url, timeout=15)
        response.raise_for_status()
        data = response.json()
        if data.get('routes'):
            return data['routes'][0]['distance'] / 1000.0, None, f"{src1} - {src2}"
        return None, "Маршрут не построен.", None
    except Exception as e:
        return None, f"Ошибка маршрута: {str(e)}", None
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

def create_project(name, description="", branch=None):
    projects = load_projects()
    new_project = {"id": str(uuid.uuid4())[:8], "name": name, "description": description, "branch": branch, "created": datetime.now().strftime("%d.%m.%Y %H:%M")}
    projects.append(new_project); save_projects(projects); return new_project

def delete_project(project_id):
    projects = [p for p in load_projects() if p["id"] != project_id]; save_projects(projects)
    history = [h for h in load_history() if h.get("project_id") != project_id]; save_history(history)

def get_projects_by_branch(branch):
    all_projects = load_projects()
    if branch is None:
        return [p for p in all_projects if not p.get('branch')]
    return [p for p in all_projects if p.get('branch') == branch]

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
    history = load_history()
    for h in history:
        if h.get("id") == entry_id: return h
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

def ask(model, prompt, max_tokens=32000):
    key = load_key()
    if not key: raise Exception("Не найден key.txt или секрет API_KEY")
    client = OpenAI(api_key=key, base_url=BASE_URL)
    response = client.chat.completions.create(
        model=model,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.3,
        max_tokens=max_tokens
    )
    return response.choices[0].message.content
    def ask_with_continuation(model, prompt, max_tokens=32000, max_attempts=5):
        """Отправляет запрос и автоматически продолжает, если ответ обрезался"""
    full_text = ""
    last_response = ""
    
    for attempt in range(max_attempts):
        if attempt == 0:
            current_prompt = prompt
        else:
            # Берём последние 500 символов предыдущего ответа
            tail = last_response[-500:] if len(last_response) > 500 else last_response
            current_prompt = f"ПРОДОЛЖИ ТЕКСТ С ТОГО МЕСТА, ГДЕ ТЫ ОСТАНОВИЛСЯ. НЕ ПОВТОРЯЙ НАЧАЛО. Последний абзац обрывался так:\n\n...{tail}\n\nПродолжай с того же места:"
        
        response = ask(model, current_prompt, max_tokens=max_tokens)
        full_text += response
        last_response = response
        
        # Проверяем, не обрывается ли ответ на середине предложения
        if response.rstrip().endswith(('.', '!', '?', '”', '"', ']', ')')):
            break  # Ответ завершён нормально
    
    return full_text

def analyze_image(file):
    key = load_key()
    if not key: raise Exception("Не найден key.txt")
    image_bytes = file.read(); base64_image = base64.b64encode(image_bytes).decode('utf-8')
    fn = file.name.lower(); mime = "image/jpeg" if fn.endswith(('.jpg', '.jpeg')) else "image/png"
    client = OpenAI(api_key=key, base_url=BASE_URL)
    response = client.chat.completions.create(model=MODEL_CHATGPT, messages=[{
        "role": "user", "content": [
            {"type": "text", "text": "Распознай весь текст на изображении. Опиши таблицы, схемы, подписи."},
            {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{base64_image}"}}
        ]}], max_tokens=2000)
    return response.choices[0].message.content

# --- ПРОМПТЫ ДЛЯ РАЗНЫХ ВЕТВЕЙ ---
def get_branch_prompts(branch, task, full_context, draft=None, critique=None):
    if branch == 'science':
        mode = st.session_state.get('science_mode', 'Написать статью с нуля')
        
        if mode == "Доработать существующую статью по рецензии":
            chatgpt_prompt = f"""Ты - научный ассистент, эксперт по доработке статей. Твоя задача - доработать существующую статью с учётом замечаний рецензента.

{full_context}

## ЗАДАЧА: {task}

## АЛГОРИТМ РАБОТЫ:
1. Внимательно изучи ИСХОДНУЮ статью (она должна быть среди загруженных файлов)
2. Изучи РЕЦЕНЗИЮ/ЗАМЕЧАНИЯ (тоже в загруженных файлах)
3. Для каждого замечания из рецензии:
   - Пойми суть замечания
   - Найди место в статье, которое нужно доработать
   - Внеси правки, сохранив стиль и логику исходной статьи
4. Составь ОТВЕТ НА РЕЦЕНЗИЮ - таблицу "Замечание - Как исправлено"

## ТРЕБОВАНИЯ:
- Сохраняй стиль исходной статьи
- Не переписывай то, что не просили дорабатывать
- Цитируй новые источники как [1], [2] и т.д.
- В конце - полный СПИСОК ЛИТЕРАТУРЫ (включая новые источники)
- В самом конце - ОТВЕТ НА РЕЦЕНЗИЮ в виде таблицы

Выдай доработанную статью целиком + ответ на рецензию."""
        else:
            chatgpt_prompt = f"""Ты - научный ассистент и исследователь. Подготовь ЧЕРНОВИК НАУЧНОЙ РАБОТЫ.

{full_context}

## ЗАДАЧА: {task}

## СТРУКТУРА:
1. Введение
2. Обзор литературы
3. Методология
4. Результаты
5. Выводы
6. СПИСОК ЛИТЕРАТУРЫ

Цитируй: книги [К1], статьи [1], данные [Данные]. Научный стиль."""

        claude_prompt = f"""Ты - строгий научный рецензент. Проверь результат.

{full_context}

## ЗАДАЧА: {task}
## РЕЗУЛЬТАТ: {draft}

Проверь:
1. Если это доработка по рецензии - учтены ли ВСЕ замечания?
2. Корректность цитат и ссылок
3. Логика и структура
4. Список литературы
Будь строгим."""

        qwen_prompt = f"""Ты - главный редактор научного журнала. Создай ФИНАЛЬНУЮ ВЕРСИЮ.

{full_context}

## ЗАДАЧА: {task}
## ЧЕРНОВИК: {draft}
## КРИТИКА: {critique}

Исправь замечания, добавь цитаты, оформи список литературы.
Если это доработка по рецензии - добавь в конец ОТВЕТ НА РЕЦЕНЗИЮ."""

    elif branch == 'legal':
        chatgpt_prompt = f"""Ты - практикующий юрист с 15-летним стажем в судебных спорах. Твоя задача - изучить ВСЕ загруженные документы и подготовить юридически грамотный отзыв.

{full_context}

## ЗАДАЧА: {task}

## АЛГОРИТМ РАБОТЫ С НЕСКОЛЬКИМИ ДОКУМЕНТАМИ:
1. Изучи ВСЕ загруженные документы по порядку (исковое заявление, отзывы, возражения и т.д.)
2. Выяви позицию каждой стороны
3. Определи ключевые правовые споры между сторонами
4. Найди слабые места в позициях оппонентов
5. Сформулируй правовую позицию по задаче

## ТРЕБОВАНИЯ:
- Строгий юридический стиль (НЕ научный!)
- ОБЯЗАТЕЛЬНЫЕ ссылки на конкретные статьи законов (ГК РФ, АПК РФ, ГПК РФ, УК РФ)
- Ссылки на актуальную СУДЕБНУЮ ПРАКТИКУ (Постановления Пленумов ВС РФ, определения ВС РФ, обзоры практики)
- Структура отзыва:
  1. Вводная часть (суть дела, стороны, предмет спора)
  2. Фактические обстоятельства (краткое изложение)
  3. Правовое обоснование (со ссылками на нормы и практику)
  4. Мотивировочная часть (почему требования обоснованны/необоснованны)
  5. Просительная часть
  6. Приложения (перечень)
- НЕ пиши научную статью - пиши готовый процессуальный документ!

Подготовь полный отзыв/возражение/иск - в зависимости от задачи."""

        claude_prompt = f"""Ты - старший партнёр юридической фирмы, эксперт по судебной практике. Проверь подготовленный документ.

{full_context}

## ЗАДАЧА: {task}
## ДОКУМЕНТ: {draft}

## ПРОВЕРЬ:
1. Учтены ли ВСЕ доводы из загруженных документов?
2. Все ли ссылки на законы корректны и актуальны?
3. Приведена ли релевантная судебная практика?
4. Нет ли юридических рисков и двусмысленностей?
5. Соблюдена ли структура процессуального документа?
6. Соответствует ли документ актуальному законодательству РФ?
7. Какие пункты могут быть оспорены в суде?
8. Что упущено?

Будь МАКСИМАЛЬНО строгим - как судья."""

        qwen_prompt = f"""Ты - старший партнёр юрфирмы. Создай ФИНАЛЬНУЮ ВЕРСИЮ юридического документа.

{full_context}

## ЗАДАЧА: {task}
## ПРОЕКТ: {draft}
## ЗАМЕЧАНИЯ: {critique}

## ТРЕБОВАНИЯ:
1. Исправь ВСЕ замечания
2. Добавь недостающие ссылки на нормы права
3. Добавь ссылки на релевантную судебную практику (Постановления Пленумов ВС РФ, определения ВС РФ, обзоры судебной практики)
4. Устрани все двусмысленности
5. Строгий юридический стиль
6. Документ должен быть ГОТОВ К ПОДАЧЕ В СУД
7. НЕ пиши научную статью!

Создай финальную версию документа."""

    else:
        chatgpt_prompt = f"""Ты - бизнес-аналитик и финансовый консультант уровня McKinsey/BCG с опытом CFO. Проведи глубокий ФИНАНСОВЫЙ АНАЛИЗ.

{full_context}

## ЗАДАЧА: {task}

## ТРЕБОВАНИЯ:

### Если анализируешь банковскую выписку:
1. Категоризируй транзакции (аренда, зарплаты, закупки, маркетинг, налоги)
2. Найди аномалии и подозрительные операции
3. Проанализируй cash flow по неделям/месяцам
4. Выяви причины убытков - конкретные статьи расходов
5. Сравни с отраслевыми нормами
6. Дай рекомендации по оптимизации

### Если создаёшь бизнес-план:
1. Резюме (Executive Summary)
2. Описание бизнеса и концепция
3. Анализ рынка и конкурентов
4. Финансовый план (выручка, расходы, точка безубыточности, NPV, IRR)
5. Оценка рисков
6. План действий на 3/6/12 месяцев

### Если проводишь финансовый анализ:
1. Ключевые метрики (рентабельность, ликвидность)
2. Вертикальный анализ (структура расходов)
3. Горизонтальный анализ (динамика)
4. Сравнение с отраслью
5. Проблемные зоны
6. Рекомендации

## СТРУКТУРА ОТЧЁТА:
1. Резюме
2. Детальный анализ
3. Ключевые выводы
4. Рекомендации

Используй реальные данные, цифры, ссылки на источники. Деловой стиль."""

        claude_prompt = f"""Ты - финансовый аудитор из Big4 (Deloitte, PwC). Проверь отчёт.

{full_context}

## ЗАДАЧА: {task}
## ОТЧЁТ: {draft}

## ПРОВЕРЬ:
1. Корректность расчётов
2. Актуальность данных
3. Реалистичность прогнозов
4. Выявлены ли все проблемы
5. Причины убытков (если есть)
6. Учтены ли все статьи расходов
7. Сравнение с отраслью
8. Реалистичность рекомендаций
9. Категоризация транзакций (если выписка)
10. Найденные аномалии

Будь МАКСИМАЛЬНО строгим."""

        qwen_prompt = f"""Ты - инвестиционный консультант и стратег с опытом CFO. Создай ФИНАЛЬНУЮ ВЕРСИЮ.

{full_context}

## ЗАДАЧА: {task}
## ЧЕРНОВИК: {draft}
## ЗАМЕЧАНИЯ АУДИТОРА: {critique}

## ТРЕБОВАНИЯ:
1. Исправь все замечания
2. Добавь финансовые расчёты
3. Используй актуальные данные (ЦБ РФ, Yahoo Finance)
4. Если есть выписка - полный анализ транзакций
5. Если есть бизнес-план - интегрируй финмодель (NPV, IRR)
6. Деловой стиль, цифры, рекомендации
7. Структура: Резюме - Анализ - Выводы - Рекомендации

Если анализируешь убытки - дай рекомендации по каждой статье.
Если бизнес-план - покажи точку безубыточности и срок окупаемости.

Создай финальный отчёт для презентации."""

    return chatgpt_prompt, claude_prompt, qwen_prompt

# --- ОТОБРАЖЕНИЕ ОТКРЫТОЙ БЕСЕДЫ ---
def display_opened_entry(entry):
    if not entry:
        return
    branch = entry.get('branch', 'science')
    branch_names = {'science': '[НАУКА] Наука', 'legal': '[ПРАВО] Юридическая практика', 'business': '[БИЗНЕС] Бизнес-консалтинг', 'grain': '[ЗЕРНО] Зерновой бизнес'}
    branch_colors = {'science': '#9B59B6', 'legal': '#2C3E50', 'business': '#27AE60', 'grain': '#FFD700'}
    
    st.markdown(f"""<div class="ai-card" style="border-left: 8px solid {branch_colors.get(branch, '#8B5E3C')};">
        <h3 style="margin: 0; color: {branch_colors.get(branch, '#8B5E3C')};"> {entry.get('task', '')[:80]}</h3>
        <p style="margin: 10px 0 0 0; color: #A89F91;">{branch_names.get(branch, '')} - {entry.get('date', '')}</p>
    </div>""", unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("### 📝 Задача")
    st.info(entry.get('task', ''))
    if entry.get('file'):
        st.markdown(f"**[ФАЙЛ] Файл:** {entry['file']}")
    st.markdown("---")
    st.markdown("""<div class="ai-card card-chatgpt"><h3 style="margin: 0 0 15px 0; color: #D4A373;">1. Черновик от ChatGPT</h3></div>""", unsafe_allow_html=True)
    st.write(entry.get('draft', ''))
    st.markdown("---")
    st.markdown("""<div class="ai-card card-claude"><h3 style="margin: 0 0 15px 0; color: #E07A5F;">2. Критика от Claude</h3></div>""", unsafe_allow_html=True)
    st.write(entry.get('critique', ''))
    st.markdown("---")
    st.markdown("""<div class="ai-card card-qwen"><h3 style="margin: 0 0 15px 0; color: #8B5E3C;">3. Финал от Qwen</h3></div>""", unsafe_allow_html=True)
    st.write(entry.get('final', ''))
    st.markdown("---")
    
    doc = docx.Document()
    doc.add_heading(f'Задача ({branch_names.get(branch, "")}):', level=1)
    doc.add_paragraph(entry.get('task', ''))
    if entry.get('file'):
        doc.add_heading('Файл:', level=2)
        doc.add_paragraph(entry['file'])
    doc.add_heading('1. Черновик', level=2)
    doc.add_paragraph(entry.get('draft', ''))
    doc.add_heading('2. Критика', level=2)
    doc.add_paragraph(entry.get('critique', ''))
    doc.add_heading('3. Финал', level=2)
    doc.add_paragraph(entry.get('final', ''))
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        doc.save(tmp.name)
        tmp_path = tmp.name
    with open(tmp_path, "rb") as f:
        st.download_button(label="📥 Скачать в Word", data=f,
            file_name=f"Беседа_{entry.get('id', '')[:6]}_{datetime.now().strftime('%d%m%Y')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    os.remove(tmp_path)

# --- ИНТЕРФЕЙС ---
if 'branch' not in st.session_state: st.session_state['branch'] = None
if 'open_entry' not in st.session_state: st.session_state['open_entry'] = None
if 'selected_project' not in st.session_state: st.session_state['selected_project'] = None
if 'show_new_chat' not in st.session_state: st.session_state['show_new_chat'] = False

# --- БОКОВАЯ ПАНЕЛЬ ---
st.sidebar.header(" Проекты")
current_branch = st.session_state.get('branch')
branch_projects = get_projects_by_branch(current_branch)

with st.sidebar.expander(" Новый проект"):
    new_project_name = st.text_input("Название проекта:", placeholder="Например: Научная статья")
    new_project_desc = st.text_area("Описание:", height=80)
    if st.button("Создать проект"):
        if new_project_name.strip():
            create_project(new_project_name.strip(), new_project_desc.strip(), branch=current_branch)
            st.success(f"Проект '{new_project_name}' создан!"); st.rerun()
        else: st.warning("Введите название")

st.sidebar.markdown("---")

if branch_projects:
    project_names = [f"{p['name']}" for p in branch_projects]
    selected_name = st.sidebar.selectbox("Выбери проект:", ["- Без проекта (быстрый диалог) -"] + project_names)
    if selected_name != "- Без проекта (быстрый диалог) -":
        selected_project = branch_projects[project_names.index(selected_name)]
        st.session_state['selected_project'] = selected_project
    else:
        selected_project = None; st.session_state['selected_project'] = None
else:
    st.sidebar.info("Нет проектов в этой ветви. Создай первый!"); selected_project = None

if selected_project:
    st.sidebar.caption(f"Проект: {selected_project['name']}")
    if st.sidebar.button("🗑 Удалить проект"):
        if st.sidebar.checkbox("Да, удалить"):
            delete_project(selected_project["id"]); st.session_state['selected_project'] = None; st.rerun()

st.sidebar.markdown("---")

if st.session_state['branch']:
    if st.sidebar.button("💬 Новая беседа", key="new_chat_sidebar", use_container_width=True):
        st.session_state['show_new_chat'] = True; st.session_state['open_entry'] = None; st.rerun()

st.sidebar.markdown("---")

st.sidebar.header(" История бесед")
history = load_history()

if selected_project:
    visible_history = [h for h in history if h.get("project_id") == selected_project["id"]]
    st.sidebar.caption(f"Проект: {selected_project['name']}")
else:
    visible_history = [h for h in history if not h.get("project_id") and h.get("branch") == current_branch]
    st.sidebar.caption("💬 Быстрый диалог")

if visible_history:
    st.sidebar.markdown("---")
    for entry in reversed(visible_history):
        short = entry["task"][:35] + "..." if len(entry["task"]) > 35 else entry["task"]
        date_str = entry.get('date', '')
        if st.sidebar.button(f"📝 {short}\n<span class='history-btn-date'>{date_str}</span>", 
                            key=f"open_{entry.get('id', '')}", use_container_width=True, type="secondary"):
            st.session_state['open_entry'] = entry.get('id'); st.session_state['show_new_chat'] = False; st.rerun()
        col1, col2 = st.sidebar.columns([4, 1])
        with col2:
            if st.button("🗑", key=f"del_{entry.get('id', '')}", help="Удалить"):
                delete_history_entry(entry.get("id"))
                if st.session_state.get('open_entry') == entry.get('id'): st.session_state['open_entry'] = None
                st.sidebar.success("[OK] Удалено"); st.rerun()
else:
    st.sidebar.info("Пока нет бесед")

st.sidebar.markdown("---")
if load_key(): st.sidebar.success("🔑 Ключ активен")
else: st.sidebar.error("[X] Ключ не найден")

if st.session_state['branch']:
    if st.sidebar.button("🏠 Вернуться к выбору ветви"):
        st.session_state['branch'] = None; st.session_state['open_entry'] = None
        st.session_state['selected_project'] = None; st.session_state['show_new_chat'] = False; st.rerun()

# === ГЛАВНАЯ ЛОГИКА ===
if st.session_state['open_entry']:
    entry = get_history_entry(st.session_state['open_entry'])
    if entry:
        display_opened_entry(entry)
        col1, col2 = st.columns(2)
        with col1:
            if st.button("← Закрыть беседу", use_container_width=True):
                st.session_state['open_entry'] = None; st.rerun()
        with col2:
            if st.button("💬 Новая беседа", use_container_width=True):
                st.session_state['open_entry'] = None; st.session_state['show_new_chat'] = True; st.rerun()
        st.stop()
    else:
        st.session_state['open_entry'] = None

if not st.session_state['branch']:
 if not st.session_state['branch']:
    # СТАРТОВАЯ СТРАНИЦА - 4 ВЕТВИ
    st.markdown("""
    <div class="main-header">
        <div>
            <h1 style="margin: 0;">[AI] Совет ИИ</h1>
            <p style="margin: 10px 0 0 0; font-size: 1.1rem; opacity: 0.85;">Три разума. Одна задача. Идеальный результат.</p>
        </div>
        <p style="margin: 0; font-size: 1rem; opacity: 0.75; max-width: 450px; text-align: right; line-height: 1.6;">
            Выбери направление работы.<br>
            В каждой ветви свои проекты, беседы и инструменты.
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown("""
        <div class="branch-card branch-science">
            <h3 style="margin: 0; font-size: 2.2rem;">Наука</h3>
        </div>
        """, unsafe_allow_html=True)
        if st.button("🚀 НАУКА", key="btn_science", use_container_width=True):
            st.session_state['branch'] = 'science'
            st.session_state['selected_project'] = None
            st.session_state['show_new_chat'] = False
            st.rerun()
    
    with col2:
        st.markdown("""
        <div class="branch-card branch-legal">
            <h3 style="margin: 0; font-size: 2.2rem;">Юридическая практика</h3>
        </div>
        """, unsafe_allow_html=True)
        if st.button("[ПРАВО] ЮРИДИЧЕСКАЯ ПРАКТИКА", key="btn_legal", use_container_width=True):
            st.session_state['branch'] = 'legal'
            st.session_state['selected_project'] = None
            st.session_state['show_new_chat'] = False
            st.rerun()
    
    with col3:
        st.markdown("""
        <div class="branch-card branch-business">
            <h3 style="margin: 0; font-size: 2.2rem;">[БИЗНЕС] Бизнес-консалтинг</h3>
        </div>
        """, unsafe_allow_html=True)
        if st.button("[БИЗНЕС] БИЗНЕС-КОНСАЛТИНГ", key="btn_business", use_container_width=True):
            st.session_state['branch'] = 'business'
            st.session_state['selected_project'] = None
            st.session_state['show_new_chat'] = False
            st.rerun()
    
    with col4:
        st.markdown("""
        <div class="branch-card branch-grain">
            <h3 style="margin: 0; font-size: 2.2rem;">[ЗЕРНО] Зерновой бизнес</h3>
        </div>
        """, unsafe_allow_html=True)
        if st.button("[ЗЕРНО] ЗЕРНОВОЙ БИЗНЕС", key="btn_grain", use_container_width=True):
            st.session_state['branch'] = 'grain'
            st.session_state['selected_project'] = None
            st.session_state['show_new_chat'] = False
            st.rerun()
    
    st.stop()
# === РАБОТА ВНУТРИ ВЕТВИ ===
branch = st.session_state['branch']
branch_names = {'science': '[НАУКА] Наука', 'legal': '[ПРАВО] Юридическая практика', 'business': '[БИЗНЕС] Бизнес-консалтинг', 'grain': '[ЗЕРНО] Зерновой бизнес'}
branch_colors = {'science': '#9B59B6', 'legal': '#2C3E50', 'business': '#27AE60', 'grain': '#FFD700'}
project_label = selected_project['name'] if selected_project else '💬 Быстрый диалог'

st.markdown(f"""<div class="ai-card" style="border-left: 8px solid {branch_colors[branch]};">
    <h3 style="margin: 0; color: {branch_colors[branch]};">{branch_names[branch]}</h3>
    <p style="margin: 10px 0 0 0; color: #A89F91;">{project_label}</p>
</div>""", unsafe_allow_html=True)

if st.session_state.get('show_new_chat'):
    st.markdown("""<div class="ai-card" style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; text-align: center;">
        <h3 style="margin: 0; color: white;"> Готов к новой беседе!</h3>
        <p style="margin: 10px 0 0 0; opacity: 0.9;">Введи новую задачу ниже</p>
    </div>""", unsafe_allow_html=True)
    st.session_state['show_new_chat'] = False

st.markdown("---")

# === ИНСТРУМЕНТЫ ВЕТВЕЙ ===
if branch == 'science':
    science_mode = st.radio(
        "Режим работы:",
        ["Написать статью с нуля", "Доработать существующую статью по рецензии", "Научный обзор по источникам"],
        index=0
    )
    st.session_state['science_mode'] = science_mode
    st.markdown("---")
    st.markdown("""<div class="ai-card card-library"><h3 style="margin: 0 0 10px 0; color: #9B59B6;">Моя научная библиотека</h3>
    <p style="margin: 0; color: #A89F91;">Укажи путь к папке с учебниками.</p></div>""", unsafe_allow_html=True)
    library_path = st.text_input("Путь к папке:", placeholder="C:\\Users\\123\\Documents\\Научная_библиотека")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("[ОТКРЫТЬ] Сканировать библиотеку", use_container_width=True):
            if library_path.strip():
                with st.spinner("Сканирую..."):
                    books = scan_library_folder(library_path)
                    if isinstance(books, dict) and 'error' in books: st.error(books['error'])
                    else: st.success(f"[OK] Найдено {len(books)} книг!"); st.session_state['library_books'] = books
            else: st.warning("Введи путь")
    with col2:
        use_library = st.checkbox("📖 Использовать библиотеку", value=False)
    if 'library_books' in st.session_state:
        st.markdown(f"### [КНИГИ] Найдено: {len(st.session_state['library_books'])} книг")
        for i, book in enumerate(st.session_state['library_books'][:5], 1):
            st.markdown(f"**{i}.** {book['title']} - {book['author']} ({book['year']})")
    st.markdown("---")
    st.markdown("""<div class="ai-card card-search"><h3 style="margin: 0 0 10px 0; color: #5B8DB8;">[НАУКА] Научный поиск</h3>
    <p style="margin: 0; color: #A89F91;">Поиск по научным базам.</p></div>""", unsafe_allow_html=True)
    search_query = st.text_input("Тема поиска:", placeholder="artificial intelligence in civil law")
    col1, col2, col3 = st.columns(3)
    with col1: use_cyberleninka = st.checkbox("🇺 КиберЛенинка", value=True)
    with col2: use_openalex = st.checkbox("🌍 OpenAlex", value=True)
    with col3: use_semantic = st.checkbox("🤖 Semantic Scholar", value=True)
    deep_analysis = st.checkbox("[НАУКА] Глубокое исследование", value=True)
    if st.button("[ПОИСК] Найти статьи", use_container_width=True):
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
                        text = download_and_extract_pdf(article['pdf_url']); article['full_text'] = text; articles_with_text.append(article)
                st.session_state['deep_analysis_articles'] = articles_with_text; st.success(f"[OK] Скачано {len(articles_with_text)}!")
            else:
                st.session_state['search_results'] = all_articles; st.success(f"Найдено: {len(all_articles)}")
        else: st.warning("Введи тему")

elif branch == 'legal':
    st.markdown("""<div class="ai-card card-external"><h3 style="margin: 0 0 10px 0; color: #6B9080;">[ПРАВО] Мнение внешнего ИИ</h3>
    <p style="margin: 0; color: #A89F91;">Cognevia, ПравоТех, Doczilla (необязательно).</p></div>""", unsafe_allow_html=True)
    external_ai_opinion = st.text_area("Ответ от внешнего ИИ:", height=120, placeholder="Скопируй анализ от Cognevia...")
    st.markdown("---")
    st.markdown("""<div class="ai-card"><h3 style="margin: 0 0 10px 0; color: #2C3E50;">[СПИСОК] Тип документа</h3>
    <p style="margin: 0; color: #A89F91;">Выбери тип - поможет ИИ понять структуру.</p></div>""", unsafe_allow_html=True)
    doc_type = st.selectbox("Тип документа:", [
        "Договор (купли-продажи, аренды, услуг)", "Исковое заявление", "Претензия",
        "Жалоба (в суд, прокуратуру)", "Юридическое письмо", "Правовое заключение", "Другое"
    ], index=0)
elif branch == 'grain':
    st.markdown("""<div class="ai-card" style="border-left: 5px solid #FFD700; background: rgba(255, 215, 0, 0.05);">
    <h3 style="margin: 0 0 10px 0; color: #FFD700;">[ЗЕРНО] Калькулятор логистики и закупки зерна</h3>
    <p style="margin: 0; color: #A89F91;">Расчёт расстояния, стоимости перевозки и итоговой цены с НДС.</p></div>""", unsafe_allow_html=True)
    
    # БЛОК 1
    st.markdown("#### 📍 Блок 1. Адреса погрузки и выгрузки")
    col1, col2 = st.columns(2)
    with col1:
        load_addr = st.text_input("Адрес погрузки:", placeholder="Например: Ростовская обл, г. Зерноград, ул. Ленина 1")
    with col2:
        unload_addr = st.text_input("Адрес выгрузки:", placeholder="Например: Краснодарский край, г. Новороссийск")
    
    col3, col4 = st.columns([1, 2])
    with col3:
        if st.button("🗺 Рассчитать расстояние", use_container_width=True):
            if load_addr.strip() and unload_addr.strip():
                with st.spinner("Строю маршрут..."):
                    km, err, src = get_route_distance_km(load_addr, unload_addr)
                    if km:
                        st.session_state['grain_distance'] = km
                        st.success(f"[OK] Расстояние: {km:.1f} км (источник: {src})")
                    else:
                        st.error(err)
            else:
                st.warning("Введи оба адреса")
    with col4:
        manual_km = st.number_input("Или введи км вручную:", min_value=0.0, value=0.0, step=10.0)
        if manual_km > 0:
            st.session_state['grain_distance'] = manual_km
    
    # === ДИАГНОСТИКА ПОИСКА ===
    if st.button("🔧 Диагностика поиска адресов"):
        test_addr = "Ростов-на-Дону"
        st.markdown(f"**Проверка геокодеров на адресе:** {test_addr}")
        
        yk = load_yandex_key()
        st.write(f"**1. Ключ Яндекс найден:** {'ДА [OK]' if yk else 'НЕТ [X]'}")
        if yk:
            try:
                r = requests.get("https://geocode-maps.yandex.ru/1.x/",
                    params={'apikey': yk, 'geocode': test_addr, 'format': 'json', 'results': 1}, timeout=15)
                st.write(f"**Яндекс:** статус {r.status_code} - {r.text[:200]}")
            except Exception as e:
                st.write(f"**Яндекс:** ошибка {str(e)}")
        else:
            st.write("**Яндекс:** пропущен (нет ключа)")
        
        try:
            r = requests.get("https://photon.komoot.io/api/", params={'q': test_addr, 'limit': 1}, timeout=15)
            st.write(f"**2. Photon:** статус {r.status_code} - {r.text[:200]}")
        except Exception as e:
            st.write(f"**2. Photon:** ошибка {str(e)}")
        
        try:
            r = requests.get("https://nominatim.openstreetmap.org/search",
                params={'q': test_addr, 'format': 'json', 'limit': 1},
                headers={'User-Agent': 'ai-council-grain/1.0'}, timeout=15)
            st.write(f"**3. Nominatim:** статус {r.status_code} - {r.text[:200]}")
        except Exception as e:
            st.write(f"**3. Nominatim:** ошибка {str(e)}")
    
    distance_km = st.session_state.get('grain_distance', 0.0)
    if distance_km > 0:
        st.info(f"🛣 Расстояние между точками: **{distance_km:.1f} км**")
    
    st.markdown("---")
    
    # БЛОК 2
    st.markdown("#### 🚛 Блок 2. Стоимость логистики")
    coef = st.number_input("Коэффициент логистики (руб/км за 1 тонну):", min_value=0.0, value=3.5, step=0.1)
    logistics_cost = distance_km * coef
    st.info(f"💰 Стоимость перевозки за 1 тонну: **{logistics_cost:,.2f} руб** (= {distance_km:.1f} км × {coef} руб/км)")
    
    st.markdown("---")
    
    # БЛОК 3
    st.markdown("#### [БИЗНЕС] Блок 3. Твой заработок")
    margin = st.number_input("Заработок за 1 тонну (руб):", min_value=0.0, value=500.0, step=50.0)
    
    st.markdown("---")
    
       # БЛОК 4
    st.markdown("#### [ЗЕРНО] Блок 4. Цена сделки (Z)")
    purchase = st.number_input("Цена, по которой экспортер покупает (руб за 1 тонну):", min_value=0.0, value=12000.0, step=100.0)
    
    st.markdown("---")
    
    # БЛОК 5
    st.markdown("#### [Excel] Блок 5. Итоговая цена")
    price_no_vat = purchase - logistics_cost - margin
    price_vat = price_no_vat * 1.10
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Цена без НДС", f"{price_no_vat:,.2f} руб")
    with col2:
        st.metric("Цена с НДС 10%", f"{price_vat:,.2f} руб")
    
    if price_no_vat < 0:
        st.error("⚠️ Внимание: результат отрицательный! Логистика и заработок превышают цену сделки - сделка убыточна.")
    
    st.markdown("**Детализация расчёта за 1 тонну (формула Z − (X×Y) − A):**")
    st.write(f"- Цена сделки (Z): {purchase:,.2f} руб")
    st.write(f"- Минус логистика (X×Y = {distance_km:.1f} км × {coef} руб/км): −{logistics_cost:,.2f} руб")
    st.write(f"- Минус твой заработок (A): −{margin:,.2f} руб")
    st.write(f"- **= Цена без НДС: {price_no_vat:,.2f} руб**")
    st.write(f"- **Цена с НДС 10%: {price_vat:,.2f} руб**")
elif branch == 'business':
    # Анализ банковских выписок
    st.markdown("""<div class="ai-card card-api"><h3 style="margin: 0 0 10px 0; color: #3498DB;">Анализ банковских выписок</h3>
    <p style="margin: 0; color: #A89F91;">Загрузи выписку (CSV, Excel) - ИИ найдёт причины убытков.</p></div>""", unsafe_allow_html=True)
    bank_statement = st.file_uploader("Загрузи банковскую выписку:", type=['csv', 'xlsx', 'xls'])
    if bank_statement is not None:
        try:
            fn = bank_statement.name.lower()
            if fn.endswith('.csv'):
                df = pd.read_csv(bank_statement, encoding='utf-8', on_bad_lines='skip')
            else:
                df = pd.read_excel(bank_statement)
            st.success(f"[OK] Выписка загружена: {len(df)} транзакций")
            st.info(f"**Столбцы:** {', '.join(df.columns.tolist())}")
            date_cols = [col for col in df.columns if any(k in col.lower() for k in ['дата', 'date', 'время', 'time'])]
            amount_cols = [col for col in df.columns if any(k in col.lower() for k in ['сумма', 'amount', 'списание', 'debit', 'зачисление', 'credit', 'операция'])]
            if date_cols: st.write(f"**Колонка даты:** {date_cols[0]}")
            if amount_cols: st.write(f"**Колонка суммы:** {amount_cols[0]}")
            st.session_state['bank_statement_df'] = df
            st.session_state['bank_date_col'] = date_cols[0] if date_cols else None
            st.session_state['bank_amount_col'] = amount_cols[0] if amount_cols else None
            with st.expander("👁️ Первые 10 транзакций"): st.dataframe(df.head(10))
        except Exception as e: st.error(f"[X] Ошибка: {str(e)}")
    
    st.markdown("---")
    
    # Генератор бизнес-планов
    st.markdown("""<div class="ai-card card-finance"><h3 style="margin: 0 0 10px 0; color: #E74C3C;">[БИЗНЕС] Генератор бизнес-планов</h3>
    <p style="margin: 0; color: #A89F91;">Создай бизнес-план с финансовой моделью.</p></div>""", unsafe_allow_html=True)
    st.markdown("**Параметры бизнес-плана:**")
    col1, col2 = st.columns(2)
    with col1:
        business_name = st.text_input("Название бизнеса:", placeholder="Кофейня в центре")
        business_type = st.selectbox("Тип бизнеса:", ["Розничная торговля", "Общепит", "Услуги", "Производство", "IT", "Другое"])
        initial_investment = st.number_input("Начальные инвестиции (руб):", min_value=0, value=1000000, step=100000)
        monthly_revenue = st.number_input("Месячная выручка (руб):", min_value=0, value=500000, step=50000)
    with col2:
        monthly_costs = st.number_input("Месячные расходы (руб):", min_value=0, value=300000, step=50000)
        growth_rate = st.number_input("Рост (%/мес):", min_value=0, max_value=50, value=5, step=1)
        planning_period = st.number_input("Горизонт (мес):", min_value=6, max_value=60, value=24, step=6)
        discount_rate = st.number_input("Ставка дисконтирования (%):", min_value=0, max_value=100, value=12, step=1)
    if st.button(" Создать бизнес-план с финмоделью", use_container_width=True):
        if business_name.strip():
            with st.spinner("Генерирую..."):
                try:
                    years = max(1, planning_period // 12)
                    excel_path = create_financial_model(initial_investment, monthly_revenue*12, monthly_costs*12, growth_rate*12, discount_rate, years)
                    st.session_state['business_plan_data'] = {
                        'name': business_name, 'type': business_type, 'investment': initial_investment,
                        'monthly_revenue': monthly_revenue, 'monthly_costs': monthly_costs,
                        'growth_rate': growth_rate, 'planning_period': planning_period, 'discount_rate': discount_rate, 'excel_path': excel_path
                    }
                    st.success("[OK] Бизнес-план подготовлен! Опиши задачу ниже.")
                    with open(excel_path, "rb") as f:
                        st.download_button(label="📥 Скачать финмодель Excel", data=f,
                            file_name=f"Финмодель_{business_name}_{datetime.now().strftime('%d%m%Y')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    os.remove(excel_path)
                except Exception as e: st.error(f"[X] Ошибка: {str(e)}")
        else: st.warning("Введи название бизнеса")
    
    st.markdown("---")
    
    # Финансовые данные
    st.markdown("""<div class="ai-card card-api"><h3 style="margin: 0 0 10px 0; color: #3498DB;">Актуальные финансовые данные</h3>
    <p style="margin: 0; color: #A89F91;">ЦБ РФ и Yahoo Finance.</p></div>""", unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1: use_cbr = st.checkbox("🇷🇺 ЦБ РФ", value=True)
    with col2:
        use_yahoo = st.checkbox("🌍 Yahoo Finance", value=False)
        yahoo_tickers = st.text_input("Тикеры конкурентов:", placeholder="SBUX, MCD", value="")
    if st.button(" Получить данные", use_container_width=True):
        financial_parts = []
        if use_cbr:
            with st.spinner("ЦБ РФ..."):
                cbr = get_cbr_data()
                if 'error' not in cbr:
                    text = f"## 🇷🇺 ЦБ РФ ({cbr.get('date', '')})\n**Ключевая ставка:** {cbr.get('key_rate', 'N/A')}%\n**Курсы:**\n"
                    for code, d in cbr.get('currencies', {}).items():
                        s = "📈" if d['change'] > 0 else "[ПАДЕНИЕ]" if d['change'] < 0 else "➡️"
                        text += f"- **{code}**: {d['value']:.2f} руб {s} ({d['change']:+.2f})\n"
                    financial_parts.append(text); st.success(f"[OK] Ставка: {cbr.get('key_rate', 'N/A')}%")
                else: st.error(cbr.get('error'))
        if use_yahoo and yahoo_tickers.strip():
            with st.spinner("Yahoo Finance..."):
                tickers = [t.strip() for t in yahoo_tickers.split(',')]
                yahoo_data = []
                for ticker in tickers:
                    data = get_yahoo_finance_data(ticker, '1mo'); yahoo_data.append(data)
                    if 'error' not in data: st.success(f"[OK] {ticker}: {data.get('current_price', 0):.2f}")
                    else: st.error(f"[X] {ticker}: {data.get('error')}")
                st.session_state['yahoo_data'] = yahoo_data
        if financial_parts or 'yahoo_data' in st.session_state:
            full_fin = "# [Excel] АКТУАЛЬНЫЕ ФИНАНСОВЫЕ ДАННЫЕ\n\n"
            for p in financial_parts: full_fin += p + "\n"
            if 'yahoo_data' in st.session_state:
                full_fin += "##  Yahoo Finance\n"
                for d in st.session_state['yahoo_data']:
                    if 'error' not in d:
                        full_fin += f"### {d.get('ticker', '')} ({d.get('company_name', '')})\n"
                        full_fin += f"- Цена: {d.get('current_price', 0):.2f} {d.get('currency', '')}\n"
                        if d.get('price_change_percent'): full_fin += f"- Изменение: {d['price_change_percent']:+.2f}%\n"
                        full_fin += "\n"
            st.session_state['financial_api_data'] = full_fin; st.success("[OK] Данные добавлены!")
    
    # Быстрый анализ выписки
    if 'bank_statement_df' in st.session_state:
        st.markdown("---")
        st.markdown("""<div class="ai-card" style="border-left: 5px solid #3498DB;"><h3 style="margin: 0 0 10px 0; color: #3498DB;">[ПОИСК] Быстрый анализ выписки</h3></div>""", unsafe_allow_html=True)
        df = st.session_state['bank_statement_df']
        amount_col = st.session_state.get('bank_amount_col')
        if amount_col:
            col1, col2, col3, col4 = st.columns(4)
            with col1: st.metric("Общая сумма", f"{df[amount_col].sum():,.0f} руб")
            with col2: st.metric("Средняя операция", f"{df[amount_col].mean():,.0f} руб")
            with col3: st.metric("Максимальная", f"{df[amount_col].max():,.0f} руб")
            with col4: st.metric("Всего операций", f"{len(df)}")
            st.markdown("**🔝 Топ-10 крупнейших операций:**")
            date_col = st.session_state.get('bank_date_col', df.columns[0])
            cols_to_show = [c for c in [date_col, amount_col] if c in df.columns]
            st.dataframe(df.nlargest(10, amount_col)[cols_to_show])

# === ОБЩИЕ ЭЛЕМЕНТЫ ===
st.markdown("---")
st.markdown("""<div class="ai-card card-excel"><h3 style="margin: 0 0 10px 0; color: #27AE60;">[ФАЙЛ] Загрузи материалы</h3>
<p style="margin: 0; color: #A89F91;">PDF, DOCX, TXT, Excel, фото</p></div>""", unsafe_allow_html=True)
uploaded_files = st.file_uploader(
    "Загрузи материалы (можно несколько файлов сразу):", 
    type=['pdf', 'docx', 'txt', 'jpg', 'jpeg', 'png', 'xlsx', 'xls'],
    accept_multiple_files=True
)
file_text = ""; file_name = ""; file_type = ""; excel_analysis = None

file_texts = []
file_names = []
excel_analysis = None

if uploaded_files:
    st.info(f"[ФАЙЛ] Загружено файлов: {len(uploaded_files)}")
    for idx, uploaded_file in enumerate(uploaded_files):
        try:
            file_name = uploaded_file.name
            file_names.append(file_name)
            fn = file_name.lower()
            
            if fn.endswith(('.jpg', '.jpeg', '.png')):
                with st.spinner(f"Анализирую изображение {file_name}..."):
                    file_text = analyze_image(uploaded_file)
                st.success(f"[OK] Изображение: {file_name}")
                file_texts.append(("image", file_name, file_text))
                
            elif fn.endswith(('.xlsx', '.xls')):
                with st.spinner(f"Анализирую Excel {file_name}..."):
                    file_text, num = extract_excel_data(uploaded_file)
                    excel_analysis = analyze_excel_structure(uploaded_file)
                st.success(f"[OK] Excel: {file_name} ({num} листов)")
                file_texts.append(("excel", file_name, file_text))
                
            else:
                file_text = extract_text(uploaded_file)
                if len(file_text) > MAX_FILE_SIZE:
                    file_text = file_text[:MAX_FILE_SIZE]
                st.success(f"[OK] Файл: {file_name}")
                file_texts.append(("document", file_name, file_text))
                
        except Exception as e:
            st.error(f"[X] Ошибка {file_name}: {e}")
    
    if len(file_texts) > 1:
        st.write("**Загруженные файлы:**")
        for i, (ftype, fname, _) in enumerate(file_texts, 1):
            icon = "[DOC]" if ftype == "document" else "🖼️" if ftype == "image" else "[Excel]"
            st.write(f"{i}. {icon} {fname}")
task = st.text_area("Опиши задачу:", height=150, 
    placeholder={
        'science': "Например: 'Напиши научную статью об ИИ в праве'",
        'legal': "Например: 'Составь договор купли-продажи автомобиля'",
        'business': "Например: 'Проанализируй выписку и найди причины убытков'"
    }.get(branch, ""))
    
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    if st.button("🚀 Запустить обсуждение", use_container_width=True):
        if not task.strip(): st.warning("⚠️ Введи задачу")
        else:
            try:
                context_parts = []
                if 'financial_api_data' in st.session_state and st.session_state['financial_api_data']:
                    context_parts.append(st.session_state['financial_api_data'])
                if branch == 'science' and use_library and 'library_books' in st.session_state:
                    with st.spinner("Ищу книги..."):
                        relevant = search_relevant_books(task, st.session_state['library_books'], top_n=3)
                        if relevant:
                            bt = "# [КНИГИ] БИБЛИОТЕКА\nИНСТРУКЦИЯ: Цитируй как [К1], [К2]\n\n"
                            for i, book in enumerate(relevant, 1):
                                with st.status(f"Загружаю: {book['title'][:50]}..."):
                                    bt += f"## [К{i}] {book['title']}\n**Автор:** {book['author']} | **Год:** {book['year']}\n\n{load_book_content(book['path'])[:10000]}\n\n---\n\n"
                            context_parts.append(bt); st.success(f"[OK] {len(relevant)} книг")
                if branch == 'legal' and 'external_ai_opinion' in dir() and external_ai_opinion.strip():
                    context_parts.append(f"# [ПРАВО] МНЕНИЕ ВНЕШНЕГО ИИ\n\n{external_ai_opinion}\n\n")
                                # === ОБРАБОТКА ВСЕХ ЗАГРУЖЕННЫХ ФАЙЛОВ ===
                for idx, (ftype, fname, ftext) in enumerate(file_texts, 1):
                    if ftype == "excel":
                        ec = f"# [Excel] ФАЙЛ №{idx}: {fname}\n\n"
                        if excel_analysis and excel_analysis.get('has_financial_data'):
                            ec += "⚠️ Финансовые данные. Выяви тренды, метрики, аномалии.\n\n"
                        ec += ftext + "\n\n"
                        context_parts.append(ec)
                    elif ftype in ['document', 'image']:
                        label = "ДОКУМЕНТ" if ftype == "document" else "ИЗОБРАЖЕНИЕ"
                        context_parts.append(f"# [DOC] {label} №{idx}: {fname}\n\n{ftext}\n\n")
                if 'deep_analysis_articles' in st.session_state and st.session_state['deep_analysis_articles']:
                    at = "# [КНИГИ] СТАТЬИ\nИНСТРУКЦИЯ: Цитируй как [1], [2]\n\n"
                    for i, a in enumerate(st.session_state['deep_analysis_articles'], 1):
                        at += f"## [{i}] {a['title']}\n**Авторы:** {a['authors']} | **Год:** {a['year']}\n\n{a.get('full_text', '')[:10000]}\n\n---\n\n"
                    context_parts.append(at)
                if file_type in ['document', 'image'] and file_text:
                    context_parts.append(f"# [DOC] ФАЙЛ\n\n{file_text}\n\n")
                full_context = "\n".join(context_parts) if context_parts else ""
                instructions = """# ⚠️ ИНСТРУКЦИИ:
1. Используй ТОЛЬКО информацию из источников.
2. НЕ выдумывай факты.
3. Цитируй источники.
4. Соблюдай стиль ветви.
"""
                full_context_with_instructions = full_context + instructions if full_context else instructions
                chatgpt_prompt, claude_prompt, qwen_prompt = get_branch_prompts(branch, task, full_context_with_instructions)  
        try:
            with st.spinner("🤖 ChatGPT пишет..."):
                draft = ask(MODEL_CHATGPT, chatgpt_prompt)
            st.subheader("1. Черновик от ChatGPT", divider="orange")
            st.write(draft)
            with st.spinner("🔍 Claude проверяет..."):
                critique = ask(MODEL_CLAUDE, claude_prompt.replace('{draft}', draft))
            st.subheader("2. Критика от Claude", divider="red")
            st.write(critique)
            with st.spinner("🧠 Qwen готовит финал..."):
                final = ask(MODEL_QWEN, qwen_prompt.replace('{draft}', draft).replace('{critique}', critique))
            st.subheader("3. Финал от Qwen", divider="green")
            st.write(final)
        except Exception as e:
            st.error(f"❌ Ошибка при генерации: {e}")
            st.stop()
                                           # === КНОПКИ СКАЧИВАНИЯ ===
                st.markdown("---")
                st.success("✅ Задача выполнена! Выбери, что скачать:")
                
                # Формируем имя файла из задачи
                safe_task = task[:60].strip()
                # Удаляем запрещённые символы для имён файлов
                for char in ['/', '\\', ':', '*', '?', '"', '<', '>', '|', '\n', '\r']:
                    safe_task = safe_task.replace(char, '')
                safe_task = safe_task.replace('  ', ' ').strip()
                if not safe_task:
                    safe_task = "Документ"
                date_str = datetime.now().strftime('%d%m%Y')
                
                col1, col2 = st.columns(2)
                
                with col1:
                    # Скачивание ТОЛЬКО финального документа
                    final_doc = docx.Document()
                    final_doc.add_heading(task, level=1)
                    final_doc.add_paragraph(final)
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                        final_doc.save(tmp.name); final_path = tmp.name
                    with open(final_path, "rb") as f:
                        st.download_button(
                            label="[DOC] Скачать готовый документ",
                            data=f,
                            file_name=f"{safe_task}_{date_str}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                    os.remove(final_path)
                
                with col2:
                    # Скачивание ПОЛНОЙ беседы
                    full_doc = docx.Document()
                    full_doc.add_heading(f'Задача ({branch_names[branch]}):', level=1)
                    full_doc.add_paragraph(task)
                    if file_name: full_doc.add_heading('Файл:', level=2); full_doc.add_paragraph(file_name)
                    full_doc.add_heading('1. Черновик от ChatGPT', level=2); full_doc.add_paragraph(draft)
                    full_doc.add_heading('2. Критика от Claude', level=2); full_doc.add_paragraph(critique)
                    full_doc.add_heading('3. Финальный результат от Qwen', level=2); full_doc.add_paragraph(final)
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                        full_doc.save(tmp.name); full_path = tmp.name
                    with open(full_path, "rb") as f:
                        st.download_button(
                            label="[КНИГИ] Скачать полную беседу",
                            data=f,
                            file_name=f"Беседа_{safe_task}_{date_str}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                    os.remove(full_path)
                
                # Для бизнес-ветви - дополнительная кнопка Excel
                if branch == 'business':
                    st.markdown("---")
                    st.subheader("[Excel] Скачать в Excel", divider="green")
                    st.caption("Структурированный отчёт с разделами.")
                    
                    excel_wb = openpyxl.Workbook()
                    ws = excel_wb.active
                    ws.title = "Готовый документ"
                    ws['A1'] = "ЗАДАЧА:"
                    ws['A1'].font = Font(bold=True, size=12)
                    ws['A2'] = task
                    ws['A4'] = "ФИНАЛЬНЫЙ РЕЗУЛЬТАТ:"
                    ws['A4'].font = Font(bold=True, size=12, color="FFFFFF")
                    ws['A4'].fill = PatternFill(start_color="27AE60", end_color="27AE60", fill_type="solid")
                    # Разбиваем финал на параграфы
                    final_lines = final.split('\n')
                    row = 5
                    for line in final_lines:
                        ws[f'A{row}'] = line
                        row += 1
                    # Ширина столбца
                    ws.column_dimensions['A'].width = 100
                    
                    # Лист с черновиком
                    ws2 = excel_wb.create_sheet("Черновик ChatGPT")
                    ws2['A1'] = "ЧЕРНОВИК:"
                    ws2['A1'].font = Font(bold=True, size=12)
                    for i, line in enumerate(draft.split('\n'), 2):
                        ws2[f'A{i}'] = line
                    ws2.column_dimensions['A'].width = 100
                    
                    # Лист с критикой
                    ws3 = excel_wb.create_sheet("Критика Claude")
                    ws3['A1'] = "КРИТИКА:"
                    ws3['A1'].font = Font(bold=True, size=12)
                    for i, line in enumerate(critique.split('\n'), 2):
                        ws3[f'A{i}'] = line
                    ws3.column_dimensions['A'].width = 100
                    
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                        excel_wb.save(tmp.name); excel_path = tmp.name
                    with open(excel_path, "rb") as f:
                        st.download_button(
                            label="[Excel] Скачать структурированный Excel",
                            data=f,
                            file_name=f"Бизнес_отчёт_{datetime.now().strftime('%d%m%Y_%H%M')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    os.remove(excel_path)
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
                
                entry_id = str(uuid.uuid4())[:8]
                project_id = selected_project["id"] if selected_project else None
                save_to_history({"id": entry_id, "project_id": project_id, "branch": branch,
                    "date": datetime.now().strftime("%d.%m.%Y %H:%M"),
                    "task": task, "file": file_name,
                    "draft": draft, "critique": critique, "final": final})
                st.sidebar.success("💾 Сохранено!")
                st.session_state['show_new_chat'] = True
            except Exception as e: st.error(f"[X] Ошибка: {str(e)}")
                    # === ЭКСПОРТ В EXCEL (ТОЛЬКО ДЛЯ БИЗНЕС-ВЕТВИ) ===
    if branch == 'business' and final:
        st.markdown("---")
        st.markdown("### [Excel] Экспорт в Excel")
        
        if st.button("📥 Скачать отчет в Excel", use_container_width=True):
            try:
                import openpyxl
                from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
                
                wb = openpyxl.Workbook()
                
                # Лист 1: Резюме
                ws1 = wb.active
                ws1.title = "Резюме"
                ws1['A1'] = "ФИНАНСОВЫЙ ОТЧЕТ"
                ws1['A1'].font = Font(size=16, bold=True, color="FFFFFF")
                ws1['A1'].fill = PatternFill(start_color="27AE60", end_color="27AE60", fill_type="solid")
                ws1['A1'].alignment = Alignment(horizontal="center")
                ws1.merge_cells('A1:B1')
                
                ws1['A3'] = "Задача:"
                ws1['A3'].font = Font(bold=True)
                ws1['B3'] = task
                
                ws1['A5'] = "Резюме:"
                ws1['A5'].font = Font(bold=True)
                ws1['B5'] = final[:500] if final else ""
                
                # Лист 2: Детальный анализ
                ws2 = wb.create_sheet("Анализ")
                ws2['A1'] = "ДЕТАЛЬНЫЙ АНАЛИЗ"
                ws2['A1'].font = Font(size=16, bold=True, color="FFFFFF")
                ws2['A1'].fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
                ws2['A1'].alignment = Alignment(horizontal="center")
                ws2.merge_cells('A1:B1')
                
                ws2['A3'] = "Полный текст анализа:"
                ws2['A3'].font = Font(bold=True)
                ws2['B3'] = final if final else ""
                ws2.column_dimensions['B'].width = 80
                
                # Лист 3: Рекомендации
                ws3 = wb.create_sheet("Рекомендации")
                ws3['A1'] = "РЕКОМЕНДАЦИИ"
                ws3['A1'].font = Font(size=16, bold=True, color="FFFFFF")
                ws3['A1'].fill = PatternFill(start_color="E74C3C", end_color="E74C3C", fill_type="solid")
                ws3['A1'].alignment = Alignment(horizontal="center")
                ws3.merge_cells('A1:B1')
                
                ws3['A3'] = "Ключевые рекомендации:"
                ws3['A3'].font = Font(bold=True)
                ws3['B3'] = final[-500:] if final else ""
                
                # Сохранение
                import tempfile
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    wb.save(tmp.name)
                    tmp_path = tmp.name
                
                with open(tmp_path, "rb") as f:
                    st.download_button(
                        label="⬇️ Скачать Excel",
                        data=f,
                        file_name=f"Финансовый_отчет_{datetime.now().strftime('%d%m%Y')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
                import os
                os.remove(tmp_path)
                st.success("[OK] Excel файл готов!")
                
            except Exception as e:
                st.error(f"Ошибка создания Excel: {e}")
