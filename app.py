import streamlit as st
from openai import OpenAI
import PyPDF2
import docx
import json
import os
from datetime import datetime
import uuid

# ================= НАСТРОЙКИ =================
KEY_FILE = "key.txt"
HISTORY_FILE = "history.json"
PROJECTS_FILE = "projects.json"
BASE_URL = "https://api.vsegpt.ru/v1"
MODEL_CHATGPT = "openai/gpt-4o-mini"
MODEL_CLAUDE  = "anthropic/claude-3-5-haiku"
MODEL_QWEN    = "qwen/qwen-plus"
MAX_FILE_SIZE = 50000 
# =============================================

# --- КАСТОМНЫЙ CSS: ТЕПЛЫЙ И МЯГКИЙ ДИЗАЙН ---
st.markdown("""
<style>
    /* Общий фон приложения - мягкий кремовый */
    .stApp {
        background-color: #FAF8F5;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    
    /* Цвет текста - теплый темно-серый, не резкий черный */
    .stApp p, .stApp div, .stApp label {
        color: #4A403A !important;
    }

    /* Заголовки */
    h1, h2, h3 {
        color: #8B5E3C !important;
        font-weight: 600 !important;
    }
    
    /* Боковая панель */
    section[data-testid="stSidebar"] {
        background-color: #FDFBF9 !important;
        border-right: 1px solid #F0EBE3 !important;
    }
    
    /* Кнопки - мягкий терракотовый/песочный */
    .stButton>button {
        background-color: #D4A373 !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 12px !important;
        padding: 10px 20px !important;
        font-weight: 600 !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 2px 10px rgba(212, 163, 115, 0.2) !important;
    }
    
    .stButton>button:hover {
        background-color: #BC8A5F !important;
        transform: translateY(-1px) !important;
        box-shadow: 0 4px 15px rgba(212, 163, 115, 0.3) !important;
    }
    
    /* Карточки для ответов ИИ */
    .ai-card {
        background: #FFFFFF;
        border-radius: 16px;
        padding: 24px;
        margin: 20px 0;
        box-shadow: 0 4px 24px rgba(139, 94, 60, 0.06);
        border: 1px solid #F0EBE3;
    }
    
    /* Индивидуальные акценты для каждого ИИ */
    .card-chatgpt { border-left: 5px solid #E9C46A; }
    .card-claude { border-left: 5px solid #E07A5F; }
    .card-qwen { border-left: 5px solid #8B5E3C; background-color: #FAF6F1; }

    /* Поле ввода задачи */
    .stTextArea textarea {
        border-radius: 12px !important;
        border: 1px solid #E0D8CC !important;
        background: #FFFFFF !important;
        color: #4A403A !important;
        padding: 15px !important;
    }
    
    /* Смягчаем стандартные уведомления Streamlit */
    .stSuccess {
        background-color: #F1F8E9 !important;
        color: #558B2F !important;
        border-radius: 10px !important;
        border: 1px solid #DCEDC8 !important;
    }
    .stError {
        background-color: #FBE9E7 !important;
        color: #D84315 !important;
        border-radius: 10px !important;
        border: 1px solid #FFCCBC !important;
    }
    .stWarning {
        background-color: #FFF8E1 !important;
        color: #F57F17 !important;
        border-radius: 10px !important;
        border: 1px solid #FFE082 !important;
    }
    .stInfo {
        background-color: #E3F2FD !important;
        color: #1565C0 !important;
        border-radius: 10px !important;
        border: 1px solid #BBDEFB !important;
    }
    
    /* Разделители */
    hr {
        border: none !important;
        height: 1px !important;
        background: #E0D8CC !important;
        margin: 30px 0 !important;
    }
</style>
""", unsafe_allow_html=True)

st.set_page_config(page_title="Совет ИИ", page_icon="🤖", layout="wide")

# Красивый заголовок
st.markdown("""
<div style="text-align: center; margin-bottom: 40px; margin-top: 20px;">
    <h1 style="color: #8B5E3C; font-size: 3rem; margin: 0;">🤖 Совет ИИ</h1>
    <p style="color: #A89F91; font-size: 1.2rem; margin-top: 10px; font-weight: 400;">
        Три разума. Одна задача. Идеальный результат.
    </p>
</div>
""", unsafe_allow_html=True)

# --- ФУНКЦИИ ДЛЯ РАБОТЫ С ПРОЕКТАМИ ---
def load_projects():
    if os.path.exists(PROJECTS_FILE):
        with open(PROJECTS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

def save_projects(projects):
    with open(PROJECTS_FILE, "w", encoding="utf-8") as f:
        json.dump(projects, f, ensure_ascii=False, indent=2)

def create_project(name, description=""):
    projects = load_projects()
    new_project = {
        "id": str(uuid.uuid4())[:8],
        "name": name,
        "description": description,
        "created": datetime.now().strftime("%d.%m.%Y %H:%M")
    }
    projects.append(new_project)
    save_projects(projects)
    return new_project

def delete_project(project_id):
    projects = load_projects()
    projects = [p for p in projects if p["id"] != project_id]
    save_projects(projects)
    history = load_history()
    history = [h for h in history if h.get("project_id") != project_id]
    save_history(history)

# --- ФУНКЦИИ ДЛЯ РАБОТЫ С ИСТОРИЕЙ ---
def load_history():
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

def save_history(history):
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(history, f, ensure_ascii=False, indent=2)

def save_to_history(entry):
    history = load_history()
    history.append(entry)
    save_history(history)

# --- ОСТАЛЬНЫЕ ФУНКЦИИ ---
def load_key():
    try:
        return st.secrets["API_KEY"]
    except Exception:
        pass
    try:
        with open(KEY_FILE, "r", encoding="utf-8") as f:
            return f.read().strip()
    except FileNotFoundError:
        return None

def extract_text(file):
    filename = file.name.lower()
    if filename.endswith('.txt'):
        return file.read().decode('utf-8')
    elif filename.endswith('.pdf'):
        pdf_reader = PyPDF2.PdfReader(file)
        text = ""
        for page in pdf_reader.pages:
            text += page.extract_text() + "\n"
        return text
    elif filename.endswith('.docx'):
        doc = docx.Document(file
