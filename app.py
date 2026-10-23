import streamlit as st
from openai import OpenAI
import PyPDF2
import docx
import json
import os
import tempfile
import base64
import requests
from datetime import datetime
import uuid
from io import BytesIO
from pathlib import Path
import re

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
MAX_BOOK_PAGES = 30  # Максимум страниц из книги для анализа
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
</style>
""", unsafe_allow_html=True)

st.set_page_config(page_title="Совет ИИ", page_icon="🤖", layout="wide")

st.markdown("""
<div style="text-align: center; margin-bottom: 40px; margin-top: 20px;">
    <h1 style="color: #8B5E3C; font-size: 3rem; margin: 0;">🤖 Совет ИИ</h1>
    <p style="color: #A89F91; font-size: 1.2rem; margin-top: 10px; font-weight: 400;">
        Три разума. Одна задача. Идеальный результат.
    </p>
</div>
""", unsafe_allow_html=True)

# --- ФУНКЦИИ ДЛЯ БИБЛИОТЕКИ ---
def scan_library_folder(folder_path):
    """Сканирует папку и создаёт индекс книг"""
    books = []
    folder = Path(folder_path)
    
    if not folder.exists():
        return {'error': f'Папка не найдена: {folder_path}'}
    
    for file_path in folder.iterdir():
        if file_path.is_file() and file_path.suffix.lower() in ['.pdf', '.docx', '.txt']:
            try:
                metadata = extract_book_metadata(file_path)
                books.append({
                    'path': str(file_path),
                    'filename': file_path.name,
                    'title': metadata.get('title', file_path.stem),
                    'author': metadata.get('author', 'Автор неизвестен'),
                    'year': metadata.get('year', 'Год неизвестен'),
                    'size_pages': metadata.get('pages', 0),
                    'extension': file_path.suffix.lower()
                })
            except Exception as e:
                books.append({
                    'path': str(file_path),
                    'filename': file_path.name,
                    'title': file_path.stem,
                    'author': 'Автор неизвестен',
                    'year': 'Год неизвестен',
                    'size_pages': 0,
                    'extension': file_path.suffix.lower(),
                    'error': str(e)
                })
    
    return books

def extract_book_metadata(file_path):
    """Извлекает метаданные из файла"""
    metadata = {}
    
    if file_path.suffix.lower() == '.pdf':
        try:
            with open(file_path, 'rb') as f:
                pdf_reader = PyPDF2.PdfReader(f)
                info = pdf_reader.metadata
                if info:
                    metadata['title'] = info.get('/Title', '')
                    metadata['author'] = info.get('/Author', '')
                    metadata['year'] = info.get('/CreationDate', '')[:4] if info.get('/CreationDate') else ''
                metadata['pages'] = len(pdf_reader.pages)
        except:
            metadata['pages'] = 0
    
    elif file_path.suffix.lower() == '.docx':
        try:
            doc = docx.Document(file_path)
            metadata['pages'] = len(doc.paragraphs) // 50  # Примерная оценка
        except:
            metadata['pages'] = 0
    
    return metadata

def search_relevant_books(query, books, top_n=3):
    """Находит релевантные книги по запросу"""
    if not books:
        return []
    
    query_words = set(query.lower().split())
    scored_books = []
    
    for book in books:
        score = 0
        # Проверяем совпадения в названии
        title_words = set(book['title'].lower().split())
        score += len(query_words & title_words) * 3
        
        # Проверяем совпадения в имени автора
        author_words = set(book['author'].lower().split())
        score += len(query_words & author_words) * 2
        
        # Проверяем совпадения в имени файла
        filename_words = set(book['filename'].lower().split())
        score += len(query_words & filename_words)
        
        if score > 0:
            scored_books.append((score, book))
    
    # Сортируем по релевантности
    scored_books.sort(reverse=True, key=lambda x: x[0])
    
    # Возвращаем топ-N
    return [book for score, book in scored_books[:top_n]]

def load_book_content(file_path, max_pages=MAX_BOOK_PAGES):
    """Загружает текст из книги (первые N страниц)"""
    file_path = Path(file_path)
    
    if not file_path.exists():
        return f"Файл не найден: {file_path}"
    
    try:
        if file_path.suffix.lower() == '.pdf':
            with open(file_path, 'rb') as f:
                pdf_reader = PyPDF2.PdfReader(f)
                text = ""
                for i, page in enumerate(pdf_reader.pages[:max_pages]):
                    text += page.extract_text() + "\n"
                return text
        
        elif file_path.suffix.lower() == '.docx':
            doc = docx.Document(file_path)
            paragraphs = [p.text for p in doc.paragraphs[:max_pages * 3]]  # Примерно
            return "\n".join(paragraphs)
        
        elif file_path.suffix.lower() == '.txt':
            with open(file_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()[:max_pages * 50]  # Примерно
                return "".join(lines)
        
        return "Неподдерживаемый формат"
    
    except Exception as e:
        return f"Ошибка чтения файла: {str(e)}"

def save_library_index(books):
    """Сохраняет индекс библиотеки"""
    with open(LIBRARY_INDEX_FILE, 'w', encoding='utf-8') as f:
        json.dump(books, f, ensure_ascii=False, indent=2)

def load_library_index():
    """Загружает индекс библиотеки"""
    if os.path.exists(LIBRARY_INDEX_FILE):
        with open(LIBRARY_INDEX_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return []

# --- ФУНКЦИИ ДЛЯ НАУЧНОГО ПОИСКА ---
def search_openalex(query, max_results=15):
    """Поиск статей в OpenAlex"""
    try:
        url = f"https://api.openalex.org/works?search={query}&per-page={max_results}&sort=relevance_score:desc&filter=open_access.is_oa:true"
        response = requests.get(url, timeout=15)
        response.raise_for_status()
        data = response.json()
        
        articles = []
        for work in data.get('results', []):
            title = work.get('title', 'Без названия')
            abstract = work.get('abstract_inverted_index')
            abstract_text = ""
            if abstract:
                words = {}
                for word, positions in abstract.items():
                    for pos in positions:
                        words[pos] = word
                abstract_text = ' '.join([words[i] for i in sorted(words.keys())])
            
            open_access = work.get('open_access', {})
            pdf_url = open_access.get('oa_url', '')
            
            authors = [authorship.get('author', {}).get('display_name', '') 
                      for authorship in work.get('authorships', [])[:3]]
            
            if pdf_url:
                articles.append({
                    'title': title,
                    'abstract': abstract_text[:500] if abstract_text else 'Аннотация недоступна',
                    'authors': ', '.join(authors) if authors else 'Авторы неизвестны',
                    'year': work.get('publication_year', 'Год неизвестен'),
                    'pdf_url': pdf_url,
                    'source': 'OpenAlex',
                    'has_full_text': True
                })
        return articles
    except Exception as e:
        return [{'error': f'Ошибка поиска в OpenAlex: {str(e)}'}]

def search_cyberleninka(query, max_results=15):
    """Поиск статей в КиберЛенинке"""
    try:
        url = f"https://cyberleninka.ru/api/n/search?q={query}&per_page={max_results}"
        response = requests.get(url, timeout=15)
        response.raise_for_status()
        data = response.json()
        
        articles = []
        for article in data.get('results', []):
            pdf_url = article.get('pdf', '')
            if pdf_url:
                articles.append({
                    'title': article.get('title', 'Без названия'),
                    'abstract': article.get('abstract', 'Аннотация недоступна')[:500],
                    'authors': article.get('authors', 'Авторы неизвестны'),
                    'year': article.get('year', 'Год неизвестен'),
                    'pdf_url': pdf_url,
                    'source': 'КиберЛенинка',
                    'has_full_text': True
                })
        return articles
    except Exception as e:
        return [{'error': f'Ошибка поиска в КиберЛенинке: {str(e)}'}]

def download_and_extract_pdf(pdf_url, max_chars=15000):
    """Скачивает PDF и извлекает текст"""
    try:
        response = requests.get(pdf_url, timeout=30)
        response.raise_for_status()
        
        pdf_file = BytesIO(response.content)
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        
        text = ""
        for page in pdf_reader.pages:
            text += page.extract_text() + "\n"
        
        if len(text) > max_chars:
            text = text[:max_chars]
        
        return text
    except Exception as e:
        return f"Не удалось скачать или прочитать PDF: {str(e)}"

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

def load_history():
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, "r", encoding="utf-8") as f: return json.load(f)
    return []

def save_history(history):
    with open(HISTORY_FILE, "w", encoding="utf-8") as f: json.dump(history, f, ensure_ascii=False, indent=2)

def save_to_history(entry):
    history = load_history()
    history.append(entry)
    save_history(history)

def load_key():
    try: return st.secrets["API_KEY"]
    except Exception: pass
    try:
        with open(KEY_FILE, "r", encoding="utf-8") as f: return f.read().strip()
    except FileNotFoundError: return None

def extract_text(file):
    filename = file.name.lower()
    if filename.endswith('.txt'): return file.read().decode('utf-8')
    elif filename.endswith('.pdf'):
        pdf_reader = PyPDF2.PdfReader(file)
        return "".join([page.extract_text() + "\n" for page in pdf_reader.pages])
    elif filename.endswith('.docx'):
        doc = docx.Document(file)
        return "".join([p.text + "\n" for p in doc.paragraphs])
    else: raise Exception("Неподдерживаемый формат.")

def ask(model, prompt):
    key = load_key()
    if not key: raise Exception("Не найден файл key.txt или секрет API_KEY")
    client = OpenAI(api_key=key, base_url=BASE_URL)
    return client.chat.completions.create(model=model, messages=[{"role": "user", "content": prompt}], temperature=0.3).choices[0].message.content

def analyze_image(file):
    """Распознает текст и описывает содержимое изображения через GPT-4o"""
    key = load_key()
    if not key: raise Exception("Не найден файл key.txt или секрет API_KEY")
    
    image_bytes = file.read()
    base64_image = base64.b64encode(image_bytes).decode('utf-8')
    filename = file.name.lower()
    mime = "image/jpeg" if filename.endswith(('.jpg', '.jpeg')) else "image/png"
    
    client = OpenAI(api_key=key, base_url=BASE_URL)
    response = client.chat.completions.create(
        model=MODEL_CHATGPT,
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": "Ты — система оптического распознавания символов (OCR) и визуальный аналитик. Твоя задача: 1. Распознать ВЕСЬ текст на изображении и вывести его дословно. 2. Описать визуальные элементы (таблицы, схемы, подписи, рукописные пометки, печати), если они есть. Верни результат в виде структурированного текста. Если текста нет, просто опиши, что на картинке."},
                    {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{base64_image}"}}
                ]
            }
        ],
        max_tokens=2000
    )
    return response.choices[0].message.content

# --- ИНТЕРФЕЙС ---
st.sidebar.header("📁 Проекты")
projects = load_projects()

with st.sidebar.expander(" Новый проект"):
    new_project_name = st.text_input("Название проекта:", placeholder="Например: Юридическая практика")
    new_project_desc = st.text_area("Описание (необязательно):", height=80)
    if st.button("Создать проект"):
        if new_project_name.strip():
            create_project(new_project_name.strip(), new_project_desc.strip())
            st.success(f"Проект '{new_project_name}' создан!")
            st.rerun()
        else: st.warning("Введите название проекта")

if projects:
    project_names = [f"{p['name']} ({p['created']})" for p in projects]
    selected_name = st.sidebar.selectbox("Выбери проект:", project_names)
    selected_project = projects[project_names.index(selected_name)]
    if st.sidebar.button("🗑 Удалить проект"):
        if st.sidebar.checkbox("Да, удалить проект и все его задачи"):
            delete_project(selected_project["id"])
            st.sidebar.success("Проект удалён")
            st.rerun()
else:
    st.sidebar.info("Нет проектов. Создай первый!")
    selected_project = None

st.sidebar.markdown("---")

if selected_project:
    st.sidebar.header(f"📂 История: {selected_project['name']}")
    history = load_history()
    project_history = [h for h in history if h.get("project_id") == selected_project["id"]]
    if project_history:
        st.sidebar.markdown("---")
        for i, entry in enumerate(reversed(project_history)):
            short_task = entry["task"][:40] + "..." if len(entry["task"]) > 40 else entry["task"]
            with st.sidebar.expander(f"📝 {short_task}\n🕒 {entry['date']}"):
                st.write("**Задача:**"); st.write(entry["task"])
                if entry.get("file"): st.write(f"**Файл:** {entry['file']}")
                if entry.get("external_ai"): st.write(f"**Внешний ИИ:** {entry['external_ai'][:100]}...")
                st.write("**Финальный результат:**"); st.success(entry["final"])
    else: st.sidebar.info("В этом проекте пока нет задач")
else:
    st.sidebar.header("📂 История задач"); st.sidebar.info("Выбери проект слева")

st.sidebar.markdown("---")
if load_key(): st.sidebar.success("🔑 Ключ активен")
else: st.sidebar.error(" Ключ не найден")

if selected_project:
    st.markdown(f"""<div class="ai-card"><h3 style="margin: 0; color: #8B5E3C;">📁 Проект: {selected_project['name']}</h3>{f'<p style="margin: 10px 0 0 0; color: #A89F91;">{selected_project["description"]}</p>' if selected_project.get("description") else ''}</div>""", unsafe_allow_html=True)
else:
    st.markdown("""<div class="ai-card" style="text-align: center;"><h3 style="margin: 0; color: #8B5E3C;">👈 Создай или выбери проект в боковой панели</h3></div>""", unsafe_allow_html=True)

st.markdown("---")
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    if st.button("🔌 Проверить связь", use_container_width=True):
        for name, model in [("ChatGPT", MODEL_CHATGPT), ("Claude", MODEL_CLAUDE), ("Qwen", MODEL_QWEN)]:
            try:
                ask(model, "ОК"); st.success(f"✅ {name} — на связи")
            except Exception as e: st.error(f"❌ {name} — ошибка: {e}")

st.markdown("---")

# === МОЯ БИБЛИОТЕКА ===
st.markdown("""<div class="ai-card card-library"><h3 style="margin: 0 0 10px 0; color: #9B59B6;">📚 Моя научная библиотека</h3><p style="margin: 0; color: #A89F91;">💡 Укажи путь к папке с учебниками и монографиями. Приложение найдёт релевантные книги и использует их для написания работы.</p></div>""", unsafe_allow_html=True)

library_path = st.text_input("Путь к папке с библиотекой:", placeholder="Например: C:\\Users\\123\\Documents\\Научная_библиотека")

col1, col2 = st.columns(2)
with col1:
    if st.button("📂 Сканировать библиотеку", use_container_width=True):
        if library_path.strip():
            with st.spinner("Сканирую папку..."):
                books = scan_library_folder(library_path)
                if isinstance(books, dict) and 'error' in books:
                    st.error(books['error'])
                else:
                    save_library_index(books)
                    st.success(f"✅ Найдено {len(books)} книг!")
                    st.session_state['library_books'] = books
        else:
            st.warning("Введи путь к папке")

with col2:
    use_library = st.checkbox("📖 Использовать мою библиотеку", value=False)

# Отображаем список книг
if 'library_books' in st.session_state and st.session_state['library_books']:
    books = st.session_state['library_books']
    st.markdown(f"### 📚 Найдено книг: {len(books)}")
    
    for i, book in enumerate(books[:10], 1):  # Показываем первые 10
        with st.container():
            st.markdown(f"""
            <div class="book-item">
                <h4 style="margin: 0 0 8px 0; color: #9B59B6;">{i}. {book['title']}</h4>
                <p style="margin: 0 0 8px 0; color: #666; font-size: 0.9em;">
                    <b>Автор:</b> {book['author']} | <b>Год:</b> {book['year']} | <b>Страниц:</b> {book['size_pages']}
                </p>
                <p style="margin: 0; color: #4A403A; font-size: 0.85em;">
                    📄 {book['filename']}
                </p>
            </div>
            """, unsafe_allow_html=True)

st.markdown("---")

# === НАУЧНЫЙ ПОИСК ===
st.markdown("""<div class="ai-card card-search"><h3 style="margin: 0 0 10px 0; color: #5B8DB8;">🔬 Научный поиск в библиотеках</h3><p style="margin: 0; color: #A89F91;">💡 Найди научные статьи по теме и добавь их в контекст.</p></div>""", unsafe_allow_html=True)

search_query = st.text_input("Введи тему для поиска:", placeholder="Например: искусственный интеллект в гражданском праве")

col1, col2 = st.columns(2)
with col1:
    use_openalex = st.checkbox("OpenAlex (мировая база)", value=True)
with col2:
    use_cyberleninka = st.checkbox("КиберЛенинка (российская база)", value=True)

deep_analysis = st.checkbox("🎓 Глубокое исследование (скачать полные тексты)", value=True)

if st.button("🔍 Найти статьи", use_container_width=True):
    if not search_query.strip():
        st.warning("Введи тему для поиска")
    else:
        all_articles = []
        
        if use_openalex:
            with st.spinner("🔍 Ищу в OpenAlex..."):
                articles = search_openalex(search_query)
                all_articles.extend(articles)
        
        if use_cyberleninka:
            with st.spinner("🔍 Ищу в КиберЛенинке..."):
                articles = search_cyberleninka(search_query)
                all_articles.extend(articles)
        
        full_text_articles = [a for a in all_articles if a.get('has_full_text', False)]
        
        if deep_analysis and full_text_articles:
            articles_to_analyze = full_text_articles[:MAX_ARTICLES]
            
            with st.spinner(f"📥 Скачиваю полные тексты {len(articles_to_analyze)} статей..."):
                articles_with_text = []
                for i, article in enumerate(articles_to_analyze):
                    with st.status(f"Скачиваю статью {i+1}/{len(articles_to_analyze)}: {article['title'][:50]}..."):
                        text = download_and_extract_pdf(article['pdf_url'])
                        article['full_text'] = text
                        articles_with_text.append(article)
            
            st.session_state['deep_analysis_articles'] = articles_with_text
            st.success(f"✅ Скачано {len(articles_with_text)} статей!")
        else:
            st.session_state['search_results'] = all_articles
            st.success(f"Найдено статей: {len(all_articles)}")

if 'deep_analysis_articles' in st.session_state and st.session_state['deep_analysis_articles']:
    st.markdown("### 📚 Статьи для глубокого анализа")
    for i, article in enumerate(st.session_state['deep_analysis_articles'], 1):
        st.markdown(f"""
        <div class="article-item">
            <h4 style="margin: 0 0 8px 0; color: #5B8DB8;">{i}. {article['title']}</h4>
            <p style="margin: 0; color: #666; font-size: 0.9em;">
                <b>Авторы:</b> {article['authors']} | <b>Год:</b> {article['year']} | <b>Источник:</b> {article['source']}
            </p>
        </div>
        """, unsafe_allow_html=True)

elif 'search_results' in st.session_state and st.session_state['search_results']:
    st.markdown("### 📚 Результаты поиска")
    selected_articles = []
    for i, article in enumerate(st.session_state['search_results']):
        if 'error' in article:
            st.error(article['error'])
            continue
        
        st.markdown(f"""
        <div class="article-item">
            <h4 style="margin: 0 0 8px 0; color: #5B8DB8;">{article['title']}</h4>
            <p style="margin: 0; color: #666; font-size: 0.9em;">
                <b>Авторы:</b> {article['authors']} | <b>Год:</b> {article['year']}
            </p>
            <p style="margin: 0; color: #4A403A; font-size: 0.95em;">{article['abstract']}</p>
        </div>
        """, unsafe_allow_html=True)
        
        if st.checkbox(f"Добавить: {article['title'][:60]}...", key=f"article_{i}"):
            selected_articles.append(article)
    
    if selected_articles:
        if st.button("📥 Добавить выбранные статьи", use_container_width=True):
            articles_text = "#  НАУЧНЫЕ СТАТЬИ\n\n"
            for i, article in enumerate(selected_articles, 1):
                articles_text += f"## [{i}] {article['title']}\n**Авторы:** {article['authors']} | **Год:** {article['year']}\n\n{article['abstract']}\n\n---\n\n"
            st.session_state['articles_context'] = articles_text
            st.success(f"✅ Добавлено {len(selected_articles)} статей!")
            st.rerun()

st.markdown("---")

# === ЗАГРУЗКА ФАЙЛОВ ===
st.markdown("""<div class="ai-card"><h3 style="margin: 0 0 10px 0; color: #8B5E3C;">📎 Загрузи материалы (необязательно)</h3><p style="margin: 0; color: #A89F91;"> PDF, DOCX, TXT, JPG, PNG</p></div>""", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Выбери файл:", type=['pdf', 'docx', 'txt', 'jpg', 'jpeg', 'png'])
file_text = ""; file_name = ""
if uploaded_file is not None:
    try:
        file_name = uploaded_file.name
        filename_lower = file_name.lower()
        
        if filename_lower.endswith(('.jpg', '.jpeg', '.png')):
            with st.spinner("👁️ Анализирую изображение..."):
                file_text = analyze_image(uploaded_file)
            st.success(f"✅ Изображение проанализировано: {file_name}")
        else:
            file_text = extract_text(uploaded_file)
            if len(file_text) > MAX_FILE_SIZE:
                file_text = file_text[:MAX_FILE_SIZE]
            st.success(f"✅ Файл загружен: {file_name}")
            
    except Exception as e: st.error(f"❌ Ошибка: {e}")

st.markdown("---")

# === МНЕНИЕ ВНЕШНЕГО ИИ ===
external_ai_opinion = st.text_area(
    "⚖️ Мнение внешнего ИИ (Cognevia, ПравоТех и т.д.) - необязательно:",
    height=100,
    placeholder="Вставь ответ от внешнего ИИ..."
)

st.markdown("---")
task = st.text_area("Опиши задачу:", height=150, placeholder="Например: 'На основе этих источников напиши научную статью...'")

col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    if st.button("🚀 Запустить обсуждение", use_container_width=True):
        if not selected_project: st.warning("⚠️ Выбери проект")
        elif not task.strip(): st.warning("️ Введи задачу")
        else:
            try:
                context_parts = []
                
                # 1. Добавляем книги из библиотеки
                if use_library and 'library_books' in st.session_state and st.session_state['library_books']:
                    with st.spinner("📚 Ищу релевантные книги в библиотеке..."):
                        relevant_books = search_relevant_books(task, st.session_state['library_books'], top_n=3)
                        
                        if relevant_books:
                            books_text = "# 📚 МОЯ НАУЧНАЯ БИБЛИОТЕКА (РЕЛЕВАНТНЫЕ КНИГИ)\n\n"
                            books_text += "ИНСТРУКЦИЯ: Цитируй эти книги в формате [К1], [К2] и т.д. В конце добавь 'СПИСОК ЛИТЕРАТУРЫ'.\n\n"
                            
                            for i, book in enumerate(relevant_books, 1):
                                with st.status(f"Загружаю книгу: {book['title'][:50]}..."):
                                    book_text = load_book_content(book['path'])
                                    books_text += f"## [К{i}] {book['title']}\n"
                                    books_text += f"**Автор:** {book['author']} | **Год:** {book['year']}\n\n"
                                    books_text += f"**Текст книги:**\n{book_text[:10000]}\n\n---\n\n"
                            
                            context_parts.append(books_text)
                            st.success(f"✅ Найдено {len(relevant_books)} релевантных книг")
                
                # 2. Добавляем статьи из глубокого анализа
                if 'deep_analysis_articles' in st.session_state and st.session_state['deep_analysis_articles']:
                    articles_text = "# 📚 НАУЧНЫЕ СТАТЬИ (ПОЛНЫЕ ТЕКСТЫ)\n\n"
                    articles_text += "ИНСТРУКЦИЯ: Цитируй в формате [1], [2] и т.д.\n\n"
                    
                    for i, article in enumerate(st.session_state['deep_analysis_articles'], 1):
                        articles_text += f"## [{i}] {article['title']}\n"
                        articles_text += f"**Авторы:** {article['authors']} | **Год:** {article['year']}\n\n"
                        articles_text += f"**Текст:**\n{article.get('full_text', '')[:10000]}\n\n---\n\n"
                    
                    context_parts.append(articles_text)
                
                # 3. Добавляем статьи из обычного поиска
                if 'articles_context' in st.session_state and st.session_state['articles_context']:
                    context_parts.append(st.session_state['articles_context'])
                
                # 4. Добавляем загруженный файл
                if file_text:
                    context_parts.append(f"# 📄 ЗАГРУЖЕННЫЙ ФАЙЛ\n\n{file_text}\n\n")
                
                # 5. Добавляем мнение внешнего ИИ
                if external_ai_opinion.strip():
                    context_parts.append(f"# ️ МНЕНИЕ ВНЕШНЕГО ИИ\n\n{external_ai_opinion}\n\n")
                
                full_context = "\n".join(context_parts) if context_parts else ""
                
                instructions = """# ⚠️ ВАЖНЫЕ ИНСТРУКЦИИ:
1. Работай с НАУЧНЫМ материалом. Точность критична.
2. Используй ТОЛЬКО информацию из источников.
3. НЕ выдумывай факты.
4. ЦИТИРУЙ источники: книги как [К1], [К2], статьи как [1], [2].
5. В конце добавь 'СПИСОК ЛИТЕРАТУРЫ'.

"""
                
                full_context_with_instructions = full_context + instructions if full_context else instructions

                with st.spinner("📝 ChatGPT пишет черновик..."):
                    draft = ask(MODEL_CHATGPT, f"""Ты — научный ассистент. Напиши ЧЕРНОВИК НАУЧНОЙ СТАТЬИ.

{full_context_with_instructions}
## ЗАДАЧА:
{task}

## СТРУКТУРА:
- Введение
- Основная часть
- Выводы
- СПИСОК ЛИТЕРАТУРЫ

Цитируй источники.""")
                
                st.markdown(f"""<div class="ai-card card-chatgpt"><h3 style="margin: 0 0 15px 0; color: #D4A373;">1️⃣ Черновик от ChatGPT</h3></div>""", unsafe_allow_html=True)
                st.write(draft)

                with st.spinner("🔍 Claude проверяет..."):
                    critique = ask(MODEL_CLAUDE, f"""Ты — научный рецензент. Проверь черновик.

{full_context_with_instructions}
## ЗАДАЧА:
{task}

## ЧЕРНОВИК:
{draft}

## ПРОВЕРЬ:
1. Все ли факты есть в источниках?
2. Корректны ли цитаты?
3. Есть ли список литературы?
4. Что упущено?

Будь строгим.""")
                
                st.markdown(f"""<div class="ai-card card-claude"><h3 style="margin: 0 0 15px 0; color: #E07A5F;">2️⃣ Критика от Claude</h3></div>""", unsafe_allow_html=True)
                st.write(critique)

                with st.spinner("🎯 Qwen готовит финал..."):
                    final = ask(MODEL_QWEN, f"""Ты — главный редактор. Создай ФИНАЛЬНУЮ ВЕРСИЮ.

{full_context_with_instructions}
## ЗАДАЧА:
{task}

## ЧЕРНОВИК:
{draft}

## КРИТИКА:
{critique}

## ТРЕБОВАНИЯ:
1. Исправь ошибки Claude
2. Добавь цитаты [К1], [1] и т.д.
3. Структура: Введение, Основная часть, Выводы, СПИСОК ЛИТЕРАТУРЫ
4. Готовая научная статья""")
                
                st.markdown(f"""<div class="ai-card card-qwen"><h3 style="margin: 0 0 15px 0; color: #8B5E3C;">3️⃣ ✅ Финал от Qwen</h3></div>""", unsafe_allow_html=True)
                st.write(final)

                # Word документ
                doc = docx.Document()
                doc.add_heading('Задача:', level=1)
                doc.add_paragraph(task)
                
                if use_library and 'library_books' in st.session_state:
                    doc.add_heading('Использованные книги из библиотеки:', level=2)
                    relevant_books = search_relevant_books(task, st.session_state['library_books'], top_n=3)
                    for i, book in enumerate(relevant_books, 1):
                        doc.add_paragraph(f"[К{i}] {book['author']}. {book['title']}. {book['year']}.")
                
                if 'deep_analysis_articles' in st.session_state:
                    doc.add_heading('Научные статьи:', level=2)
                    for i, article in enumerate(st.session_state['deep_analysis_articles'], 1):
                        doc.add_paragraph(f"[{i}] {article['authors']}. {article['title']}. {article['year']}.")
                
                doc.add_heading('1. Черновик от ChatGPT', level=2)
                doc.add_paragraph(draft)
                doc.add_heading('2. Критика от Claude', level=2)
                doc.add_paragraph(critique)
                doc.add_heading('3. Финальный результат от Qwen', level=2)
                doc.add_paragraph(final)

                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                    doc.save(tmp.name)
                    tmp_path = tmp.name
                
                with open(tmp_path, "rb") as f:
                    st.download_button(
                        label="📥 Скачать в Word",
                        data=f,
                        file_name=f"Научная_статья_{datetime.now().strftime('%d%m%Y_%H%M')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                os.remove(tmp_path)

                entry = {
                    "project_id": selected_project["id"],
                    "date": datetime.now().strftime("%d.%m.%Y %H:%M"),
                    "task": task,
                    "file": file_name,
                    "external_ai": external_ai_opinion[:200] if external_ai_opinion.strip() else "",
                    "library_books_used": len(search_relevant_books(task, st.session_state.get('library_books', []), top_n=3)) if use_library else 0,
                    "articles_count": len(st.session_state.get('deep_analysis_articles', [])),
                    "draft": draft,
                    "critique": critique,
                    "final": final
                }
                save_to_history(entry)
                st.sidebar.success("💾 Сохранено!")

            except Exception as e: st.error(f"❌ Ошибка: {str(e)}")
