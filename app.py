import streamlit as st
from openai import OpenAI
import PyPDF2
import docx
import json
import os
import tempfile
import base64
from datetime import datetime
import uuid

# ================= НАСТРОЙКИ =================
KEY_FILE = "key.txt"
HISTORY_FILE = "history.json"
PROJECTS_FILE = "projects.json"
BASE_URL = "https://api.vsegpt.ru/v1"
MODEL_CHATGPT = "openai/gpt-4o-mini" # Для картинок лучше использовать gpt-4o, если есть в агрегаторе, но и mini справится
MODEL_CLAUDE  = "anthropic/claude-3-5-haiku"
MODEL_QWEN    = "qwen/qwen-plus"
MAX_FILE_SIZE = 50000 
# =============================================

# --- КАСТОМНЫЙ CSS: ТЕПЛЫЙ И МЯГКИЙ ДИЗАЙН ---
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

    .stTextArea textarea { border-radius: 12px !important; border: 1px solid #E0D8CC !important; background: #FFFFFF !important; color: #4A403A !important; padding: 15px !important; }
    
    .stSuccess { background-color: #F1F8E9 !important; color: #558B2F !important; border-radius: 10px !important; border: 1px solid #DCEDC8 !important; }
    .stError { background-color: #FBE9E7 !important; color: #D84315 !important; border-radius: 10px !important; border: 1px solid #FFCCBC !important; }
    .stWarning { background-color: #FFF8E1 !important; color: #F57F17 !important; border-radius: 10px !important; border: 1px solid #FFE082 !important; }
    .stInfo { background-color: #E3F2FD !important; color: #1565C0 !important; border-radius: 10px !important; border: 1px solid #BBDEFB !important; }
    
    hr { border: none !important; height: 1px !important; background: #E0D8CC !important; margin: 30px 0 !important; }
</style>
""", unsafe_allow_html=True)

st.set_page_config(page_title="Совет ИИ", page_icon="", layout="wide")

st.markdown("""
<div style="text-align: center; margin-bottom: 40px; margin-top: 20px;">
    <h1 style="color: #8B5E3C; font-size: 3rem; margin: 0;">🤖 Совет ИИ</h1>
    <p style="color: #A89F91; font-size: 1.2rem; margin-top: 10px; font-weight: 400;">
        Три разума. Одна задача. Идеальный результат.
    </p>
</div>
""", unsafe_allow_html=True)

# --- ФУНКЦИИ ---
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
        model=MODEL_CHATGPT, # Используем GPT-4o-mini (или gpt-4o, если доступен) для зрения
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

with st.sidebar.expander("➕ Новый проект"):
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
            with st.sidebar.expander(f"📝 {short_task}\n {entry['date']}"):
                st.write("**Задача:**"); st.write(entry["task"])
                if entry.get("file"): st.write(f"**Файл:** {entry['file']}")
                if entry.get("external_ai"): st.write(f"**Внешний ИИ:** {entry['external_ai'][:100]}...")
                st.write("**Финальный результат:**"); st.success(entry["final"])
    else: st.sidebar.info("В этом проекте пока нет задач")
else:
    st.sidebar.header(" История задач"); st.sidebar.info("Выбери проект слева")

st.sidebar.markdown("---")
if load_key(): st.sidebar.success("🔑 Ключ активен")
else: st.sidebar.error("❌ Ключ не найден")

if selected_project:
    st.markdown(f"""<div class="ai-card"><h3 style="margin: 0; color: #8B5E3C;">📁 Проект: {selected_project['name']}</h3>{f'<p style="margin: 10px 0 0 0; color: #A89F91;">{selected_project["description"]}</p>' if selected_project.get("description") else ''}</div>""", unsafe_allow_html=True)
else:
    st.markdown("""<div class="ai-card" style="text-align: center;"><h3 style="margin: 0; color: #8B5E3C;"> Создай или выбери проект в боковой панели</h3></div>""", unsafe_allow_html=True)

st.markdown("---")
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    if st.button("🔌 Проверить связь", use_container_width=True):
        for name, model in [("ChatGPT", MODEL_CHATGPT), ("Claude", MODEL_CLAUDE), ("Qwen", MODEL_QWEN)]:
            try:
                ask(model, "ОК"); st.success(f"✅ {name} — на связи")
            except Exception as e: st.error(f"❌ {name} — ошибка: {e}")

st.markdown("---")
st.markdown("""<div class="ai-card"><h3 style="margin: 0 0 10px 0; color: #8B5E3C;">📎 Загрузи материалы (необязательно)</h3><p style="margin: 0; color: #A89F91;">💡 Поддерживаются PDF, DOCX, TXT, а также <b>фотографии и сканы (JPG, PNG)</b>. ИИ сам распознает текст на фото.</p></div>""", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Выбери PDF, DOCX, TXT или фото (JPG, PNG):", type=['pdf', 'docx', 'txt', 'jpg', 'jpeg', 'png'])
file_text = ""; file_name = ""
if uploaded_file is not None:
    try:
        file_name = uploaded_file.name
        filename_lower = file_name.lower()
        
        if filename_lower.endswith(('.jpg', '.jpeg', '.png')):
            with st.spinner("👁️ Анализирую изображение (распознаю текст и схемы)..."):
                file_text = analyze_image(uploaded_file)
            st.success(f"✅ Изображение проанализировано: {file_name}")
            with st.expander("👁️ Посмотреть распознанный текст и описание"):
                st.text(file_text[:2000] + "..." if len(file_text) > 2000 else file_text)
        else:
            file_text = extract_text(uploaded_file)
            if len(file_text) > MAX_FILE_SIZE:
                file_text = file_text[:MAX_FILE_SIZE]; st.warning(f"⚠️ Файл большой. Взяты только первые {MAX_FILE_SIZE} символов.")
            st.success(f"✅ Файл загружен: {file_name} ({len(file_text)} символов)")
            with st.expander("👁️ Посмотреть загруженный текст"): st.text(file_text[:2000] + "..." if len(file_text) > 2000 else file_text)
            
    except Exception as e: st.error(f"❌ Ошибка чтения файла: {e}")

st.markdown("---")

# === ПОЛЕ: МНЕНИЕ ВНЕШНЕГО ИИ ===
st.markdown("""<div class="ai-card card-external"><h3 style="margin: 0 0 10px 0; color: #6B9080;">⚖️ Мнение внешнего юридического ИИ (необязательно)</h3><p style="margin: 0; color: #A89F91;">💡 Вставь сюда ответ от Cognevia, ПравоТех, Doczilla или другого специализированного ИИ.</p></div>""", unsafe_allow_html=True)

external_ai_opinion = st.text_area(
    "Вставь ответ от внешнего ИИ (Cognevia, ПравоТех, Doczilla и т.д.):",
    height=150,
    placeholder="Например: скопируй сюда анализ судебного дела или правовую позицию..."
)

st.markdown("---")
task = st.text_area("Опиши задачу:", height=150, placeholder="Например: 'На основе этого текста и мнения Cognevia подготовь отзыв на исковое заявление'")

col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    if st.button("🚀 Запустить обсуждение", use_container_width=True):
        if not selected_project: st.warning("⚠️ Сначала создай или выбери проект в боковой панели.")
        elif not task.strip(): st.warning("⚠️ Сначала введи задачу.")
        else:
            try:
                if file_text:
                    context_block = f"# 📖 ИСТОЧНИК (материал)\n\n{file_text}\n\n"
                else:
                    context_block = ""
                
                if external_ai_opinion.strip():
                    external_block = f"# ⚖️ МНЕНИЕ ВНЕШНЕГО ЮРИДИЧЕСКОГО ИИ\n\n{external_ai_opinion}\n\n"
                else:
                    external_block = ""
                
                instructions = """# ️ ВАЖНЫЕ ИНСТРУКЦИИ:
1. Ты работаешь с НАУЧНЫМ/ЮРИДИЧЕСКИМ материалом. Точность критически важна.
2. Используй ТОЛЬКО информацию из предоставленных источников выше.
3. НЕ выдумывай факты, которых нет в тексте.
4. Если информации недостаточно — прямо скажи об этом.
5. ЦИТИРУЙ конкретные места из источников (можно дословно).
6. Если делаешь вывод — объясни, на чем он основан.
7. Если есть мнение внешнего ИИ — учти его, но проверь на соответствие источникам.

"""
                
                full_context = context_block + external_block + instructions

                # ChatGPT
                with st.spinner("📝 ChatGPT анализирует материалы..."):
                    draft = ask(MODEL_CHATGPT, f"""Ты — научный/юридический ассистент. Твоя задача: подготовить ЧЕРНОВИК на основе предоставленных материалов.

{full_context}
## ЗАДАЧА:
{task}

## ТВОЙ ОТВЕТ ДОЛЖЕН:
- Строго опираться на источники (если они есть)
- Содержать конкретные факты и цитаты
- Иметь четкую структуру (заголовки, пункты)
- Выделять ключевые идеи и аргументы
- Если есть мнение внешнего ИИ — интегрируй его в структуру

Начни с анализа источников, затем предложи структуру ответа.""")
                
                st.markdown(f"""<div class="ai-card card-chatgpt"><h3 style="margin: 0 0 15px 0; color: #D4A373;">1️⃣ Черновик от ChatGPT</h3></div>""", unsafe_allow_html=True)
                st.write(draft)

                # Claude
                with st.spinner("🔍 Claude проверяет точность..."):
                    critique = ask(MODEL_CLAUDE, f"""Ты — строгий научный/юридический рецензент. Твоя задача: ПРОВЕРИТЬ черновик на соответствие источникам.

{full_context}
## ЗАДАЧА, которую ставили:
{task}

## ЧЕРНОВИК, который нужно проверить:
{draft}

## ПРОВЕРЬ:
1. ✅ Все ли факты из черновика ЕСТЬ в источниках?
2. ✅ Нет ли выдуманных фактов или искажений?
3. ✅ Корректно ли процитированы источники?
4. ⚠️ Какие утверждения требуют уточнения или доказательства?
5. ❌ Что упущено из источников, но важно для задачи?
6. ️ Если есть мнение внешнего ИИ — проверь его на соответствие источникам и актуальность

Будь МАКСИМАЛЬНО строгим. Если черновик противоречит источникам — укажи это явно.""")
                
                st.markdown(f"""<div class="ai-card card-claude"><h3 style="margin: 0 0 15px 0; color: #E07A5F;">2️⃣ Критика и проверка от Claude</h3></div>""", unsafe_allow_html=True)
                st.write(critique)

                # Qwen
                with st.spinner(" Qwen готовит финальную версию..."):
                    final = ask(MODEL_QWEN, f"""Ты — главный редактор научной/юридической работы. Твоя задача: создать ФИНАЛЬНУЮ ВЕРСИЮ, исправив все ошибки.

{full_context}
## ЗАДАЧА:
{task}

## ЧЕРНОВИК (ChatGPT):
{draft}

## КРИТИКА И ЗАМЕЧАНИЯ (Claude):
{critique}

## ТВОЙ ОТВЕТ ДОЛЖЕН:
1. ✅ Исправить ВСЕ ошибки, указанные Claude
2. ✅ Убрать всё, чего нет в источниках (если есть источники)
3. ✅ Добавить цитаты и ссылки на конкретные места в источниках
4. ✅ Сохранить структуру и логику
5. ✅ Если чего-то не хватает в источниках — честно указать: "В источниках нет информации о..."
6. ️ Если есть мнение внешнего ИИ — интегрируй его в финальный текст, если оно корректно и подтверждено источниками
7. ✅ Оформить ответ как готовый научный/юридический материал (заголовки, списки, абзацы)

Создай финальную, ПРОВЕРЕННУЮ версию.""")
                
                st.markdown(f"""<div class="ai-card card-qwen"><h3 style="margin: 0 0 15px 0; color: #8B5E3C;">3️⃣ ✅ Финальный результат от Qwen</h3><p style="margin: 0; color: #A89F91;">Совет ИИ завершил работу!</p></div>""", unsafe_allow_html=True)
                st.write(final)

                # --- СОЗДАНИЕ WORD ДОКУМЕНТА ---
                doc = docx.Document()
                doc.add_heading('Задача:', level=1)
                doc.add_paragraph(task)
                
                if file_name:
                    doc.add_heading('Загруженный файл:', level=2)
                    doc.add_paragraph(file_name)
                
                if external_ai_opinion.strip():
                    doc.add_heading('Мнение внешнего ИИ:', level=2)
                    doc.add_paragraph(external_ai_opinion)
                
                doc.add_heading('1. Черновик от ChatGPT', level=2)
                doc.add_paragraph(draft)
                doc.add_heading('2. Критика и проверка от Claude', level=2)
                doc.add_paragraph(critique)
                doc.add_heading('3. Финальный результат от Qwen', level=2)
                doc.add_paragraph(final)

                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                    doc.save(tmp.name)
                    tmp_path = tmp.name
                
                with open(tmp_path, "rb") as f:
                    st.download_button(
                        label="📥 Скачать полный отчёт в Word",
                        data=f,
                        file_name=f"Совет_ИИ_{datetime.now().strftime('%d%m%Y_%H%M')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                os.remove(tmp_path)

                entry = {
                    "project_id": selected_project["id"],
                    "date": datetime.now().strftime("%d.%m.%Y %H:%M"),
                    "task": task,
                    "file": file_name,
                    "external_ai": external_ai_opinion[:200] if external_ai_opinion.strip() else "",
                    "draft": draft,
                    "critique": critique,
                    "final": final
                }
                save_to_history(entry)
                st.sidebar.success("💾 Сохранено в историю проекта!")

            except Exception as e: st.error(f" Что-то пошло не так: {str(e)}")
