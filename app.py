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

st.set_page_config(page_title="Совет ИИ", page_icon="🤖", layout="wide")
st.title("🤖 Совет ИИ")

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
    # Также удаляем задачи этого проекта из истории
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
        doc = docx.Document(file)
        text = ""
        for paragraph in doc.paragraphs:
            text += paragraph.text + "\n"
        return text
    else:
        raise Exception("Неподдерживаемый формат. Используй PDF, DOCX или TXT.")

def ask(model, prompt):
    key = load_key()
    if not key:
        raise Exception("Не найден файл key.txt или секрет API_KEY")
    client = OpenAI(api_key=key, base_url=BASE_URL)
    answer = client.chat.completions.create(
        model=model,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.3,
    )
    return answer.choices[0].message.content

# --- БОКОВАЯ ПАНЕЛЬ: ПРОЕКТЫ ---
st.sidebar.header("📁 Проекты")

projects = load_projects()

# Кнопка создания нового проекта
with st.sidebar.expander("➕ Новый проект"):
    new_project_name = st.text_input("Название проекта:", placeholder="Например: Научные статьи")
    new_project_desc = st.text_area("Описание (необязательно):", height=80)
    if st.button("Создать проект"):
        if new_project_name.strip():
            create_project(new_project_name.strip(), new_project_desc.strip())
            st.success(f"Проект '{new_project_name}' создан!")
            st.rerun()
        else:
            st.warning("Введите название проекта")

# Выбор активного проекта
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

# --- БОКОВАЯ ПАНЕЛЬ: ИСТОРИЯ ЗАДАЧ ---
if selected_project:
    st.sidebar.header(f"📂 История: {selected_project['name']}")
    history = load_history()
    project_history = [h for h in history if h.get("project_id") == selected_project["id"]]
    
    if project_history:
        st.sidebar.markdown("---")
        for i, entry in enumerate(reversed(project_history)):
            short_task = entry["task"][:40] + "..." if len(entry["task"]) > 40 else entry["task"]
            with st.sidebar.expander(f"📝 {short_task}\n {entry['date']}"):
                st.write("**Задача:**")
                st.write(entry["task"])
                if entry.get("file"):
                    st.write(f"**Файл:** {entry['file']}")
                st.write("**Финальный результат:**")
                st.success(entry["final"])
    else:
        st.sidebar.info("В этом проекте пока нет задач")
else:
    st.sidebar.header("📂 История задач")
    st.sidebar.info("Выбери проект слева")

st.sidebar.markdown("---")
if load_key():
    st.sidebar.success("🔑 Ключ активен")
else:
    st.sidebar.error("❌ Ключ не найден")

# --- ОСНОВНОЙ ИНТЕРФЕЙС ---
if selected_project:
    st.write(f"Проект: **{selected_project['name']}**")
    if selected_project.get("description"):
        st.caption(selected_project["description"])
else:
    st.write("Создай или выбери проект в боковой панели слева")

st.markdown("---")
st.write("Научный совет: ChatGPT, Claude и Qwen работают с твоими источниками.")

if st.button("🔌 Проверить связь"):
    for name, model in [("ChatGPT", MODEL_CHATGPT), ("Claude", MODEL_CLAUDE), ("Qwen", MODEL_QWEN)]:
        try:
            ask(model, "ОК")
            st.success(f"{name} — на связи")
        except Exception as e:
            st.error(f"{name} — ошибка: {e}")

st.markdown("---")
st.subheader("📎 Загрузи научные материалы")
st.info("💡 Модели будут СТРОГО следовать загруженному тексту и не станут выдумывать.")

uploaded_file = st.file_uploader("Выбери PDF, DOCX или TXT:", type=['pdf', 'docx', 'txt'])

file_text = ""
file_name = ""
if uploaded_file is not None:
    try:
        file_text = extract_text(uploaded_file)
        file_name = uploaded_file.name
        if len(file_text) > MAX_FILE_SIZE:
            file_text = file_text[:MAX_FILE_SIZE]
            st.warning(f"⚠️ Файл большой. Взяты только первые {MAX_FILE_SIZE} символов.")
        
        st.success(f"✅ Файл загружен: {file_name} ({len(file_text)} символов)")
        with st.expander("👁 Посмотреть загруженный текст"):
            st.text(file_text[:2000] + "..." if len(file_text) > 2000 else file_text)
            
    except Exception as e:
        st.error(f"❌ Ошибка чтения файла: {e}")

st.markdown("---")
task = st.text_area("Опиши задачу:", height=150, 
                    placeholder="Например: 'На основе этого текста выдели ключевые тезисы для научной статьи'")

if st.button("🚀 Запустить обсуждение"):
    if not selected_project:
        st.warning("Сначала создай или выбери проект в боковой панели.")
    elif not task.strip():
        st.warning("Сначала введи задачу.")
    else:
        try:
            if file_text:
                context_block = f"""
#  ИСТОЧНИК (научный материал)

{file_text}

# ⚠️ ВАЖНЫЕ ИНСТРУКЦИИ:
1. Ты работаешь с НАУЧНЫМ материалом. Точность критически важна.
2. Используй ТОЛЬКО информацию из источника выше.
3. НЕ выдумывай факты, которых нет в тексте.
4. Если информации недостаточно — прямо скажи об этом.
5. ЦИТИРУЙ конкретные места из источника (можно дословно).
6. Если делаешь вывод — объясни, на чем он основан в тексте.

"""
            else:
                context_block = "\n# ⚠️ ВНИМАНИЕ: Источник не загружен. Отвечай на основе общих знаний.\n\n"

            with st.spinner("📝 ChatGPT анализирует источник..."):
                draft = ask(
                    MODEL_CHATGPT, 
                    f"""Ты — научный ассистент. Твоя задача: подготовить ЧЕРНОВИК на основе источника.

{context_block}
## ЗАДАЧА:
{task}

## ТВОЙ ОТВЕТ ДОЛЖЕН:
- Строго опираться на источник (если он есть)
- Содержать конкретные факты и цитаты из текста
- Иметь четкую структуру (заголовки, пункты)
- Выделять ключевые идеи и аргументы

Начни с анализа источника, затем предложи структуру ответа."""
                )
            st.subheader("1️⃣ Черновик от ChatGPT")
            st.write(draft)

            with st.spinner("🔍 Claude проверяет точность..."):
                critique = ask(
                    MODEL_CLAUDE,
                    f"""Ты — строгий научный рецензент. Твоя задача: ПРОВЕРИТЬ черновик на соответствие источнику.

{context_block}
## ЗАДАЧА, которую ставили:
{task}

## ЧЕРНОВИК, который нужно проверить:
{draft}

## ПРОВЕРЬ:
1. ✅ Все ли факты из черновика ЕСТЬ в источнике?
2. ✅ Нет ли выдуманных фактов или искажений?
3. ✅ Корректно ли процитирован источник?
4. ⚠️ Какие утверждения требуют уточнения или доказательства?
5. ❌ Что упущено из источника, но важно для задачи?

Будь МАКСИМАЛЬНО строгим. Если черновик противоречит источнику — укажи это явно."""
                )
            st.subheader("2️⃣ Критика и проверка от Claude")
            st.write(critique)

            with st.spinner("🎯 Qwen готовит финальную версию..."):
                final = ask(
                    MODEL_QWEN,
                    f"""Ты — главный редактор научной работы. Твоя задача: создать ФИНАЛЬНУЮ ВЕРСИЮ, исправив все ошибки.

{context_block}
## ЗАДАЧА:
{task}

## ЧЕРНОВИК (ChatGPT):
{draft}

## КРИТИКА И ЗАМЕЧАНИЯ (Claude):
{critique}

## ТВОЙ ОТВЕТ ДОЛЖЕН:
1. ✅ Исправить ВСЕ ошибки, указанные Claude
2. ✅ Убрать всё, чего нет в источнике (если есть источник)
3. ✅ Добавить цитаты и ссылки на конкретные места в источнике
4. ✅ Сохранить структуру и логику
5. ✅ Если чего-то не хватает в источнике — честно указать: "В источнике нет информации о..."
6. ✅ Оформить ответ как готовый научный материал (заголовки, списки, абзацы)

Создай финальную, ПРОВЕРЕННУЮ версию."""
                )
            st.subheader("3️⃣ ✅ Финальный результат от Qwen")
            st.success("Научный совет завершил работу!")
            st.write(final)

            entry = {
                "project_id": selected_project["id"],
                "date": datetime.now().strftime("%d.%m.%Y %H:%M"),
                "task": task,
                "file": file_name,
                "draft": draft,
                "critique": critique,
                "final": final
            }
            save_to_history(entry)
            st.sidebar.success("💾 Сохранено в историю проекта!")

        except Exception as e:
            st.error("❌ Что-то пошло не так: " + str(e))
