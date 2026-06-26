import os
import json
import logging
import re
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
import telebot
import openai
from dotenv import load_dotenv
from docx import Document
from docx.shared import Inches, Pt

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - [%(levelname)s] - %(message)s',
    handlers=[
        logging.FileHandler("bot.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)


def strip_tags(html_fragment):
    """Убирает HTML-теги из фрагмента, заменяя их на пробелы."""
    text = re.sub(r'<br\s*/?>', '\n', html_fragment, flags=re.IGNORECASE)
    text = re.sub(r'<p[^>]*>', '\n', text, flags=re.IGNORECASE)
    text = re.sub(r'</p>', '\n', text, flags=re.IGNORECASE)
    text = re.sub(r'<[^>]+>', ' ', text)
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'\n\s*\n+', '\n', text)
    return text.strip()


def extract_html_content(html_content):
    """
    Умный парсер HTML-таблиц: извлекает содержимое ПРАВОЙ колонки каждой строки таблицы.
    Если HTML не содержит таблиц — извлекает весь текст как обычно (fallback).

    Логика:
    - Для каждой строки <tr> находим все ячейки <td>.
    - Если ячеек >= 2, берём ПОСЛЕДНЮЮ (правую) — там спецификация.
    - Если ячейка одна — берём её (заголовки и широкие строки).
    - Короткие строки левой колонки (< 120 символов) при наличии правой — пропускаются.
    """
    # Удаляем стили и скрипты
    html_content = re.sub(r'<style[^>]*>[\s\S]*?</style>', '', html_content, flags=re.IGNORECASE)
    html_content = re.sub(r'<script[^>]*>[\s\S]*?</script>', '', html_content, flags=re.IGNORECASE)

    # Ищем строки таблицы
    rows = re.findall(r'<tr[^>]*>([\s\S]*?)</tr>', html_content, flags=re.IGNORECASE)

    if not rows:
        # Таблиц нет — просто убираем теги
        logging.info("HTML-таблицы не обнаружены. Используется стандартная очистка.")
        text = re.sub(r'<[^>]+>', '\n', html_content)
        text = re.sub(r'[ \t]+', ' ', text)
        text = re.sub(r'\n\s*\n+', '\n', text)
        return text.strip()

    texts = []
    for row in rows:
        # Находим все <td> в строке
        cells = re.findall(r'<td[^>]*>([\s\S]*?)</td>', row, flags=re.IGNORECASE)

        if not cells:
            # Строка без <td> — возможно, заголовок <th>
            header_cells = re.findall(r'<th[^>]*>([\s\S]*?)</th>', row, flags=re.IGNORECASE)
            for cell in header_cells:
                t = strip_tags(cell)
                if t and len(t) > 250:
                    texts.append(t)
            continue

        if len(cells) == 1:
            # Одна ячейка — берём её (широкая строка или заголовок)
            t = strip_tags(cells[0])
            if t and len(t) > 250:
                texts.append(t)
        else:
            # Несколько ячеек — берём ТОЛЬКО правую (последнюю)
            right_cell_text = strip_tags(cells[-1])
            # Оставляем только большие блоки текста (спецификации), игнорируя короткие метаданные
            if right_cell_text and len(right_cell_text) > 250:
                texts.append(right_cell_text)

    result = '\n'.join(texts)
    logging.info(f"Извлечено строк из правой колонки таблицы: {len(texts)}. Символов: {len(result)}.")
    return result


def deduplicate_text_lines(text, window=10):
    """
    Убирает повторяющиеся блоки строк из текста ещё до отправки в ИИ.
    Алгоритм: скользящее окно из `window` строк. Если текущая строка уже встречалась
    ранее в тексте, она пропускается.
    """
    lines = text.split('\n')
    seen = set()
    result = []

    for line in lines:
        normalized = line.strip().lower()
        # Пропускаем пустые и очень короткие строки (не дедуплицируем)
        if len(normalized) < 10:
            result.append(line)
            continue
        if normalized in seen:
            continue
        seen.add(normalized)
        result.append(line)

    deduped_text = '\n'.join(result)
    original_lines = len(lines)
    new_lines = len(result)
    removed = original_lines - new_lines
    if removed > 0:
        logging.info(f"Текстовая дедупликация: удалено {removed} повторяющихся строк ({original_lines} → {new_lines}).")
    return deduped_text


load_dotenv()
BOT_TOKEN = os.getenv("TELEGRAM_TOKEN")
API_KEY = os.getenv("OPENAI_API_KEY")
API_BASE = os.getenv("OPENAI_API_BASE")
AI_MODEL = os.getenv("AI_MODEL")

if not API_KEY:
    # Обратная совместимость с VselLM через старый ключ GEMINI_API_KEY
    API_KEY = os.getenv("GEMINI_API_KEY")
    if not API_BASE:
        API_BASE = "https://api.vsellm.ru/v1"
    if not AI_MODEL:
        AI_MODEL = "openai/gpt-4o-mini"
else:
    # Использование официального OpenAI
    if not AI_MODEL:
        AI_MODEL = "gpt-4o-mini"

print(f"DEBUG: Загруженный ключ: {API_KEY[:8] + '...' if API_KEY else 'None'}")
print(f"DEBUG: Используемая модель: {AI_MODEL}")
print(f"DEBUG: Базовый URL API: {API_BASE or 'https://api.openai.com/v1'}")

bot = telebot.TeleBot(BOT_TOKEN)

if API_BASE:
    ai_client = openai.OpenAI(api_key=API_KEY, base_url=API_BASE)
else:
    ai_client = openai.OpenAI(api_key=API_KEY)

allowed_users_raw = os.getenv("ALLOWED_USERS", "")
ALLOWED_USERS = [int(uid.strip()) for uid in allowed_users_raw.split(",") if uid.strip().isdigit()]

AI_SYSTEM_PROMPT = """
Ты — эксперт по анализу технических спецификаций. Твоя задача — взять текст из HTML-файла, убрать дубликаты и выдать строго структурированную разметку документа в формате JSON для последующей сборки в Word.

ВЫХОДНОЙ ФОРМАТ:
Выдай ответ СТРОГО в формате JSON-объекта, содержащего ключ "items" с массивом объектов внутри. Не добавляй никаких markdown-разметок (типа ```json), только чистый JSON-объект.
Пример структуры ответа:
{
  "items": [
    {"type": "item_description", "text": "Вводный текст или абзац спецификации..."},
    {"type": "item_heading", "text": "1. ДОСКА МАГНИТНАЯ НА КОЛЕСИКАХ – 1 шт"},
    {"type": "item_description", "text": "Модель... Гарантийный срок... Характеристики..."}
  ]
}

ЖЕСТКИЕ ПРАВИЛА ЛОГИКИ И ФОРМАТИРОВАНИЯ:
1. ВВОДНЫЙ ТЕКСТ И ОБЫЧНЫЕ АБЗАЦЫ: Если в самом начале HTML-файла идет общий текст (ссылки на пункты правил, требования к заявкам, инструкции для поставщика), ОБЯЗАТЕЛЬНО извлеки его полностью и отметь типом "item_description". Каждые отдельные абзацы этого текста разделяй на разные объекты "item_description", чтобы в Word они легли красивыми стандартными абзацами.
2. СОХРАНЕНИЕ ОРИГИНАЛЬНОЙ СТРУКТУРЫ И НУМЕРАЦИИ: Категорически запрещено придумывать искусственные глобальные разделы (такие как "1. МЕБЕЛЬ" или "2. ОБОРУДОВАНИЕ"), если их явно нет в исходном HTML. Сохраняй ту нумерацию позиций, которая идет в файле (например, "1. ДОСКА...", "2. КОВЕР..."). Если в нумерации источника есть пропуски (например, после пункта 9 сразу идет пункт 11) или у позиции вовсе нет номера (например, "КРЕСЛО ДЛЯ УЧИТЕЛЯ"), переноси этот текст ОДИН В ОДИН, ничего не исправляя и не добавляя от себя.
3. НАЗВАНИЯ ПОЗИЦИЙ (item_heading): Отмечай этим типом строго названия товаров/комплектов с их количеством. Не объединяй название позиции и её характеристики в одну строку — название должно идти отдельным заголовком.
4. ОПИСАНИЕ ПОЗИЦИЙ (item_description): Все технические характеристики, модели, заводы-изготовители, адреса и требования, которые идут ПОСЛЕ названия позиции, переноси целиком. Категорически запрещено сокращать параметры, удалять артикулы, изменять формулировки или сжимать текст.
5. УДАЛЕНИЕ ДУБЛИКАТОВ: Если из-за табличной верстки HTML одна и та же позиция со всем текстом дублируется несколько раз подряд, оставь только ОДНУ копию, а повторы отфильтруй.
"""


def extract_all_valid_items(s):
    """
    Сканирует текст и находит все подстроки, представляющие собой валидные JSON-объекты
    с ключами 'type' и 'text'. Это нужно для восстановления усеченного JSON.
    """
    items = []
    n = len(s)
    i = 0
    while i < n:
        start = s.find('{', i)
        if start == -1:
            break
        
        bracket_count = 0
        in_str = False
        esc = False
        found_end = -1
        
        for j in range(start, n):
            char = s[j]
            if char == '"' and not esc:
                in_str = not in_str
            if char == '\\' and in_str:
                esc = not esc
            else:
                esc = False
                
            if not in_str:
                if char == '{':
                    bracket_count += 1
                elif char == '}':
                    bracket_count -= 1
                    if bracket_count == 0:
                        found_end = j
                        break
        
        if found_end != -1:
            candidate = s[start:found_end+1]
            try:
                obj = json.loads(candidate)
                if isinstance(obj, dict) and "type" in obj and "text" in obj:
                    items.append(obj)
                    i = found_end + 1
                    continue
            except Exception:
                pass
        i = start + 1
    return items


def parse_ai_json(raw_response_text):
    """
    Парсит JSON-ответ от ИИ с поддержкой восстановления усеченных данных.
    """
    raw_response_text = raw_response_text.strip()
    # Убираем разметку markdown ```json ... ```
    if raw_response_text.startswith("```json"):
        raw_response_text = raw_response_text[7:]
    elif raw_response_text.startswith("```"):
        raw_response_text = raw_response_text[3:]
    if raw_response_text.endswith("```"):
        raw_response_text = raw_response_text[:-3]
    raw_response_text = raw_response_text.strip()

    try:
        json_data = json.loads(raw_response_text)
        items = json_data.get("items", [])
        if isinstance(items, list):
            return items
    except json.JSONDecodeError as e:
        logging.warning(f"Ошибка декодирования полного JSON от ИИ: {e}. Пытаемся извлечь валидные блоки...")

    # Если не удалось распарсить весь JSON, извлекаем отдельные валидные объекты
    items = extract_all_valid_items(raw_response_text)
    if items:
        logging.info(f"Успешно извлечено {len(items)} объектов из поврежденного/усеченного JSON.")
        return items
    
    raise ValueError("Не удалось распарсить JSON и извлечь из него валидные объекты.")


def deduplicate_items(items):
    """
    Удаляет повторяющиеся элементы из сборного списка позиций после объединения всех чанков.
    Сравнение по нормализованному тексту (lowercase, без пробелов).
    """
    seen_texts = set()
    result = []
    for item in items:
        text = item.get('text', '')
        key = ' '.join(text.lower().split())  # нормализация: lowercase + убираем лишние пробелы
        if key and key in seen_texts:
            continue
        seen_texts.add(key)
        result.append(item)
    removed = len(items) - len(result)
    if removed > 0:
        logging.info(f"Пост-обработка: удалено {removed} дублирующихся элементов документа ({len(items)} → {len(result)}).")
    return result


def create_word_document(structured_data, output_path="output.docx"):
    """Создает документ Word на основе структуры от ИИ"""
    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    for block in structured_data:
        b_type = block.get("type")
        b_text = block.get("text", "").strip()

        if not b_text:
            continue

        p = doc.add_paragraph()

        if b_type == "main_heading":
            p.paragraph_format.space_before = Pt(18)
            p.paragraph_format.space_after = Pt(6)
            run = p.add_run(b_text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(14)
            run.bold = True

        elif b_type == "item_heading":
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(b_text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)
            run.bold = True

        elif b_type == "item_description":
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(6)
            run = p.add_run(b_text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)
            run.bold = False

    doc.save(output_path)
    return output_path


# ЗАЩИТНЫЙ ЩИТ
@bot.message_handler(
    func=lambda message: message.from_user.id not in ALLOWED_USERS,
    content_types=['text', 'document', 'photo', 'audio', 'video', 'voice', 'sticker']
)
def block_unauthorized(message):
    logging.warning(f"Попытка чужого доступа! ID: {message.from_user.id}, Username: @{message.from_user.username}")
    bot.reply_to(message, "🛑 **Доступ ограничен.**\n\n"
                          "Вы не находитесь в белом списке администратора этого бота, "
                          "поэтому я не могу обрабатывать ваши файлы. 🔐")


# КОМАНДЫ /start И /help
@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):
    markup = telebot.types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=1)
    btn_info = telebot.types.KeyboardButton("ℹ️ Что мне делать?")
    markup.add(btn_info)

    welcome_text = (
        f"👋 Привет, {message.from_user.first_name}!\n\n"
        f"Я — твой интеллектуальный ассистент по конвертации документов. 🚀\n\n"
        f"**Что я делаю:** Ты кидаешь мне сложный, нечитаемый `.html` файл, а я с помощью нейросети "
        f"вытаскиваю из него суть и превращаю в идеально оформленный документ Word (`.docx`).\n\n"
        f"👇 Просто перетащи сюда нужный файл или нажми кнопку ниже!"
    )
    bot.send_message(message.chat.id, welcome_text, reply_markup=markup, parse_mode="Markdown")


# ОБРАБОТКА НАЖАТИЯ НА КНОПКУ ИНСТРУКЦИИ
@bot.message_handler(func=lambda message: message.text == "ℹ️ Что мне делать?")
def show_instruction(message):
    instruction = (
        "📖 **Инструкция очень простая:**\n\n"
        "1. Найди на компьютере или телефоне нужный файл с расширением `.html`.\n"
        "2. Отправь его мне **как файл** (не как текст, не ссылкой).\n"
        "3. Подожди немного, пока ИИ анализирует структуру.\n"
        "4. Забирай готовый, красивый `.docx` файл! 📝"
    )
    bot.send_message(message.chat.id, instruction, parse_mode="Markdown")


def split_text_into_chunks(text, max_chunk_size=6000):
    """Делит текст на части примерно по max_chunk_size символов, разбивая строго по переносам строк,
    чтобы не разрывать предложения посреди слова."""
    lines = text.split('\n')
    chunks = []
    current_chunk = []
    current_length = 0

    for line in lines:
        if current_length + len(line) + 1 > max_chunk_size:
            if current_chunk:
                chunks.append('\n'.join(current_chunk))
            current_chunk = [line]
            current_length = len(line)
        else:
            current_chunk.append(line)
            current_length += len(line) + 1

    if current_chunk:
        chunks.append('\n'.join(current_chunk))
    return chunks


# ОБРАБОТКА HTML-ДОКУМЕНТОВ (ОСНОВНАЯ ЛОГИКА С ЧАНКИНГОМ)
@bot.message_handler(content_types=['document'])
def handle_docs(message):
    if not message.document.file_name.lower().endswith('.html'):
        bot.reply_to(message,
                     "❌ Ошибка: Я умею работать **только с HTML-файлами** (название должно заканчиваться на `.html`).\n"
                     "Пожалуйста, пришлите правильный файл!")
        return

    status_msg = bot.reply_to(message, "⏳ Принял файл! Начинаю обработку...")
    out_filename = f"processed_{message.document.file_name.replace('.html', '.docx')}"

    try:
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        html_content = downloaded_file.decode('utf-8', errors='ignore')

        # 1. Парсим HTML и извлекаем содержимое ПРАВОЙ колонки таблицы
        cleaned_text = extract_html_content(html_content)

        # 1.5 Удаляем повторяющиеся строки из очищенного текста (ДО отправки в ИИ)
        cleaned_text = deduplicate_text_lines(cleaned_text)
        bot.edit_message_text(
            f"⏳ Структура файла разобрана. Удаление дубликатов... Делю на части...",
            message.chat.id, status_msg.message_id
        )

        # 2. Делим текст на безопасные части
        text_chunks = split_text_into_chunks(cleaned_text, max_chunk_size=6000)
        logging.info(f"Текст разбит на {len(text_chunks)} частей для отправки ИИ.")

        bot.edit_message_text(f"⏳ Файл успешно разбит на {len(text_chunks)} частей. Начинаю анализ...",
                               message.chat.id, status_msg.message_id)

        combined_items = [None] * len(text_chunks)
        completed_count = 0
        count_lock = threading.Lock()

        def process_chunk(index, chunk):
            """Обрабатывает один чанк и возвращает (index, items)."""
            logging.info(f"Отправка части {index + 1}/{len(text_chunks)} к ИИ API ({AI_MODEL})...")
            response = ai_client.chat.completions.create(
                model=AI_MODEL,
                messages=[
                    {"role": "system", "content": AI_SYSTEM_PROMPT},
                    {"role": "user",
                     "content": f"Вот часть текста спецификации (Часть {index + 1} из {len(text_chunks)}):\n\n{chunk}"}
                ],
                response_format={"type": "json_object"},
                temperature=0.2,
                max_tokens=16384
            )
            raw = response.choices[0].message.content
            items = parse_ai_json(raw)
            logging.info(f"Часть {index + 1} успешно обработана. Добавлено элементов: {len(items)}")
            return index, items

        # 3. Отправляем чанки параллельно (до 8 одновременных запросов)
        MAX_WORKERS = 8
        bot.edit_message_text(
            f"⏳ Начинаю параллельный анализ {len(text_chunks)} частей (до {MAX_WORKERS} одновременно)...",
            message.chat.id, status_msg.message_id
        )

        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            futures = {executor.submit(process_chunk, i, chunk): i for i, chunk in enumerate(text_chunks)}
            for future in as_completed(futures):
                idx, items = future.result()  # выбросит исключение наверх, если что-то пошло не так
                combined_items[idx] = items
                with count_lock:
                    completed_count += 1
                    done = completed_count
                bot.edit_message_text(
                    f"⏳ Обработано частей: {done}/{len(text_chunks)}...",
                    message.chat.id, status_msg.message_id
                )

        # Собираем итоговый список в правильном порядке
        combined_items = [item for part in combined_items if part for item in part]

        # Финальная пост-обработка: удаляем дубликаты между чанками
        combined_items = deduplicate_items(combined_items)

        logging.info("Все части успешно получены от ИИ и объединены.")
        bot.edit_message_text("⏳ Все части обработаны! Собираю итоговый Word-документ...", message.chat.id,
                              status_msg.message_id)

        # 4. Создаем документ на основе ВСЕХ собранных частей
        create_word_document(combined_items, out_filename)
        logging.info(f"Документ {out_filename} успешно создан.")

        # 5. Отправляем готовый файл пользователю
        with open(out_filename, 'rb') as doc_file:
            bot.send_document(message.chat.id, doc_file)

        bot.edit_message_text("✅ Готово! Файл отправлен.", message.chat.id, status_msg.message_id)

    except openai.OpenAIError as e:
        logging.error(f"Ошибка со стороны API: {e}")
        bot.edit_message_text(f"⚠️ Произошла ошибка при запросе к нейросети. Подробности: {e}",
                               message.chat.id, status_msg.message_id)

    except Exception as e:
        logging.exception("Произошла непредвиденная ошибка:")
        bot.edit_message_text(f"❌ Произошла ошибка. Подробности записаны в файл bot.log.", message.chat.id,
                              status_msg.message_id)
    finally:
        # В самом конце всегда удаляем временный файл
        if os.path.exists(out_filename):
            os.remove(out_filename)
            logging.info(f"Файл {out_filename} успешно удален с диска.")


# ЗАЩИТА ОТ ДУРАКА №2
@bot.message_handler(content_types=['text', 'photo', 'audio', 'video', 'voice', 'sticker'])
def handle_wrong_content(message):
    if message.text and message.text.startswith('/'):
        return

    bot.reply_to(message, "🤖 Хей! Я — специализированный бот-конвертер.\n\n"
                          "Я не умею просто болтать, обрабатывать фото или стикеры. "
                          "Чтобы я сработал, просто **пришли мне HTML-файл как документ**, "
                          "и я превращу его в структурированный Word-файл! 📄➡️📝")


if __name__ == '__main__':
    logging.info("Бот успешно запущен и готов к работе...")
    bot.infinity_polling()