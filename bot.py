import os
import json
import logging
import re
import telebot
import openai
from dotenv import load_dotenv
from docx import Document
from docx.shared import Inches, Pt  # <-- Импортировали и Inches, и Pt для работы со шрифтами

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - [%(levelname)s] - %(message)s',
    handlers=[
        logging.FileHandler("bot.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)


def clean_html_garbage(html_content):
    """Удаляет из HTML-кода тяжелые теги <style> и <script>, которые путают ИИ"""
    html_content = re.sub(r'<style[^>]*>([\s\S]*?)</style>', '', html_content)
    html_content = re.sub(r'<script[^>]*>([\s\S]*?)</script>', '', html_content)
    html_content = re.sub(r'\n\s*\n', '\n', html_content)
    return html_content.strip()


# Загрузка переменных окружения
load_dotenv()
BOT_TOKEN = os.getenv("TELEGRAM_TOKEN")
API_KEY_VSELLM = os.getenv("GEMINI_API_KEY")

# Инициализация клиентов
bot = telebot.TeleBot(BOT_TOKEN)
ai_client = openai.OpenAI(
    api_key=API_KEY_VSELLM,
    base_url="https://api.vsellm.ru/v1"
)

# Читаем белый список пользователей
allowed_users_raw = os.getenv("ALLOWED_USERS", "")
ALLOWED_USERS = [int(uid.strip()) for uid in allowed_users_raw.split(",") if uid.strip().isdigit()]

# СИСТЕМНЫЙ ПРОМПТ (Исправлен под жесткие требования json_object)
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
        f"GPT-4o-mini вытаскиваю из него суть и превращаю в идеально оформленный документ Word (`.docx`).\n\n"
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


# ОБРАБОТКА HTML-ДОКУМЕНТОВ (ОСНОВНАЯ ЛОГИКА)
@bot.message_handler(content_types=['document'])
def handle_docs(message):
    if not message.document.file_name.lower().endswith('.html'):
        bot.reply_to(message,
                     "❌ Ошибка: Я умею работать **только с HTML-файлами** (название должно заканчиваться на `.html`).\n"
                     "Пожалуйста, пришлите правильный файл!")
        return

    status_msg = bot.reply_to(message, "⏳ Принял файл! Начинаю обработку...")

    try:
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        html_content = downloaded_file.decode('utf-8', errors='ignore')

        # Очищаем от мусора
        html_content = clean_html_garbage(html_content)

        logging.info("Отправка запроса к VselLM API (openai/gpt-4o-mini)...")

        response = ai_client.chat.completions.create(
            model='openai/gpt-4o-mini',
            messages=[
                {"role": "system", "content": AI_SYSTEM_PROMPT},
                {"role": "user", "content": f"Вот текст из HTML-файла:\n\n{html_content}"}
            ],
            response_format={"type": "json_object"},
            temperature=0.2
        )

        raw_response_text = response.choices[0].message.content

        # Безопасно достаем массив "items" из объекта ответа
        json_data = json.loads(raw_response_text)
        structured_data = json_data.get("items", [])

        logging.info("Ответ от нейросети успешно получен и десериализован.")

        out_filename = f"processed_{message.document.file_name.replace('.html', '.docx')}"
        create_word_document(structured_data, out_filename)
        logging.info(f"Документ {out_filename} успешно создан.")

        with open(out_filename, 'rb') as doc_file:
            bot.send_document(message.chat.id, doc_file)

        bot.edit_message_text("✅ Готово! Отправляю файл.", message.chat.id, status_msg.message_id)
        os.remove(out_filename)
        logging.info(f"Файл {out_filename} отправлен и удален.")

    except openai.OpenAIError as e:
        logging.error(f"Ошибка со стороны VselLM API: {e}")
        bot.edit_message_text(f"⚠️ Произошла ошибка при запросе к нейросети. Подробности: {e}",
                              message.chat.id, status_msg.message_id)

    except json.JSONDecodeError as e:
        logging.error(f"ИИ вернул невалидный JSON: {e}")
        bot.edit_message_text("❌ Ошибка: ИИ вернул некорректный формат данных.", message.chat.id, status_msg.message_id)

    except Exception as e:
        logging.exception("Произошла непредвиденная ошибка:")
        bot.edit_message_text(f"❌ Произошла ошибка. Подробности записаны в файл bot.log.", message.chat.id,
                              status_msg.message_id)


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