import os
import json
import logging
import telebot
from dotenv import load_dotenv

# Импорты для работы с Google Gemini
from google import genai
from google.genai import errors
from google.genai import types

# Настройка логирования (твоя рабочая конфигурация)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - [%(levelname)s] - %(message)s',
    handlers=[
        logging.FileHandler("bot.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)

# Загрузка переменных окружения из файла .env
load_dotenv()
BOT_TOKEN = os.getenv("TELEGRAM_TOKEN")  # Проверь, что имя переменной совпадает с твоим .env
GEMINI_KEY = os.getenv("GEMINI_API_KEY")

# Инициализация клиентов
bot = telebot.TeleBot(BOT_TOKEN)
ai_client = genai.Client(api_key=GEMINI_KEY)


# Читаем белый список и превращаем его в список чисел Python
allowed_users_raw = os.getenv("ALLOWED_USERS", "")
ALLOWED_USERS = [int(uid.strip()) for uid in allowed_users_raw.split(",") if uid.strip().isdigit()]
# =====================================================================
# ВСТАВЬ СЮДА СВОИ ДАННЫЕ: Твой системный промпт и функцию создания Word
# =====================================================================
AI_SYSTEM_PROMPT = """ЗДЕСЬ ДОЛЖЕН БЫТЬ ТВОЙ СУЩЕСТВУЮЩИЙ СИСТЕМНЫЙ ПРОМПТ"""


def create_word_document(structured_data, out_filename):
    """Твоя существующая функция, которая собирает из JSON файл Word"""
    # ... (весь твой код для docx библиотеки) ...
    pass


# =====================================================================
# =====================================================================
# ЗАЩИТНЫЙ ЩИТ: Блокируем всех, кого нет в белом списке
# =====================================================================
@bot.message_handler(
    func=lambda message: message.from_user.id not in ALLOWED_USERS,
    content_types=['text', 'document', 'photo', 'audio', 'video', 'voice', 'sticker']
)
def block_unauthorized(message):
    # Записываем в лог, кто пытался зайти без разрешения
    logging.warning(f"Попытка чужого доступа! ID: {message.from_user.id}, Username: @{message.from_user.username}")

    bot.reply_to(message, "🛑 **Доступ ограничен.**\n\n"
                          "Вы не находитесь в белом списке администратора этого бота, "
                          "поэтому я не могу обрабатывать ваши файлы. 🔐")


# =====================================================================




# КОМАНДЫ /start И /help (КРАСИВЫЙ ИНТЕРФЕЙС)
@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):
    # Используем полный путь telebot.types, чтобы избежать конфликта с типами Gemini
    markup = telebot.types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=1)
    btn_info = telebot.types.KeyboardButton("ℹ️ Что мне делать?")
    markup.add(btn_info)

    welcome_text = (
        f"👋 Привет, {message.from_user.first_name}!\n\n"
        f"Я — твой интеллектуальный ассистент по конвертации документов. 🚀\n\n"
        f"**Что я делаю:** Ты кидаешь мне сложный, нечитаемый `.html` файл, а я с помощью нейросети "
        f"Gemini 2.5 вытаскиваю из него суть и превращаю в идеально оформленный документ Word (`.docx`).\n\n"
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
    # ЗАЩИТА ОТ ДУРАКА №1: Проверяем, что файл действительно .html
    if not message.document.file_name.lower().endswith('.html'):
        bot.reply_to(message,
                     "❌ Ошибка: Я умею работать **только с HTML-файлами** (название должно заканчиваться на `.html`).\n"
                     "Пожалуйста, пришлите правильный файл!")
        return  # Прерываем функцию, код дальше не пойдет

    status_msg = bot.reply_to(message, "⏳ Принял файл! Начинаю обработку...")

    try:
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        html_content = downloaded_file.decode('utf-8', errors='ignore')

        logging.info("Отправка запроса к Gemini API...")
        response = ai_client.models.generate_content(
            model='gemini-2.5-flash',
            contents=f"Вот текст из HTML-файла:\n\n{html_content}",
            config=types.GenerateContentConfig(
                system_instruction=AI_SYSTEM_PROMPT,
                response_mime_type="application/json",
                temperature=0.2
            )
        )

        structured_data = json.loads(response.text)
        logging.info("Ответ от Gemini успешно получен.")

        out_filename = f"processed_{message.document.file_name.replace('.html', '.docx')}"
        create_word_document(structured_data, out_filename)
        logging.info(f"Документ {out_filename} успешно создан.")

        with open(out_filename, 'rb') as doc_file:
            bot.send_document(message.chat.id, doc_file)

        bot.edit_message_text("✅ Готово! Отправляю файл.", message.chat.id, status_msg.message_id)
        os.remove(out_filename)
        logging.info(f"Файл {out_filename} отправлен и удален с компьютера.")

    # 1. Ловим перегрузку серверов Google (Ошибка 503)
    except errors.ServerError as e:
        logging.warning(f"Сервера Google перегружены: {e}")
        bot.edit_message_text("⚠️ Извините, сервера искусственного интеллекта сейчас перегружены (ошибка 503). "
                              "Google просит подождать пару минут. Пожалуйста, попробуйте отправить файл чуть позже! ⏳",
                              message.chat.id, status_msg.message_id)

    # 2. Ловим невалидный JSON от нейросети
    except json.JSONDecodeError as e:
        logging.error(f"ИИ вернул невалидный JSON: {e}")
        bot.edit_message_text("❌ Ошибка: ИИ вернул некорректный формат данных.", message.chat.id, status_msg.message_id)

    # 3. Ловим любые другие непредвиденные системные ошибки
    except Exception as e:
        logging.exception("Произошла непредвиденная ошибка:")
        bot.edit_message_text(f"❌ Произошла ошибка. Подробности записаны в файл bot.log.", message.chat.id,
                              status_msg.message_id)


# ЗАЩИТА ОТ ДУРАКА №2: Ловим всё, кроме документов (текст, фото, стикеры и т.д.)
@bot.message_handler(content_types=['text', 'photo', 'audio', 'video', 'voice', 'sticker'])
def handle_wrong_content(message):
    # Пропускаем команды /start и /help, у них своя логика выше
    if message.text and message.text.startswith('/'):
        return

    bot.reply_to(message, "🤖 Хей! Я — специализированный бот-конвертер.\n\n"
                          "Я не умею просто болтать, обрабатывать фото или стикеры. "
                          "Чтобы я сработал, просто **пришли мне HTML-файл как документ**, "
                          "и я превращу его в структурированный Word-файл! 📄➡️📝")


# ЗАПУСК БОТА В БЕСКОНЕЧНОМ ЦИКЛЕ
if __name__ == '__main__':
    logging.info("Бот успешно запущен и готов к работе...")
    bot.infinity_polling()