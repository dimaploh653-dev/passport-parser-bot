import os
import zipfile
import tempfile
import asyncio
from dotenv import load_dotenv
from telegram import Update
from telegram.ext import (
    ApplicationBuilder,
    MessageHandler,
    filters,
    ContextTypes,
)
from universal_parser import process_word_files
import nest_asyncio
from threading import Thread
from http.server import SimpleHTTPRequestHandler, HTTPServer

# =============================
# Настройки окружения
# =============================
load_dotenv()
BOT_TOKEN = os.getenv("BOT_TOKEN")

if not BOT_TOKEN:
    raise ValueError("❌ Не найден BOT_TOKEN! Добавь его в Environment на Render.")

PORT = int(os.environ.get("PORT", 10000))  # Render требует открытый порт

# =============================
# HTTP KeepAlive-сервер
# =============================
def run_web_server():
    """Простой HTTP сервер для Render KeepAlive"""
    handler = SimpleHTTPRequestHandler
    httpd = HTTPServer(("", PORT), handler)
    print(f"🌐 KeepAlive сервер запущен на порту {PORT}")
    httpd.serve_forever()

# =============================
# Обработка ZIP-файлов
# =============================
async def handle_zip(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.message.from_user
    file = await update.message.document.get_file()
    await update.message.reply_text("📦 Получен архив, начинаю обработку...")

    with tempfile.TemporaryDirectory() as tmpdir:
        zip_path = os.path.join(tmpdir, "input.zip")
        await file.download_to_drive(zip_path)

        try:
            with zipfile.ZipFile(zip_path, "r") as zip_ref:
                zip_ref.extractall(tmpdir)
        except zipfile.BadZipFile:
            await update.message.reply_text("❌ Ошибка: файл повреждён или не является ZIP-архивом.")
            return

        files = [
            os.path.join(tmpdir, f)
            for f in os.listdir(tmpdir)
            if f.lower().endswith(".docx")
        ]
        if not files:
            await update.message.reply_text("⚠️ В архиве нет файлов Word (.docx)")
            return

        await update.message.reply_text(f"🔍 Найдено {len(files)} файлов, запускаю парсер...")

        try:
            output_path = os.path.join(tmpdir, f"result_{user.id}.xlsx")
            process_word_files(files, output_path)
        except Exception as e:
            await update.message.reply_text(f"❌ Ошибка при парсинге: {e}")
            return

        await update.message.reply_document(
            document=open(output_path, "rb"),
            filename=f"parsed_{user.id}.xlsx",
            caption="✅ Парсинг завершён успешно!"
        )

# =============================
# Ответ по умолчанию
# =============================
async def handle_other(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("📩 Отправь мне ZIP-архив с файлами Word (.docx) для парсинга.")

# =============================
# Основной бот
# =============================
async def main():
    print("🤖 Бот запущен и готов к работе...")
    app = ApplicationBuilder().token(BOT_TOKEN).build()

    app.add_handler(MessageHandler(filters.Document.ZIP, handle_zip))
    app.add_handler(MessageHandler(filters.ALL, handle_other))

    await app.run_polling()

# =============================
# Запуск: KeepAlive + бот
# =============================
if __name__ == "__main__":
    nest_asyncio.apply()

    # Запускаем HTTP-сервер в отдельном потоке
    Thread(target=run_web_server, daemon=True).start()

    # Запускаем Telegram-бота
    asyncio.get_event_loop().run_until_complete(main())
