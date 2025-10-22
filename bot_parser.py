import os
import zipfile
import tempfile
import asyncio
import logging
import subprocess
from telegram import Update, InputFile
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    filters,
    ContextTypes,
)
from dotenv import load_dotenv

# ==========================================================
# 🧠 Настройки
# ==========================================================

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)

load_dotenv()  # для локального запуска
BOT_TOKEN = os.getenv("BOT_TOKEN")  # Render автоматически подставит это значение

if not BOT_TOKEN:
    raise ValueError("❌ Не найден BOT_TOKEN! Убедись, что он добавлен в Environment на Render.")

# ==========================================================
# 🚀 Команды
# ==========================================================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 Привет! Отправь мне ZIP-файл с паспортами (.docx), "
        "и я пришлю Excel-таблицу с результатами 📊"
    )

# ==========================================================
# 📦 Обработка ZIP-файлов
# ==========================================================

async def handle_zip(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        file = update.message.document
        if not file.file_name.endswith(".zip"):
            await update.message.reply_text("⚠️ Пожалуйста, отправь ZIP-файл.")
            return

        await update.message.reply_text("📂 Получил файл, обрабатываю... Это может занять до 1 минуты ⏳")

        # создаём временную директорию
        with tempfile.TemporaryDirectory() as tmpdir:
            zip_path = os.path.join(tmpdir, file.file_name)

            # скачиваем файл
            new_file = await file.get_file()
            await new_file.download_to_drive(zip_path)

            # распаковываем zip
            with zipfile.ZipFile(zip_path, "r") as zip_ref:
                zip_ref.extractall(tmpdir)

            # запускаем universal_parser.py
            parser_path = os.path.join(os.getcwd(), "universal_parser.py")
            result = subprocess.run(["python", parser_path], cwd=tmpdir, capture_output=True, text=True)

            logging.info(f"Parser output: {result.stdout}")
            if result.returncode != 0:
                raise Exception(result.stderr)

            # ищем итоговый файл
            output_file = os.path.join(os.path.expanduser("~"), "Desktop", "ИТОГ.xlsx")
            if not os.path.exists(output_file):
                output_file = os.path.join(tmpdir, "ИТОГ.xlsx")

            # отправляем пользователю
            await update.message.reply_document(InputFile(output_file, filename="ИТОГ.xlsx"))
            await update.message.reply_text("✅ Готово! Таблица успешно создана и отправлена.")
    except Exception as e:
        logging.error(f"Ошибка: {e}")
        await update.message.reply_text(f"❌ Произошла ошибка при обработке файла: {e}")

# ==========================================================
# ⚙️ Основной запуск
# ==========================================================

async def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_zip))

    print("🤖 Бот запущен и готов к работе...")
    await app.run_polling()

if __name__ == "__main__":
    asyncio.run(main())
