import logging
import os
import re
from io import BytesIO

from telegram import Update, InputFile
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
from docx import Document

# Логирование
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)

# Экранируемые символы
ESCAPE_CHARS = r'."!?)([]%:;\-'

def escape_punct(text: str) -> str:
    return re.sub(f'([{re.escape(ESCAPE_CHARS)}])', r'\\\1', text)

def format_run(run):
    text = run.text
    if not text.strip():
        return escape_punct(text)

    leading_spaces = len(text) - len(text.lstrip())
    trailing_spaces = len(text) - len(text.rstrip())
    leading = text[:leading_spaces]
    trailing = text[len(text.rstrip()):]
    core = text.strip()

    core = escape_punct(core)

    # Стилизация
    try:
        strike = run.font.strike
    except AttributeError:
        strike = False

    # 1. Подчёркнутый + курсив → только подчёркнутый
    if run.underline and run.italic:
        core = f"__{core}__"
    else:
        if run.bold:
            core = f"*{core}*"
        if run.italic:
            core = f"_{core}_"
        if run.underline:
            core = f"__{core}__"
        if strike:
            core = f"~{core}~"

    return f"{leading}{core}{trailing}"

def process_paragraph(paragraph):
    return ''.join(format_run(run) for run in paragraph.runs)

def process_document(docx_bytes):
    doc = Document(docx_bytes)
    lines = [process_paragraph(p) for p in doc.paragraphs]
    return '\n'.join(lines)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Присылай .docx файл — я отформатирую текст с Markdown-разметкой (жирный, курсив, подчёркнутый)."
    )

async def handle_docx(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc_file = update.message.document
    if not doc_file.file_name.endswith('.docx'):
        await update.message.reply_text("Пришли файл с расширением .docx.")
        return

    file = await doc_file.get_file()
    content = await file.download_as_bytearray()

    formatted = process_document(BytesIO(content))

    output = BytesIO()
    output.write(formatted.encode('utf-8'))
    output.seek(0)

    await update.message.reply_document(
        document=InputFile(output, filename="formatted.txt"),
        caption="Готово ✅"
    )

def main():
    TOKEN = os.getenv("BOT_TOKEN")
    WEBHOOK_URL = os.getenv("WEBHOOK_URL")
    PORT = int(os.environ.get("PORT", 10000))

    if not TOKEN or not WEBHOOK_URL:
        raise RuntimeError("Не заданы переменные окружения BOT_TOKEN или WEBHOOK_URL!")

    app = ApplicationBuilder().token(TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.FileExtension("docx"), handle_docx))

    print("Бот запускается через вебхук...")

    app.run_webhook(
        listen="0.0.0.0",
        port=PORT,
        webhook_url=WEBHOOK_URL,
    )

if __name__ == '__main__':
    main()
