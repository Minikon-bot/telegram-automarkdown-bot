import logging
import os
import re
from io import BytesIO

from telegram import Update, InputFile
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    ContextTypes,
    filters,
)
from docx import Document

# Логирование
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO
)

ESCAPE_CHARS = r'."!?)([]%:;\-'

def escape_punct(text: str) -> str:
    return re.sub(f'([{re.escape(ESCAPE_CHARS)}])', r'\\\1', text)

def format_run(run):
    original_text = run.text
    if not original_text.strip():
        return escape_punct(original_text)

    leading_spaces = len(original_text) - len(original_text.lstrip())
    trailing_spaces = len(original_text) - len(original_text.rstrip())
    leading = original_text[:leading_spaces]
    trailing = original_text[len(original_text.rstrip()):]
    core = original_text.strip()

    core = escape_punct(core)

    try:
        strike = run.font.strike
    except AttributeError:
        strike = False

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
    document = Document(docx_bytes)
    result_lines = [process_paragraph(para) for para in document.paragraphs]
    return '\n'.join(result_lines)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Присылай файл .docx, а я верну тебе Markdown-текст со стилями."
    )

async def handle_docx(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc_file = update.message.document
    if not doc_file.file_name.endswith('.docx'):
        await update.message.reply_text("Пожалуйста, пришли .docx файл.")
        return

    file = await doc_file.get_file()
    doc_bytes = await file.download_as_bytearray()

    formatted_text = process_document(BytesIO(doc_bytes))

    output = BytesIO()
    output.write(formatted_text.encode('utf-8'))
    output.seek(0)

    await update.message.reply_document(
        document=InputFile(output, filename='formatted.txt'),
        caption="Вот ваш Markdown-текст"
    )

async def main():
    TOKEN = os.getenv("BOT_TOKEN")
    APP_URL = os.getenv("APP_URL")  # например, https://yourapp.onrender.com

    if not TOKEN or not APP_URL:
        raise RuntimeError("BOT_TOKEN и APP_URL должны быть установлены!")

    app = ApplicationBuilder().token(TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.FileExtension("docx"), handle_docx))

    # Запуск webhook (асинхронно)
    await app.run_webhook(
        listen="0.0.0.0",
        port=int(os.environ.get("PORT", 10000)),
        webhook_url=f"{APP_URL}/webhook",
    )

if __name__ == '__main__':
    import asyncio
    asyncio.run(main())
