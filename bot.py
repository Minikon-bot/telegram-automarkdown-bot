import logging
import re
import os
from io import BytesIO
from telegram import Update, InputFile
from telegram.ext import Application, CommandHandler, MessageHandler, ContextTypes, filters
from telegram.ext.webhookhandler import WebhookHandler
from docx import Document
from aiohttp import web

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
    core = escape_punct(original_text.strip())
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
    return '\n'.join(process_paragraph(p) for p in document.paragraphs)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Пришли файл .docx, я отформатирую его в Markdown.")

async def handle_docx(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc_file = update.message.document
    if not doc_file.file_name.endswith('.docx'):
        await update.message.reply_text("Пришли .docx файл.")
        return
    file = await doc_file.get_file()
    doc_bytes = await file.download_as_bytearray()
    formatted = process_document(BytesIO(doc_bytes))
    output = BytesIO()
    output.write(formatted.encode('utf-8'))
    output.seek(0)
    await update.message.reply_document(InputFile(output, 'formatted.txt'), caption="Готово!")

async def main():
    TOKEN = os.getenv("BOT_TOKEN")
    WEBHOOK_URL = os.getenv("WEBHOOK_URL")  # Сюда Render подставит публичный адрес
    PORT = int(os.getenv("PORT", "8080"))
    app = Application.builder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.FileExtension("docx"), handle_docx))
    async def handler(request):
        data = await request.json()
        await app.update_queue.put(Update.de_json(data, app.bot))
        return web.Response()
    app_runner = web.Application()
    app_runner.router.add_post("/", handler)
    await app.initialize()
    await app.bot.set_webhook(WEBHOOK_URL)
    await web._run_app(app_runner, port=PORT)

if __name__ == '__main__':
    import asyncio
    asyncio.run(main())
