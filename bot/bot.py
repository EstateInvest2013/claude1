import subprocess, os
from telegram import Update
from telegram.ext import ApplicationBuilder, MessageHandler, filters, ContextTypes

TELEGRAM_TOKEN = os.environ['TELEGRAM_TOKEN']
ANTHROPIC_KEY = os.environ['ANTHROPIC_API_KEY']

async def handle(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    msg = ''

    if update.message.text:
        msg = update.message.text
    elif update.message.document:
        file = await ctx.bot.get_file(update.message.document.file_id)
        path = f"/tmp/{update.message.document.file_name}"
        await file.download_to_drive(path)
        with open(path, 'r', errors='ignore') as f:
            content = f.read()
        msg = f"Ось файл {update.message.document.file_name}:\n{content[:8000]}"
        if update.message.caption:
            msg += f"\n\nЗавдання: {update.message.caption}"

    if not msg:
        return

    await update.message.reply_text('Виконую...')
    result = subprocess.run(
        ['claude', '-p', msg, '--allowedTools', 'Bash,Read,Write,Edit', '--output-format', 'text'],
        capture_output=True, text=True, timeout=300,
        env={**os.environ, 'ANTHROPIC_API_KEY': ANTHROPIC_KEY}
    )
    reply = result.stdout or result.stderr or 'Немає відповіді'
    for i in range(0, len(reply), 4000):
        await update.message.reply_text(reply[i:i+4000])

app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
app.add_handler(MessageHandler(filters.ALL, handle))
app.run_polling()
