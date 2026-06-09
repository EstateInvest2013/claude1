import subprocess, os, asyncio
from datetime import datetime, date
import openpyxl
from telegram import Update, Bot
from telegram.ext import ApplicationBuilder, MessageHandler, CommandHandler, filters, ContextTypes

TELEGRAM_TOKEN = os.environ.get('TELEGRAM_TOKEN', '8852326670:AAEItInxXpi-Ynyoc0MvmwDVTZFJtw5o0WI')
ANTHROPIC_KEY = os.environ.get('ANTHROPIC_API_KEY', '')
GROUP_CHAT_ID = -1003951496246
DATA_FILE = '/opt/tgbot/colleagues.xlsx'

def load_colleagues():
    if not os.path.exists(DATA_FILE):
        return []
    wb = openpyxl.load_workbook(DATA_FILE, data_only=True)
    colleagues = []
    for sheet in wb.worksheets:
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not row[0]:
                continue
            colleagues.append({
                'name': row[0],
                'birthday': row[1],
                'work_anniversary': row[2],
                'children': row[3],
                'hobbies': row[4],
                'flowers': row[5],
                'cake': row[6],
            })
    return colleagues

def parse_date(val):
    if not val:
        return None
    if isinstance(val, (datetime, date)):
        return val if isinstance(val, date) else val.date()
    try:
        return datetime.strptime(str(val).strip(), '%d.%m.%Y').date()
    except:
        return None

def days_until(event_date):
    today = date.today()
    next_event = event_date.replace(year=today.year)
    if next_event < today:
        next_event = next_event.replace(year=today.year + 1)
    return (next_event - today).days

def format_event_message(person, event_type, days_left):
    name = person['name']
    hobbies = person.get('hobbies') or '—'
    flowers = person.get('flowers') or '—'
    cake = person.get('cake') or '—'

    if event_type == 'birthday':
        emoji = '🎂'
        title = f"День народження — {name}"
    elif event_type == 'work_anniversary':
        emoji = '🏆'
        title = f"Річниця роботи — {name}"
    else:
        emoji = '👶'
        title = f"День народження дитини — {name}"

    if days_left == 0:
        timing = "🎉 СЬОГОДНІ!"
    else:
        timing = f"⏰ Через {days_left} днів"

    return (
        f"{emoji} {title}\n"
        f"{timing}\n\n"
        f"🌸 Улюблені квіти: {flowers}\n"
        f"🍰 Улюблений торт: {cake}\n"
        f"🎯 Хобі: {hobbies}"
    )

async def check_and_notify(bot: Bot):
    colleagues = load_colleagues()
    today = date.today()

    for person in colleagues:
        # День народження
        bd = parse_date(person.get('birthday'))
        if bd:
            days = days_until(bd)
            if days in [0, 7]:
                msg = format_event_message(person, 'birthday', days)
                await bot.send_message(chat_id=GROUP_CHAT_ID, text=msg)

        # Річниця роботи
        wa = parse_date(person.get('work_anniversary'))
        if wa:
            days = days_until(wa)
            if days in [0, 7]:
                msg = format_event_message(person, 'work_anniversary', days)
                await bot.send_message(chat_id=GROUP_CHAT_ID, text=msg)

        # Діти
        children_raw = person.get('children')
        if children_raw:
            import re
            dates = re.findall(r'\d{2}\.\d{2}\.\d{4}', str(children_raw))
            for d_str in dates:
                child_date = parse_date(d_str)
                if child_date:
                    days = days_until(child_date)
                    if days in [0, 7]:
                        msg = format_event_message(person, 'child_birthday', days)
                        await bot.send_message(chat_id=GROUP_CHAT_ID, text=msg)

async def daily_check(context):
    await check_and_notify(context.bot)

async def handle(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not update.message:
        return

    chat_id = update.message.chat.id

    if update.message.text == '/id':
        await update.message.reply_text(f"Chat ID: {chat_id}\nТип: {update.message.chat.type}")
        return

    if update.message.text == '/check':
        await update.message.reply_text('Перевіряю календар...')
        await check_and_notify(ctx.bot)
        await update.message.reply_text('Готово! Всі актуальні нагадування надіслано.')
        return

    if update.message.document:
        file_name = update.message.document.file_name
        if file_name.endswith('.xlsx'):
            file = await ctx.bot.get_file(update.message.document.file_id)
            await file.download_to_drive(DATA_FILE)
            colleagues = load_colleagues()
            await update.message.reply_text(
                f'✅ Файл збережено! Завантажено {len(colleagues)} колег.\n\n'
                f'Команди:\n'
                f'/check — перевірити і надіслати актуальні нагадування\n'
                f'/list — показати всіх колег'
            )
            return

    if update.message.text == '/list':
        colleagues = load_colleagues()
        if not colleagues:
            await update.message.reply_text('Список порожній. Завантаж Excel файл.')
            return
        text = '👥 Список колег:\n\n'
        for c in colleagues:
            text += f"• {c['name']}"
            bd = parse_date(c.get('birthday'))
            if bd:
                text += f" — ДН: {bd.strftime('%d.%m')}"
            text += '\n'
        await update.message.reply_text(text[:4000])
        return

    if update.message.text and chat_id != GROUP_CHAT_ID:
        msg = update.message.text
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

# Щоденна перевірка о 9:00
app.job_queue.run_daily(daily_check, time=datetime.strptime('09:00', '%H:%M').time())

app.run_polling()
