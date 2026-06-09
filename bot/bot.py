import subprocess, os, re
from datetime import datetime, date
import openpyxl
from telegram import Update
from telegram.ext import ApplicationBuilder, MessageHandler, filters, ContextTypes

TELEGRAM_TOKEN = os.environ.get('TELEGRAM_TOKEN', '8852326670:AAEItInxXpi-Ynyoc0MvmwDVTZFJtw5o0WI')
ANTHROPIC_KEY = os.environ.get('ANTHROPIC_API_KEY', '')
GROUP_CHAT_ID = -1003951496246
DATA_FILE = '/opt/tgbot/colleagues.xlsx'

def parse_date(val):
    if not val:
        return None
    s = str(val).strip().replace(' ', '')
    # Формати: 01.06, 01.06.96, 01.06.1996, 01.06.2024
    patterns = [
        (r'^(\d{1,2})\.(\d{1,2})\.(\d{4})$', '%d.%m.%Y'),
        (r'^(\d{1,2})\.(\d{1,2})\.(\d{2})$', '%d.%m.%y'),
        (r'^(\d{1,2})\.(\d{1,2})$', None),  # тільки день.місяць
    ]
    for pattern, fmt in patterns:
        if re.match(pattern, s):
            if fmt:
                try:
                    return datetime.strptime(s, fmt).date()
                except:
                    pass
            else:
                # тільки день.місяць — додаємо поточний рік
                try:
                    parts = s.split('.')
                    return date(date.today().year, int(parts[1]), int(parts[0]))
                except:
                    pass
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    return None

def days_until(d):
    today = date.today()
    try:
        next_event = d.replace(year=today.year)
    except ValueError:
        next_event = d.replace(year=today.year, day=28)
    if next_event < today:
        next_event = next_event.replace(year=today.year + 1)
    return (next_event - today).days

def load_events():
    if not os.path.exists(DATA_FILE):
        return []
    wb = openpyxl.load_workbook(DATA_FILE, data_only=True)
    events = []
    for sheet in wb.worksheets:
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
            name = str(row[0]).strip() if row[0] else ''
            reason = str(row[4]).strip() if len(row) > 4 and row[4] else ''
            date_val = row[5] if len(row) > 5 else None
            d = parse_date(date_val)
            if name and reason and d:
                events.append({
                    'name': name,
                    'reason': reason,
                    'date': d,
                    'telegram': str(row[3]).strip() if len(row) > 3 and row[3] else '',
                })
    return events

def make_message(event, days_left):
    name = event['name']
    reason = event['reason']
    tg = event['telegram']
    timing = '🎉 СЬОГОДНІ!' if days_left == 0 else f'⏰ Через {days_left} днів'

    # Визначаємо емодзі за приводом
    r = reason.lower()
    if 'народження' in r and ('доньк' in r or 'сина' in r or 'дитин' in r):
        emoji = '👶'
    elif 'народження' in r:
        emoji = '🎂'
    elif 'річниця' in r or 'заснування' in r:
        emoji = '🏆'
    else:
        emoji = '🎊'

    msg = f'{emoji} {reason} — {name}\n{timing}'
    if tg:
        msg += f'\n📱 {tg}'
    return msg

async def do_check(bot, notify_days=[0, 7]):
    events = load_events()
    count = 0
    for e in events:
        d = days_until(e['date'])
        if d in notify_days:
            await bot.send_message(chat_id=GROUP_CHAT_ID, text=make_message(e, d))
            count += 1
    return count

async def handle(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not update.message:
        return

    msg_text = update.message.text or ''
    chat_id = update.message.chat.id

    if msg_text == '/id':
        await update.message.reply_text(f'Chat ID: {chat_id}')
        return

    if msg_text.startswith('/delete '):
        name_to_delete = msg_text[8:].strip()
        if not os.path.exists(DATA_FILE):
            await update.message.reply_text('Файл не знайдено.')
            return
        wb = openpyxl.load_workbook(DATA_FILE)
        deleted = 0
        for sheet in wb.worksheets:
            rows_to_delete = []
            for row in sheet.iter_rows(min_row=2):
                if row[0].value and name_to_delete.lower() in str(row[0].value).lower():
                    rows_to_delete.append(row[0].row)
            for row_num in reversed(rows_to_delete):
                sheet.delete_rows(row_num)
                deleted += 1
        wb.save(DATA_FILE)
        if deleted:
            await update.message.reply_text(f'✅ Видалено {deleted} запис(ів) для "{name_to_delete}".')
        else:
            await update.message.reply_text(f'❌ Не знайдено "{name_to_delete}" у файлі.')
        return

    if msg_text.startswith('/add'):
        text = msg_text[4:].strip()
        if not text:
            await update.message.reply_text(
                'Надішли інформацію про нового колегу після /add, наприклад:\n\n'
                '/add\n'
                'Іванова Марія Петрівна\n'
                'Дата народження: 18.05.1986\n'
                'Посада: Агент\n'
                'Офіс: Київ\n'
                'Річниця роботи: 01.06.2024\n'
                'Діти: немає\n'
                'Хобі: йога, читання\n'
                'Квіти: троянди\n'
                'Торт: медовик\n'
                'Телеграм: @username'
            )
            return
        await update.message.reply_text('Обробляю інформацію...')
        prompt = f"""Розбери інформацію про нового співробітника і поверни JSON з полями:
- name (ПІБ або ім'я)
- birthday (дата народження у форматі DD.MM.YYYY, або null)
- work_anniversary (річниця роботи у форматі DD.MM.YYYY, або null)
- children (інформація про дітей або null)
- hobbies (хобі або null)
- flowers (улюблені квіти або null)
- cake (улюблений торт або null)
- telegram (телеграм нікі через кому або null)
- position (посада або null)
- office (офіс або null)

Інформація:
{text}

Поверни ТІЛЬКИ JSON без пояснень."""

        result = subprocess.run(
            ['claude', '-p', prompt, '--output-format', 'text'],
            capture_output=True, text=True, timeout=60,
            env={**os.environ, 'ANTHROPIC_API_KEY': ANTHROPIC_KEY}
        )
        import json
        try:
            raw = result.stdout.strip()
            raw = re.sub(r'^```json\s*', '', raw)
            raw = re.sub(r'^```\s*', '', raw)
            raw = re.sub(r'\s*```$', '', raw)
            data = json.loads(raw)
        except:
            await update.message.reply_text(f'❌ Не вдалось розібрати інформацію.\n{result.stdout[:500]}')
            return

        if not os.path.exists(DATA_FILE):
            await update.message.reply_text('❌ Файл не знайдено. Спочатку завантаж Excel.')
            return

        wb = openpyxl.load_workbook(DATA_FILE)
        ws = wb.worksheets[0]

        def fmt_date(d):
            if not d:
                return None
            try:
                return datetime.strptime(str(d), '%d.%m.%Y').date()
            except:
                return None

        bd = fmt_date(data.get('birthday'))
        wa = fmt_date(data.get('work_anniversary'))
        tg = data.get('telegram') or ''
        name = data.get('name', '')

        rows_added = 0
        if bd:
            ws.append([name, data.get('office',''), data.get('position',''), tg,
                       f"день народження {name.split()[1] if len(name.split())>1 else name}", bd])
            rows_added += 1
        if wa:
            ws.append([name, data.get('office',''), data.get('position',''), tg,
                       'річниця роботи в ЕІ', wa])
            rows_added += 1

        wb.save(DATA_FILE)

        summary = f'✅ Додано {name}!\n\n'
        summary += f'📅 День народження: {data.get("birthday") or "не вказано"}\n'
        summary += f'🏆 Річниця роботи: {data.get("work_anniversary") or "не вказано"}\n'
        summary += f'🌸 Квіти: {data.get("flowers") or "не вказано"}\n'
        summary += f'🍰 Торт: {data.get("cake") or "не вказано"}\n'
        summary += f'🎯 Хобі: {data.get("hobbies") or "не вказано"}\n'
        summary += f'📱 Телеграм: {tg or "не вказано"}'
        await update.message.reply_text(summary)
        return

    if msg_text == '/check':
        await update.message.reply_text('Перевіряю календар...')
        count = await do_check(ctx.bot)
        await update.message.reply_text(f'Готово! Надіслано {count} нагадувань.')
        return

    if msg_text == '/today':
        await update.message.reply_text('Події на сьогодні і через 7 днів:')
        events = load_events()
        found = []
        for e in events:
            d = days_until(e['date'])
            if d in [0, 7]:
                found.append(make_message(e, d))
        if found:
            for msg in found:
                await update.message.reply_text(msg)
        else:
            await update.message.reply_text('Немає подій на найближчі 7 днів.')
        return

    if msg_text == '/list':
        events = load_events()
        if not events:
            await update.message.reply_text('Список порожній. Завантаж Excel файл.')
            return
        lines = [f'📋 Всього подій: {len(events)}\n']
        for e in events:
            lines.append(f"• {e['name']} — {e['reason']} ({e['date'].strftime('%d.%m')})")
        text = '\n'.join(lines)
        for i in range(0, len(text), 4000):
            await update.message.reply_text(text[i:i+4000])
        return

    if update.message.document:
        fname = update.message.document.file_name or ''
        if fname.endswith('.xlsx'):
            await update.message.reply_text('Завантажую файл...')
            file = await ctx.bot.get_file(update.message.document.file_id)
            await file.download_to_drive(DATA_FILE)
            events = load_events()
            await update.message.reply_text(
                f'✅ Файл збережено! Знайдено {len(events)} подій.\n\n'
                f'Команди:\n'
                f'/today — найближчі 7 днів\n'
                f'/check — надіслати нагадування в групу\n'
                f'/list — всі події'
            )
            return

    if msg_text and chat_id != GROUP_CHAT_ID:
        await update.message.reply_text('Виконую...')
        result = subprocess.run(
            ['claude', '-p', msg_text, '--allowedTools', 'Bash,Read,Write,Edit', '--output-format', 'text'],
            capture_output=True, text=True, timeout=120,
            env={**os.environ, 'ANTHROPIC_API_KEY': ANTHROPIC_KEY}
        )
        reply = result.stdout or result.stderr or 'Немає відповіді'
        for i in range(0, len(reply), 4000):
            await update.message.reply_text(reply[i:i+4000])

app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
app.add_handler(MessageHandler(filters.ALL, handle))
app.run_polling()
