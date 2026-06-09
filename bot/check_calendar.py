import asyncio, os, re
from datetime import datetime, date
import openpyxl
from telegram import Bot

TELEGRAM_TOKEN = os.environ.get('TELEGRAM_TOKEN', '8852326670:AAEItInxXpi-Ynyoc0MvmwDVTZFJtw5o0WI')
GROUP_CHAT_ID = -1003951496246
DATA_FILE = '/opt/tgbot/colleagues.xlsx'

def load_colleagues():
    if not os.path.exists(DATA_FILE):
        return []
    wb = openpyxl.load_workbook(DATA_FILE, data_only=True)
    colleagues = []
    for sheet in wb.worksheets:
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
            colleagues.append({
                'name': str(row[0]).strip(),
                'birthday': row[1] if len(row) > 1 else None,
                'work_anniversary': row[2] if len(row) > 2 else None,
                'children': str(row[3]).strip() if len(row) > 3 and row[3] else '',
                'hobbies': str(row[4]).strip() if len(row) > 4 and row[4] else '',
                'flowers': str(row[5]).strip() if len(row) > 5 and row[5] else '',
                'cake': str(row[6]).strip() if len(row) > 6 and row[6] else '',
            })
    return colleagues

def parse_date(val):
    if not val:
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    try:
        return datetime.strptime(str(val).strip(), '%d.%m.%Y').date()
    except:
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

def make_message(person, event_type, days_left):
    name = person['name']
    flowers = person.get('flowers') or '—'
    cake = person.get('cake') or '—'
    hobbies = person.get('hobbies') or '—'
    if event_type == 'birthday':
        emoji, title = '🎂', f'День народження — {name}'
    elif event_type == 'work':
        emoji, title = '🏆', f'Річниця роботи — {name}'
    else:
        emoji, title = '👶', f'День народження дитини — {name}'
    timing = '🎉 СЬОГОДНІ!' if days_left == 0 else f'⏰ Через {days_left} днів'
    return f'{emoji} {title}\n{timing}\n\n🌸 Квіти: {flowers}\n🍰 Торт: {cake}\n🎯 Хобі: {hobbies}'

async def main():
    bot = Bot(token=TELEGRAM_TOKEN)
    colleagues = load_colleagues()
    for p in colleagues:
        bd = parse_date(p.get('birthday'))
        if bd:
            d = days_until(bd)
            if d in [0, 7]:
                await bot.send_message(chat_id=GROUP_CHAT_ID, text=make_message(p, 'birthday', d))

        wa = parse_date(p.get('work_anniversary'))
        if wa:
            d = days_until(wa)
            if d in [0, 7]:
                await bot.send_message(chat_id=GROUP_CHAT_ID, text=make_message(p, 'work', d))

        children_raw = p.get('children', '')
        if children_raw:
            for d_str in re.findall(r'\d{2}\.\d{2}\.\d{4}', children_raw):
                cd = parse_date(d_str)
                if cd:
                    d = days_until(cd)
                    if d in [0, 7]:
                        await bot.send_message(chat_id=GROUP_CHAT_ID, text=make_message(p, 'child', d))

asyncio.run(main())
