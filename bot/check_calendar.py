import asyncio, os, re
from datetime import datetime, date
import openpyxl
from telegram import Bot

TELEGRAM_TOKEN = os.environ.get('TELEGRAM_TOKEN', '8852326670:AAEItInxXpi-Ynyoc0MvmwDVTZFJtw5o0WI')
GROUP_CHAT_ID = -1003951496246
DATA_FILE = '/opt/tgbot/colleagues.xlsx'

def parse_date(val):
    if not val:
        return None
    s = str(val).strip().replace(' ', '')
    patterns = [
        (r'^(\d{1,2})\.(\d{1,2})\.(\d{4})$', '%d.%m.%Y'),
        (r'^(\d{1,2})\.(\d{1,2})\.(\d{2})$', '%d.%m.%y'),
    ]
    for pattern, fmt in patterns:
        if re.match(pattern, s):
            try:
                return datetime.strptime(s, fmt).date()
            except:
                pass
    if re.match(r'^(\d{1,2})\.(\d{1,2})$', s):
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

def make_message(event, days_left):
    name = event['name']
    reason = event['reason']
    tg = event['telegram']
    timing = '🎉 СЬОГОДНІ!' if days_left == 0 else f'⏰ Через {days_left} днів'
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

async def send_monthly_preview(bot):
    today = date.today()
    import calendar
    # Наступний місяць
    if today.month == 12:
        next_month, next_year = 1, today.year + 1
    else:
        next_month, next_year = today.month + 1, today.year

    events = load_events()
    found = []
    for e in events:
        try:
            next_event = e['date'].replace(year=next_year, month=next_month)
        except ValueError:
            continue
        if next_event.month == next_month:
            found.append((next_event, e))
    found.sort(key=lambda x: x[0])

    month_name = ['', 'Січень', 'Лютий', 'Березень', 'Квітень', 'Травень', 'Червень',
                  'Липень', 'Серпень', 'Вересень', 'Жовтень', 'Листопад', 'Грудень'][next_month]

    if found:
        lines = [f'📅 Події на {month_name} {next_year}:\n']
        for d, e in found:
            lines.append(f"• {d.strftime('%d.%m')} — {e['reason']} — {e['name']}")
        await bot.send_message(chat_id=GROUP_CHAT_ID, text='\n'.join(lines)[:4000])
    else:
        await bot.send_message(chat_id=GROUP_CHAT_ID, text=f'📅 {month_name} {next_year}: подій немає.')

async def main():
    today = date.today()
    bot = Bot(token=TELEGRAM_TOKEN)

    # Щоденні нагадування
    events = load_events()
    for e in events:
        d = days_until(e['date'])
        if d in [0, 7]:
            await bot.send_message(chat_id=GROUP_CHAT_ID, text=make_message(e, d))

    # 29 або 30 числа — надсилаємо список на наступний місяць
    if today.day in [29, 30]:
        await send_monthly_preview(bot)

asyncio.run(main())
