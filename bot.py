import os
import re
import random
import string
import sqlite3
import requests
from dotenv import load_dotenv
import telebot
from telebot import types
from openpyxl import Workbook
import time
from threading import Timer


load_dotenv()

BOT_TOKEN = os.getenv('BOT_TOKEN')
ADMIN_ID = int(os.getenv('ADMIN_ID'))

# Контакты поддержки из .env
SUPPORT_PHONE = os.getenv('SUPPORT_PHONE')        # пример: +7 123 456 78 90
SUPPORT_TELEGRAM = os.getenv('SUPPORT_TELEGRAM')  # пример: @support_manager

# Ссылка на мини-апп МАГАЗИНА. Кнопка «💰 Купить Код» откроет его.
# ?shop в конце = мини-апп открывается сразу в режиме магазина.
SHOP_URL = os.getenv('SHOP_URL', 'https://historical-occasionally-founder-office.trycloudflare.com/?shop')

bot = telebot.TeleBot(BOT_TOKEN)

conn = sqlite3.connect('codes.db', check_same_thread=False)
cursor = conn.cursor()

cursor.execute('''
 CREATE TABLE IF NOT EXISTS codes (
     id INTEGER PRIMARY KEY AUTOINCREMENT,
     code TEXT UNIQUE,
     product TEXT,
     status TEXT,
     instruction TEXT
 )
''')
conn.commit()

# Создаем таблицу пользователей, если её ещё нет
cursor.execute('''
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY
)
''')
conn.commit()

cursor.execute('''
CREATE TABLE IF NOT EXISTS blocked_users (
    user_id INTEGER PRIMARY KEY,
    reason TEXT
)
''')
conn.commit()


user_states = {}
user_data = {}

PAYMENT_TIMEOUT = 15 * 60  # 15 минут на оплату
payment_timers = {}


# ---------- Генерация кода ----------
def generate_code():
   while True:
       raw = ''.join(random.choices(string.digits, k=9))
       code = f"{raw[:3]}-{raw[3:6]}-{raw[6:]}"
       cursor.execute("SELECT id FROM codes WHERE code = ?", (code,))
       if not cursor.fetchone():
           return code

# ---------- Меню ----------
# ---------- Меню ----------
# ---------- Меню ----------
def send_main_menu(chat_id):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)

    # Первая строка
    markup.row(types.KeyboardButton("♻️Активировать Код Robux♻️"))

    # Вторая строка — две кнопки рядом
    # «Купить Код» — обычная кнопка; по нажатию бот пришлёт инлайн-кнопку с web_app
    # (инлайн-кнопка отдаёт подписанный initData, в отличие от reply-кнопки)
    markup.row(
        types.KeyboardButton("💰 Купить Код"),
        types.KeyboardButton("📞Тех Поддержка")
    )

    bot.send_message(chat_id, "Выберите действие:", reply_markup=markup)


def send_admin_menu(chat_id):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)

    # Первая строка
    markup.add(
        types.KeyboardButton("/new"),
        types.KeyboardButton("/addbulk")
    )

    # Вторая строка
    markup.add(
        types.KeyboardButton("/addcode"),
        types.KeyboardButton("/addcodes")
    )

    # Третья строка
    markup.add(
        types.KeyboardButton("/list"),
        types.KeyboardButton("/setinstruction")
    )

    # Четвертая строка
    markup.add(
        types.KeyboardButton("/export"),
        types.KeyboardButton("/msg")
    )

    # Roblox: никнейм -> Place ID
    markup.add(types.KeyboardButton("🎮 Roblox Place ID"))

    # Отдельно кнопка закрытия меню
    markup.add(types.KeyboardButton("Закрыть меню"))

    bot.send_message(chat_id, "⚙️ Админ меню:", reply_markup=markup)


# ---------- Команды ----------
@bot.message_handler(commands=['start'])
def start_handler(message):
   send_main_menu(message.chat.id)

@bot.message_handler(commands=['admin'])
def admin_menu(message):
   if message.from_user.id != ADMIN_ID:
       bot.reply_to(message, "❌ Нет доступа.")
       return
   send_admin_menu(message.chat.id)

@bot.message_handler(commands=['new'])
def new_code(message):
   if message.from_user.id != ADMIN_ID:
       bot.reply_to(message, "❌ Нет доступа.")
       return
   parts = message.text.split("|", maxsplit=1)
   if len(parts) != 2:
       bot.reply_to(message, "Использование:\n/new <товар> | <инструкция>")
       return
   product = parts[0].replace('/new', '').strip()
   instruction = parts[1].strip()
   code = generate_code()
   try:
       cursor.execute(
           "INSERT INTO codes (code, product, status, instruction) VALUES (?, ?, ?, ?)",
           (code, product, 'свободен', instruction)
       )
       conn.commit()
       bot.reply_to(message, f"✅ Код создан:\n🔑 `{code}`\n📦 {product}", parse_mode='Markdown')
   except sqlite3.IntegrityError:
       bot.reply_to(message, "⚠️ Такой код уже существует. Повтори попытку.")

# ---------- Roblox: никнейм -> Place ID ----------
def roblox_get_user(username):
    """Резолвит никнейм Roblox в {id, name, displayName} или None."""
    r = requests.post(
        "https://users.roblox.com/v1/usernames/users",
        json={"usernames": [username], "excludeBannedUsers": False},
        headers={"User-Agent": "Mozilla/5.0"},
        timeout=15,
    )
    r.raise_for_status()
    data = r.json().get("data", [])
    return data[0] if data else None


def roblox_get_games(user_id):
    """Возвращает список игр пользователя Roblox (с пагинацией)."""
    games, cur = [], None
    while True:
        url = f"https://games.roblox.com/v2/users/{user_id}/games?sortOrder=Asc&limit=50"
        if cur:
            url += f"&cursor={cur}"
        r = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=15)
        r.raise_for_status()
        payload = r.json()
        games.extend(payload.get("data", []))
        cur = payload.get("nextPageCursor")
        if not cur:
            break
    return games


def roblox_process_nickname(message, nickname):
    """Формирует и отправляет отчёт по никнейму: User ID + список Place ID."""
    user_states.pop(message.from_user.id, None)
    nickname = nickname.lstrip('@').strip()
    status = bot.reply_to(message, f"🔎 Ищу игры пользователя {nickname}...")
    try:
        user = roblox_get_user(nickname)
        if not user:
            bot.edit_message_text(f"🚫 Пользователь {nickname} не найден в Roblox.",
                                  message.chat.id, status.message_id)
            return

        games = roblox_get_games(user["id"])
        text = (
            f"👤 Никнейм: {user.get('name')}\n"
            f"🏷 Имя: {user.get('displayName')}\n"
            f"🆔 User ID: {user['id']}\n\n"
        )
        if not games:
            text += "📭 Публичных игр (Place ID) нет."
        else:
            text += f"🎮 Игр: {len(games)}\n\n"
            for g in games:
                place_id = (g.get("rootPlace") or {}).get("id", "—")
                text += f"• {g.get('name', 'Без названия')}\n  🔑 Place ID: {place_id}\n"

        try:
            bot.edit_message_text(text, message.chat.id, status.message_id)
        except Exception:
            bot.send_message(message.chat.id, text)
    except Exception as e:
        bot.edit_message_text(f"❌ Ошибка при запросе к Roblox: {e}",
                              message.chat.id, status.message_id)


@bot.message_handler(commands=['roblox'])
def roblox_command(message):
    if message.from_user.id != ADMIN_ID:
        bot.reply_to(message, "❌ Нет доступа.")
        return
    nickname = message.text.replace('/roblox', '', 1).strip()
    if not nickname:
        user_states[message.from_user.id] = 'awaiting_roblox_nick'
        bot.reply_to(message, "🎮 Отправьте никнейм Roblox — пришлю Place ID его игр.")
        return
    roblox_process_nickname(message, nickname)


@bot.message_handler(commands=['block'])
def block_user(message):
    if message.from_user.id != ADMIN_ID:
        return

    try:
        cmd = message.text.replace('/block', '', 1).strip()
        user_id, reason = cmd.split('|', 1)
        user_id = int(user_id.strip())
        reason = reason.strip()

        cursor.execute(
            "INSERT OR REPLACE INTO blocked_users (user_id, reason) VALUES (?, ?)",
            (user_id, reason)
        )
        conn.commit()

        bot.reply_to(message, f"🔒 Пользователь {user_id} заблокирован.\nПричина: {reason}")

        try:
            bot.send_message(
                user_id,
                f"⛔ Доступ к боту заблокирован.\nПричина: {reason}"
            )
        except:
            pass

    except:
        bot.reply_to(
            message,
            "Использование:\n/block <user_id> | <причина>\n\nПример:\n/block 123456 Нарушение правил"
        )
@bot.message_handler(commands=['paylink'])
def send_payment_link(message):
    if message.from_user.id != ADMIN_ID:
        return

    try:
        _, target_id, link = message.text.split(maxsplit=2)
        target_id = int(target_id)
    except ValueError:
        bot.send_message(ADMIN_ID, "❌ Формат:\n/paylink USER_ID ссылка")
        return

    if user_states.get(target_id) != 'waiting_for_payment_link':
        bot.send_message(ADMIN_ID, "❌ Пользователь не ожидает оплату.")
        return

    user_states[target_id] = 'payment_link_sent'

    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    markup.add("✅ Счёт оплачен")
    markup.add("🏠 Главное меню")

    bot.send_message(
        target_id,
        f"💳 Ссылка на оплату:\n{link}\n\n"
        "После оплаты нажмите «Счёт оплачен».",
        reply_markup=markup
    )

    bot.send_message(ADMIN_ID, "✅ Ссылка отправлена пользователю.")

@bot.message_handler(commands=['blocked'])
def list_blocked(message):
    if message.from_user.id != ADMIN_ID:
        return

    cursor.execute("SELECT user_id, reason FROM blocked_users")
    rows = cursor.fetchall()

    if not rows:
        bot.reply_to(message, "✅ Заблокированных пользователей нет.")
        return

    msg = "⛔ Заблокированные пользователи:\n\n"
    for uid, reason in rows:
        msg += f"🆔 {uid} — {reason}\n"

    bot.reply_to(message, msg)

@bot.message_handler(commands=['unblock'])
def unblock_user(message):
    if message.from_user.id != ADMIN_ID:
        return

    try:
        user_id = int(message.text.replace('/unblock', '', 1).strip())

        cursor.execute("DELETE FROM blocked_users WHERE user_id = ?", (user_id,))
        conn.commit()

        bot.reply_to(message, f"🔓 Пользователь {user_id} разблокирован.")

        try:
            bot.send_message(user_id, "✅ Доступ к боту восстановлен.")
        except:
            pass

    except:
        bot.reply_to(message, "Использование:\n/unblock <user_id>")


@bot.message_handler(commands=['addbulk'])
def add_bulk_codes(message):
   if message.from_user.id != ADMIN_ID:
       bot.reply_to(message, "❌ Нет доступа.")
       return
   try:
       cmd = message.text.replace('/addbulk', '', 1).strip()
       left, count_str = cmd.rsplit(' ', 1)
       product, instruction = left.split('|', 1)
       product = product.strip()
       instruction = instruction.strip()
       count = int(count_str)
       if not (1 <= count <= 100):
           raise ValueError
   except:
       bot.reply_to(message, "Использование:\n/addbulk <товар> | <инструкция> <кол-во>\nПример:\n/addbulk VPN | Введите код на vpn.com 5")
       return
   codes = []
   for _ in range(count):
       code = generate_code()
       try:
           cursor.execute(
               "INSERT INTO codes (code, product, status, instruction) VALUES (?, ?, ?, ?)",
               (code, product, 'свободен', instruction)
           )
           conn.commit()
           codes.append(code)
       except sqlite3.IntegrityError:
           pass
   bot.reply_to(message, f"✅ Создано {len(codes)} кодов для '{product}':\n\n" + "\n".join(codes))

@bot.message_handler(commands=['list'])
def list_codes(message):
   if message.from_user.id != ADMIN_ID:
       bot.reply_to(message, "❌ Нет доступа.")
       return
   count = 10
   parts = message.text.split()
   if len(parts) == 2:
       try:
           count = min(int(parts[1]), 50)
       except ValueError:
           bot.reply_to(message, "⚠️ Укажите число: /list 10")
           return
   cursor.execute("SELECT code, product, status FROM codes ORDER BY id DESC LIMIT ?", (count,))
   rows = cursor.fetchall()
   if not rows:
       bot.reply_to(message, "📭 Кодов пока нет.")
       return
   msg = f"📋 Последние {len(rows)} кодов:\n\n"
   for code, product, status in rows:
       emoji = "✅" if status == "использован" else "❎"
       msg += f"{emoji} `{code}` — {product} ({status})\n"
   bot.reply_to(message, msg, parse_mode='Markdown')

@bot.message_handler(commands=['clearcodes'])
def clear_codes(message):
    if message.from_user.id != ADMIN_ID:
        bot.reply_to(message, "❌ Нет доступа.")
        return

    try:
        cursor.execute("DELETE FROM codes")
        conn.commit()
        bot.reply_to(message, "🗑️ Все коды успешно удалены из базы!")
    except Exception as e:
        bot.reply_to(message, f"❌ Ошибка при очистке: {e}")

@bot.message_handler(commands=['setinstruction'])
def set_instruction(message):
   if message.from_user.id != ADMIN_ID:
       bot.reply_to(message, "❌ Нет доступа.")
       return
   try:
       cmd = message.text.replace('/setinstruction', '', 1).strip()
       code, instruction = cmd.split('|', 1)
       code = code.strip()
       instruction = instruction.strip()
       cursor.execute("SELECT id FROM codes WHERE code = ?", (code,))
       if not cursor.fetchone():
           bot.reply_to(message, f"🚫 Код `{code}` не найден.", parse_mode='Markdown')
           return
       cursor.execute("UPDATE codes SET instruction = ? WHERE code = ?", (instruction, code))
       conn.commit()
       bot.reply_to(message, f"✅ Инструкция обновлена для `{code}`.", parse_mode='Markdown')
   except:
       bot.reply_to(message, "Использование:\n/setinstruction <код> | <инструкция>")
# ---------- Проверка статуса кода ----------
@bot.message_handler(commands=['status'])
def check_code_status(message):
    if message.from_user.id != ADMIN_ID:
        bot.reply_to(message, "❌ Нет доступа.")
        return

    try:
        cmd = message.text.replace('/status', '', 1).strip()
        if not cmd:
            bot.reply_to(message, "Использование:\n/status <код>")
            return

        code = cmd
        cursor.execute("SELECT product, status, instruction FROM codes WHERE code = ?", (code,))
        result = cursor.fetchone()

        if not result:
            bot.reply_to(message, f"🚫 Код `{code}` не найден.", parse_mode="Markdown")
            return

        product, status, instruction = result
        bot.reply_to(
            message,
            f"🔎 Информация по коду:\n\n"
            f"🔑 Код: `{code}`\n"
            f"📦 Товар: {product}\n"
            f"📄 Инструкция: {instruction}\n"
            f"📊 Статус: *{status}*",
            parse_mode="Markdown"
        )
    except Exception as e:
        bot.reply_to(message, f"❌ Ошибка: {e}")


# ---------- Установка статуса ----------
@bot.message_handler(commands=['setstatus'])
def set_code_status(message):
    if message.from_user.id != ADMIN_ID:
        bot.reply_to(message, "❌ Нет доступа.")
        return

    try:
        cmd = message.text.replace('/setstatus', '', 1).strip()
        if "|" not in cmd:
            bot.reply_to(message, "Использование:\n/setstatus <код> | <статус>\nПример:\n/setstatus 123-456-789 | свободен")
            return

        code, new_status = cmd.split('|', 1)
        code = code.strip()
        new_status = new_status.strip().lower()

        if new_status not in ['свободен', 'использован']:
            bot.reply_to(message, "❌ Статус может быть только: `свободен` или `использован`", parse_mode="Markdown")
            return

        cursor.execute("SELECT id FROM codes WHERE code = ?", (code,))
        if not cursor.fetchone():
            bot.reply_to(message, f"🚫 Код `{code}` не найден.", parse_mode="Markdown")
            return

        cursor.execute("UPDATE codes SET status = ? WHERE code = ?", (new_status, code))
        conn.commit()
        bot.reply_to(message, f"✅ Статус кода `{code}` обновлён на *{new_status}*.", parse_mode="Markdown")
    except Exception as e:
        bot.reply_to(message, f"❌ Ошибка: {e}")


@bot.message_handler(commands=['addcode'])
def add_code(message):
    if message.from_user.id != ADMIN_ID:
        bot.reply_to(message, "❌ Нет доступа.")
        return
    try:
        cmd = message.text.replace('/addcode', '', 1).strip()
        code, product, instruction = cmd.split('|', 2)
        code = code.strip()
        product = product.strip()
        instruction = instruction.strip()

        # Проверка: есть ли уже такой код
        cursor.execute("SELECT id FROM codes WHERE code = ?", (code,))
        if cursor.fetchone():
            bot.reply_to(message, f"⚠️ Код `{code}` уже существует.", parse_mode='Markdown')
            return

        cursor.execute(
            "INSERT INTO codes (code, product, status, instruction) VALUES (?, ?, ?, ?)",
            (code, product, 'свободен', instruction)
        )
        conn.commit()
        bot.reply_to(message, f"✅ Код `{code}` добавлен.\n📦 {product}", parse_mode='Markdown')
    except:
        bot.reply_to(
            message,
            "Использование:\n/addcode <код> | <товар> | <инструкция>\n\nПример:\n/addcode ABC-123-XYZ | VPN | Ввести на vpn.com")

        # Обработчик команды /addcodes
# ---------- Добавление сразу нескольких кодов ----------
# ---------- Добавление сразу нескольких кодов ----------
@bot.message_handler(commands=['addcodes'])
def add_multiple_codes(message):
    if message.from_user.id != ADMIN_ID:
        bot.reply_to(message, "❌ Нет доступа.")
        return

    try:
        lines = message.text.splitlines()
        if len(lines) < 2:
            bot.reply_to(
                message,
                "❌ Неверный формат.\n\nИспользуй:\n"
                "/addcodes <товар>\n<инструкция>\n\nкод1\nкод2\nкод3..."
            )
            return

        # Первая строка -> /addcodes <товар>
        first_line = lines[0].replace("/addcodes", "", 1).strip()
        if not first_line:
            bot.reply_to(message, "❌ Укажи название товара после команды /addcodes.")
            return
        product_name = first_line

        # Разделяем инструкцию и коды по первой пустой строке
        instruction_lines = []
        code_lines = []
        found_empty = False

        for line in lines[1:]:
            stripped = line.strip()
            if not stripped and not found_empty:
                found_empty = True  # пустая строка = разделитель
                continue
            if not found_empty:
                instruction_lines.append(stripped)
            else:
                if stripped:
                    code_lines.append(stripped)

        instruction = "\n".join(instruction_lines) if instruction_lines else "Нет инструкции"

        if not code_lines:
            bot.reply_to(message, "❌ Не указаны коды для добавления.")
            return

        added = 0
        duplicates = []

        for code in code_lines:
            try:
                cursor.execute(
                    "INSERT INTO codes (code, product, status, instruction) VALUES (?, ?, ?, ?)",
                    (code, product_name, 'свободен', instruction)
                )
                added += 1
            except sqlite3.IntegrityError:
                duplicates.append(code)

        conn.commit()

        # Формируем ответ
        response = f"✅ Добавлено {added} код(ов) для товара *{product_name}*\n📄 Инструкция:\n{instruction}"
        if duplicates:
            response += "\n\n⚠️ Следующие коды уже есть в базе и не были добавлены:\n" + "\n".join(duplicates)

        bot.reply_to(message, response, parse_mode="Markdown")

    except Exception as e:
        bot.reply_to(message, f"❌ Ошибка: {e}")


@bot.message_handler(commands=['msg'])
def send_message_to_user(message):
   if message.from_user.id != ADMIN_ID:
       bot.reply_to(message, "❌ Нет доступа.")
       return
   try:
       cmd = message.text.replace('/msg', '', 1).strip()
       user_id_str, user_msg = cmd.split('|', 1)
       user_id = int(user_id_str.strip())
       user_msg = user_msg.strip()
       bot.send_message(user_id, f"📢 Сообщение от администратора:\n\n{user_msg}")
       bot.reply_to(message, "✅ Сообщение отправлено.")
   except Exception:
       bot.reply_to(message, "Использование:\n/msg <user_id> | <сообщение>\n\nПример:\n/msg 123456789 | Привет!")

@bot.message_handler(commands=['export'])
def export_codes(message):
   if message.from_user.id != ADMIN_ID:
       bot.reply_to(message, "❌ Нет доступа.")
       return

   cursor.execute("SELECT code, product, status, instruction FROM codes")
   rows = cursor.fetchall()

   if not rows:
       bot.reply_to(message, "📭 Нет данных для экспорта.")
       return

   wb = Workbook()
   ws = wb.active
   ws.append(["Код", "Товар", "Статус", "Инструкция"])
   for row in rows:
       ws.append(row)

   file_path = "codes_export.xlsx"
   wb.save(file_path)

   with open(file_path, "rb") as f:
       bot.send_document(message.chat.id, f)

@bot.message_handler(commands=['addusers'])
def add_users(message):
    if message.from_user.id != ADMIN_ID:
        bot.reply_to(message, "❌ Нет доступа.")
        return

    cmd = message.text.replace('/addusers', '', 1).strip()
    if not cmd:
        bot.reply_to(message, "Использование:\n/addusers <user_id1>,<user_id2>,<user_id3>...")
        return

    ids_str = [x.strip() for x in cmd.split(',')]
    added = 0
    for uid_str in ids_str:
        if not uid_str.isdigit():
            continue
        uid = int(uid_str)
        try:
            cursor.execute("INSERT OR IGNORE INTO users (id) VALUES (?)", (uid,))
            conn.commit()
            added += 1
        except Exception as e:
            print(f"Ошибка при добавлении {uid}: {e}")

    bot.reply_to(message, f"✅ Добавлено {added} пользователей.")


@bot.message_handler(commands=['usercount'])
def user_count(message):
    if message.from_user.id != ADMIN_ID:
        bot.reply_to(message, "❌ Нет доступа.")
        return

    cursor.execute("SELECT COUNT(*) FROM users")
    count = cursor.fetchone()[0]

    if count == 0:
        bot.reply_to(message, "👥 Пользователей пока нет.")
    else:
        bot.reply_to(message, f"👥 Всего пользователей: {count}")

# ---------- Рассылка сообщений всем пользователям ----------
@bot.message_handler(commands=['broadcast'])
def broadcast(message):
    if message.from_user.id != ADMIN_ID:
        bot.reply_to(message, "❌ Нет доступа.")
        return

    # Получаем текст и опциональные кнопки
    data = message.text.replace('/broadcast', '', 1).strip()
    if not data:
        bot.reply_to(message,
            "Формат:\n/broadcast Текст | Кнопка | URL\n\nМожно отправлять и просто текст."
        )
        return

    parts = [p.strip() for p in data.split("|")]
    text = parts[0]
    button_text = None
    button_url = None

    if len(parts) == 3:
        button_text = parts[1]
        button_url = parts[2]

    reply_markup = None
    if button_text and button_url:
        keyboard = types.InlineKeyboardMarkup()
        keyboard.add(types.InlineKeyboardButton(button_text, url=button_url))
        reply_markup = keyboard

    # Получаем всех пользователей
    cursor.execute("SELECT id FROM users")
    users = cursor.fetchall()

    sent = 0
    failed = 0
    for (uid,) in users:
        try:
            bot.send_message(uid, text, reply_markup=reply_markup)
            sent += 1
        except:
            failed += 1

    bot.reply_to(message, f"✅ Рассылка отправлена!\n\n📨 Отправлено: {sent}\n⚠️ Ошибок: {failed}")


# ---------- Подсказки по кнопкам для админа ----------
@bot.message_handler(func=lambda message: message.from_user.id == ADMIN_ID)
def admin_buttons_handler(message):
   text = message.text

   # Roblox: пошаговый ввод никнейма после нажатия кнопки
   if user_states.get(message.from_user.id) == 'awaiting_roblox_nick' and text and not text.startswith('/'):
       roblox_process_nickname(message, text.strip())
       return

   if text == "🎮 Roblox Place ID":
       user_states[message.from_user.id] = 'awaiting_roblox_nick'
       bot.send_message(message.chat.id, "🎮 Отправьте никнейм Roblox — пришлю Place ID его игр.")
       return

   if text == "/new":
       bot.send_message(message.chat.id, "Используйте команду:\n/new <товар> | <инструкция>")
   elif text == "/addbulk":
       bot.send_message(message.chat.id, "Используйте команду:\n/addbulk <товар> | <инструкция> <кол-во>")
   elif text == "/list":
       bot.send_message(message.chat.id, "Используйте команду:\n/list <число>")
   elif text == "/setinstruction":
       bot.send_message(message.chat.id, "Используйте команду:\n/setinstruction <код> | <инструкция>")
   elif text == "/msg":
       bot.send_message(message.chat.id, "Используйте команду:\n/msg <user_id> | <сообщение>")
   elif text == "/export":
       bot.send_message(message.chat.id, "Используйте команду:\n/export")
   elif text == "Закрыть меню":
       bot.send_message(message.chat.id, "Меню закрыто.", reply_markup=types.ReplyKeyboardRemove())

# ---------- Обработка всех сообщений ----------
@bot.message_handler(func=lambda message: True)
def all_messages_handler(message):
    user_id = message.from_user.id
    text = (message.text or "").strip()

    # Сохраняем пользователя в базу, если его ещё нет
    cursor.execute("INSERT OR IGNORE INTO users (id) VALUES (?)", (user_id,))
    conn.commit()
        # 🔒 ПРОВЕРКА БЛОКИРОВКИ
    cursor.execute("SELECT reason FROM blocked_users WHERE user_id = ?", (user_id,))
    blocked = cursor.fetchone()
    if blocked:
        if text == "📞Тех Поддержка":
            pass
        else:
            bot.send_message(
                message.chat.id,
                f"⛔ Ваш доступ к боту ограничен.\nПричина: {blocked[0]}"
            )
            return

    # 🔹 Универсальная обработка кнопки "Главное меню"
    if text == "🏠 Главное меню" or text == "Главное меню":
        send_main_menu(message.chat.id)
        user_states.pop(user_id, None)
        user_data.pop(user_id, None)
        return

    # --- Пересылка админу всех входящих сообщений ---
    if user_id != ADMIN_ID:
        try:
            user_info = f"✉️ Сообщение от {message.from_user.first_name or 'Без имени'}"
            if message.from_user.username:
                user_info += f" (@{message.from_user.username})"
            user_info += f"\n🆔 ID: {user_id}"
            bot.send_message(ADMIN_ID, user_info)
            bot.forward_message(ADMIN_ID, message.chat.id, message.message_id)
        except Exception as e:
            print("Ошибка пересылки админу:", e)

    # ====================================================
    # --- АКТИВАЦИЯ КОДА ---
    # ====================================================
    if user_states.get(user_id) == 'waiting_for_code':
        if re.fullmatch(r"\d{3}-\d{3}-\d{3}", text):
            cursor.execute("SELECT product, status, instruction FROM codes WHERE code = ?", (text,))
            result = cursor.fetchone()

            if not result:
                bot.send_message(
                    message.chat.id,
                    "🚫 Код не найден. Попробуйте снова или вернитесь в меню. Напишите @vallmanager если это ошибка",
                    reply_markup=types.ReplyKeyboardRemove()
                )
                user_states.pop(user_id, None)
                send_main_menu(message.chat.id)
                return

            product, status, instruction = result
            if status == 'использован':
                bot.send_message(
                    message.chat.id,
                    "❌ Этот код уже был активирован. Напишите @vallmanager если это ошибка",
                    reply_markup=types.ReplyKeyboardRemove()
                )
                user_states.pop(user_id, None)
                send_main_menu(message.chat.id)
            else:
                user_data[user_id] = {'code': text, 'product': product, 'instruction': instruction}
                user_states[user_id] = 'waiting_for_gamepass_confirmation'

                markup = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
                markup.row("Гейм Пасс Создан")
                markup.row("🏠 Главное меню")
                bot.send_message(
                    message.chat.id,
                    f"✅Код Найден!\n📦Товар: {product}\n🏷Необходимо Создать Game Pass:\n▶️Следуйте Инструкции:\n\n{instruction}\n"
                    "\nНажмите кнопку «Гейм Пасс Создан» после создания:",
                    reply_markup=markup
                )
        else:
            bot.send_message(
                message.chat.id,
                "❗ Код должен быть в формате XXX-XXX-XXX. Попробуйте снова или вернитесь в меню."
            )
        return

    # ====================================================
    # --- ПОДТВЕРЖДЕНИЕ СОЗДАНИЯ ГЕЙМ ПАССА ---
    # ====================================================
    elif user_states.get(user_id) == 'waiting_for_gamepass_confirmation':
        if text == "Гейм Пасс Создан":
            user_states[user_id] = 'waiting_for_nickname'
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
            markup.row("🏠 Главное меню")
            bot.send_message(
                message.chat.id,
                "Отправьте свой игровой никнейм, который начинается с @\n"
                "Где посмотреть свой никнейм: https://telegra.ph/Kak-Uznat-svoj-niknejm-09-04",
                reply_markup=markup
            )
        else:
            bot.send_message(
                message.chat.id,
                "Пожалуйста, нажмите кнопку «Гейм Пасс Создан», если его сделали."
            )
        return

    # ====================================================
    # --- ВВОД НИКНЕЙМА ---
    # ====================================================
    elif user_states.get(user_id) == 'waiting_for_nickname':
        data = user_data.get(user_id)
        if data:
            user_data[user_id]['nickname'] = text
            user_states[user_id] = 'waiting_for_checkbox'

            markup = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
            markup.row("Галочка убрана")
            markup.row("🏠 Главное меню")
            bot.send_message(
                message.chat.id,
                "⚠️ Убедитесь, что вы убрали галочку возле *Enable Regional Pricing* "
                "в настройках Game Pass.\n\nНажмите кнопку «Галочка убрана» после проверки.\n"
                "Подробнее: https://telegra.ph/Galochka-Enable-Regional-pricing-09-04",
                parse_mode="Markdown",
                reply_markup=markup
            )
        else:
            bot.send_message(message.chat.id, "❌ Ошибка. Попробуйте заново.")
        return

    # ====================================================
    # --- ПОДТВЕРЖДЕНИЕ УБРАННОЙ ГАЛОЧКИ ---
    # ====================================================
    elif user_states.get(user_id) == 'waiting_for_checkbox':
        if text == "Галочка убрана":
            data = user_data.get(user_id)
            if data:
                cursor.execute("UPDATE codes SET status = 'использован' WHERE code = ?", (data['code'],))
                conn.commit()

                markup = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
                markup.add("Главное меню")
                bot.send_message(
                    message.chat.id,
                    f"♻️ Гейм Пасс на проверке!\n\n"
                    f"💬 Никнейм: {data['nickname']}\n"
                    f"📦 Товар: {data['product']}\n"
                    f"🔑 Код: {data['code']}\n\n"
                    "⚠️ После обработки придёт уведомление о статусе транзакции.",
                    reply_markup=markup
                )

                # Уведомление админу
                username = f"@{message.from_user.username}" if message.from_user.username else "—"
                admin_msg = (
                    f"📢 Активация завершена!\n\n"
                    f"👤 Пользователь: {message.from_user.first_name or ''} {message.from_user.last_name or ''}\n"
                    f"💬 Username: {username}\n🆔 ID: {user_id}\n\n"
                    f"🔑 Код: {data['code']}\n📦 Товар: {data['product']}\n🎮 Никнейм: {data['nickname']}"
                )
                try:
                    bot.send_message(ADMIN_ID, admin_msg)
                except Exception as e:
                    print("Ошибка отправки уведомления админу:", e)

                user_states.pop(user_id, None)
                user_data.pop(user_id, None)
        else:
            bot.send_message(message.chat.id, "Нажмите кнопку «Галочка убрана», чтобы продолжить.")
        return

    # ====================================================
    # --- ОСНОВНОЕ МЕНЮ И КУПЛЯ ROBUX ---
    # ====================================================
    if text == "♻️Активировать Код Robux♻️":
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        markup.add("🏠 Главное меню")
        bot.send_message(
            message.chat.id,
            "Введите код с карты пополнения в формате XXX-XXX-XXX:",
            reply_markup=markup
        )
        user_states[user_id] = 'waiting_for_code'
        return

    elif text == "💰 Купить Код":
        kb = types.InlineKeyboardMarkup()
        kb.add(types.InlineKeyboardButton(
            "🛒 Открыть магазин",
            web_app=types.WebAppInfo(SHOP_URL)
        ))
        bot.send_message(
            message.chat.id,
            "Нажмите кнопку, чтобы открыть магазин Robux 👇",
            reply_markup=kb
        )
        return

    elif text == "📞Тех Поддержка":
        lines = ["📞 Техническая поддержка:"]
        if SUPPORT_PHONE:
            lines.append(f"📱 WhatsApp: {SUPPORT_PHONE}")
        if SUPPORT_TELEGRAM:
            lines.append(f"💬 Telegram: {SUPPORT_TELEGRAM}")
        if len(lines) == 1:
            lines.append("Просто напишите сюда, и мы поможем!")
        else:
            lines.append("\nНапишите сюда, и мы поможем!")
        bot.send_message(message.chat.id, "\n".join(lines))
        return

    # ====================================================
    # --- ВВОД КОЛИЧЕСТВА ROBUX ---
    # ====================================================
    elif user_states.get(user_id) == 'waiting_for_robux_amount':
        if text.isdigit():
            amount = int(text)
            if 100 <= amount <= 5000:
                total = round(amount * 1.1, 2)
                user_data[user_id] = {'amount': amount, 'total': total}
                user_states[user_id] = 'ready_to_pay'

                markup = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
                markup.add("Оплатить")
                markup.add("🏠 Главное меню")

                bot.send_message(
                    message.chat.id,
                    f"Отлично!\n\n✏️Детали заказа:\n📦 Товар: Код Roblox\n💰 Количество: {amount} Robux\n"
                    f"💵 Сумма к оплате: {total}₽\n\n⚠️ Код будет выслан после оплаты.",
                    reply_markup=markup
                )
            else:
                bot.send_message(message.chat.id, "❌ Количество должно быть от 100 до 5000 Robux.")
        else:
            bot.send_message(message.chat.id, "❌ Введите корректное число Robux (только цифры).")
        return

    # ====================================================
    # --- НАЖАТИЕ КНОПКИ ОПЛАТИТЬ ---
    # ====================================================
    elif user_states.get(user_id) == 'ready_to_pay':
        if text == "Оплатить":
            user_states[user_id] = 'waiting_for_payment_method'

            markup = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
            markup.add("💳 СБП")
            markup.add("🏠 Главное меню")
            bot.send_message(
               message.chat.id,
               "Выберите способ оплаты:",
               reply_markup=markup
            )
        else:
            bot.send_message(
                message.chat.id,
               "Нажмите «Оплатить» или вернитесь в главное меню ⬇️"
            )
        return

    # ====================================================
    # --- ВЫБОР СПОСОБА ОПЛАТЫ ---
    # ====================================================
    elif user_states.get(user_id) == 'waiting_for_payment_method' and text == "💳 СБП":
        data = user_data.get(user_id)
        if not data:
            bot.send_message(message.chat.id, "❌ Ошибка данных заказа.")
            user_states.pop(user_id, None)
            return
        user_states[user_id] = 'waiting_for_payment_link'
        user_data[user_id]['created_at'] = time.time()

        bot.send_message(
           message.chat.id,
           "✅ Заказ создан\n\n"
           "⏳ Ожидайте ссылку на оплату от оператора.\n"
           "После получения ссылки появится кнопка «Счёт оплачен».",
       )

        # ⏳ ТАЙМЕР
        def payment_timeout():
           if user_states.get(user_id) == 'waiting_for_payment_link':
               user_states.pop(user_id, None)
               user_data.pop(user_id, None)
               bot.send_message(
                   message.chat.id,
                   "⌛ Время на оплату истекло. Заказ отменён."
               )

        timer = Timer(PAYMENT_TIMEOUT, payment_timeout)
        timer.start()
        payment_timers[user_id] = timer

        # 📢 УВЕДОМЛЕНИЕ АДМИНУ
        username = f"@{message.from_user.username}" if message.from_user.username else "—"
        admin_msg = (
           f"📢 Новый заказ!\n\n"
           f"👤 {message.from_user.first_name or ''}\n"
           f"💬 Username: {username}\n"
           f"🆔 ID: {user_id}\n"
           f"💰 Robux: {data['amount']}\n"
           f"💵 Сумма: {data['total']} ₽\n\n"
           f"➡️ Отправь ссылку через:\n/paylink {user_id} ССЫЛКА"
        )
        bot.send_message(ADMIN_ID, admin_msg)
        return

    elif user_states.get(user_id) == 'payment_link_sent' and text == "✅ Счёт оплачен":
       user_states[user_id] = 'waiting_for_admin_check'

       if user_id in payment_timers:
          payment_timers[user_id].cancel()
          payment_timers.pop(user_id)

       bot.send_message(
          message.chat.id,
         "⏳ Оплата отмечена.\nПожалуйста Ожидайте проверки оплаты.\nВ случае оплаты счёта вам поступит Код на указанное количество"
       )

       bot.send_message(
          ADMIN_ID,
          f"✅ Пользователь {user_id} нажал «Счёт оплачен».\nПроверь платёж."
       )
       return

    else:
        bot.send_message(message.chat.id, "Выберите действие из меню.", reply_markup=types.ReplyKeyboardRemove())
        send_main_menu(message.chat.id)

# ---------- Запуск ----------
if __name__ == "__main__":
   print("Бот запущен...")
   bot.infinity_polling()
