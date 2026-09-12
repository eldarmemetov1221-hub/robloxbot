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

# Заявки на проверку гейм-пасса, ожидающие решения админа.
# Ключ — user_id покупателя, значение — данные заказа (ник, товар, код и т.д.)
pending_orders = {}

# Текст, отправляемый пользователю при одобрении (✅ Успешно)
GAMEPASS_OK_TEXT = (
    "✅ Гейм Пасс Проверен\n\n"
    "✅ Транзакция успешно проведена. \n\n"
    "Проверить транзакцию робуксов можно здесь https://www.roblox.com/transactions \n\n"
    "Робуксы будут отображаться в строке “Pending Robux “ваше количество»\n\n"
    "( Транзакция может отображаться с задержкой до 24ч)\n\n"
    "Робуксы Будут начислены на ваш баланс через 5-7 дней, таковы правила игры.\n\n"
    "⚠️ Важно , не менять цену Гейм Пасс до зачисления Робуксов!\n\n"
    "В случае возникновения проблем или вопросов пишите - @vallmanager\n\n"
    "🛑Важно, после Активации кода товар возврату не Подлежит Попытки оформить возврат "
    "товара после активации кода рассматриваются как мошеннические действия и могут "
    "повлечь ответственность в соответствии с действующим законодательством и правилами площадки."
)

# Готовые причины отклонения (❌ Ошибка). Индекс кнопки -> (краткое, шаблон текста).
# В шаблонах доступны переменные {expected} (нужная цена) и {actual} (текущая цена).
GAMEPASS_REASONS = [
    (
        "Неверная цена",
        "❌ Транзакция не выполнена! \n"
        "Причина: Не найден Гейм Пасс с ценой {expected} на данном Аккаунте.\n\n"
        "Из-за Комиссии Игры 30% Цена Гейм Пасс должна быть {expected}! Не {actual}\n\n"
        "Измените цену на {expected} Робуксов в настройках Гейм Пасс или сделайте новый с Нужной ценой \n"
        "И активируйте код повторно\n\n"
        "Или Обратитесь к менеджеру - @vallmanager или @vallmanager1 \n"
        "Для уточнения и решения.",
    ),
    (
        "Нет публичного Place",
        "❌ Гейм Пасс Не проверен\n"
        "❌ Транзакция не выполнена! \n"
        "Причина:Не найден Place\n\n"
        "Place должен быть Публичным! Текущий place Приватный\n\n"
        "Измените настройки Place сделайте его Публичным и активируйте код повторно\n\n"
        "Или Обратитесь к менеджеру - @vallmanager или @vallmanager1 \n"
        "Для уточнения и решения.",
    ),
    (
        "Пасс не создан",
        "❌ Гейм Пасс Не проверен\n"
        "❌ Транзакция не выполнена! \n"
        "Причина: Не найден Гейм Пасс на данном Аккаунте.\n\n"
        "Создайте Гейм Пасс по инструкции и активируйте код повторно\n\n"
        "Или Обратитесь к менеджеру - @vallmanager или @vallmanager1 \n"
        "Для уточнения и решения.",
    ),
    (
        "Не убрана Regional Pricing",
        "❌ Гейм Пасс Не проверен\n"
        "❌ Транзакция не выполнена! \n"
        "Причина: Не убрана галочка Enable Regional Pricing\n\n"
        "Уберите галочку Enable Regional Pricing в настройках Гейм Пасс и активируйте код повторно\n"
        "Подробнее: https://telegra.ph/Galochka-Enable-Regional-pricing-09-04\n\n"
        "Или Обратитесь к менеджеру - @vallmanager или @vallmanager1 \n"
        "Для уточнения и решения.",
    ),
]


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
    """Возвращает список опубликованных игр (Experiences) пользователя."""
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


def roblox_get_inventory_places(user_id):
    """Возвращает Place-ассеты из инвентаря пользователя (тип 9).

    None -> инвентарь скрыт настройками приватности (403).
    """
    places, cur = [], None
    while True:
        url = f"https://inventory.roblox.com/v2/users/{user_id}/inventory/9?limit=100&sortOrder=Asc"
        if cur:
            url += f"&cursor={cur}"
        r = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=15)
        if r.status_code == 403:
            return None
        r.raise_for_status()
        payload = r.json()
        places.extend(payload.get("data", []))
        cur = payload.get("nextPageCursor")
        if not cur:
            break
    return places


def roblox_get_gamepass_price(gamepass_id):
    """Возвращает цену гейм-пасса в Robux или None, если недоступна/не на продаже."""
    try:
        r = requests.get(
            f"https://apis.roblox.com/game-passes/v1/game-passes/{gamepass_id}/product-info",
            headers={"User-Agent": "Mozilla/5.0"},
            timeout=15,
        )
        if r.status_code != 200:
            return None
        info = r.json()
        # Возможные варианты формата ответа
        if info.get("PriceInRobux") is not None:
            return info.get("PriceInRobux")
        price_info = info.get("PriceInformation") or {}
        if price_info.get("DefaultPriceInRobux") is not None:
            return price_info.get("DefaultPriceInRobux")
        return info.get("price")
    except Exception:
        return None


def roblox_get_gamepasses(user_id):
    """Список гейм-пассов пользователя: [{id, name, price}]. None -> ошибка запроса."""
    try:
        result = []
        cur = None
        while True:
            url = f"https://apis.roblox.com/game-passes/v1/users/{user_id}/game-passes?count=100"
            if cur:
                url += f"&cursor={cur}"
            r = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=15)
            r.raise_for_status()
            payload = r.json()
            items = payload.get("gamePasses") or payload.get("data") or []
            for it in items:
                gp_id = it.get("gamePassId") or it.get("id")
                if gp_id is None:
                    continue
                price = it.get("price")
                if price is None:
                    price = it.get("priceInRobux")
                result.append({
                    "id": gp_id,
                    "name": it.get("name", "Без названия"),
                    "price": price,
                })
            cur = payload.get("nextPageCursor") or payload.get("cursor")
            if not cur:
                break
        return result
    except Exception:
        return None


def parse_robux_amount(product):
    """Пытается вытащить количество Robux из названия товара (первое число)."""
    m = re.search(r"\d+", product or "")
    return int(m.group()) if m else None


def expected_gamepass_price(robux_amount):
    """Ожидаемая цена гейм-пасса, чтобы продавец получил robux_amount Robux.

    Roblox берёт 30% комиссии: продавец получает floor(цена * 0.7).
    Нужна минимальная цена P, при которой floor(P*0.7) >= robux_amount,
    что равно ceil(10*R/7) = (10*R + 6) // 7 (целочисленно, без ошибок float).
    Совпадает с таблицей: 100->143, 400->572, 700->1000, 1000->1429.
    """
    if not robux_amount:
        return None
    return (10 * robux_amount + 6) // 7


def build_gamepass_report(nickname, expected_price=None, robux_amount=None):
    """Возвращает (текст с инфой, user_id или None, фактическая_цена или None).

    Фактическая цена — цена гейм-пасса, ближайшего к ожидаемой (лучшая догадка о
    том, какую цену выставил пользователь) — подставляется в {actual} шаблона ошибки.
    """
    nick = (nickname or "").lstrip('@').strip()
    try:
        user = roblox_get_user(nick)
    except Exception as e:
        return f"⚠️ Не удалось найти пользователя `{nick}`: {e}", None, None

    if not user:
        return f"⚠️ Пользователь `{nick}` не найден в Roblox.", None, None

    user_id = user["id"]
    header = f"🔗 Профиль: https://www.roblox.com/users/{user_id}/profile\n"
    if expected_price is not None:
        header += f"🎯 Ожидаемая цена пасса: {expected_price} R$ (за {robux_amount} Robux)\n"

    passes = roblox_get_gamepasses(user_id)
    if passes is None:
        return header + "⚠️ Не удалось автоматически получить гейм-пассы (проверьте вручную).", user_id, None
    if not passes:
        return header + "📭 Гейм-пассы у пользователя не найдены.", user_id, None

    total = len(passes)
    # Если пассов много — показываем ближайшие к ожидаемой цене (чтобы влезть в лимит Telegram)
    MAX_SHOW = 15
    if expected_price is not None:
        passes = sorted(
            passes,
            key=lambda p: abs((p.get("price") if p.get("price") is not None else 10 ** 9) - expected_price),
        )
    shown = passes[:MAX_SHOW]

    detected_price = None  # цена пасса, ближайшая к ожидаемой
    lines = [header + f"🎟 Гейм-пассы ({total}):"]
    for p in shown:
        price = p.get("price")
        # цена может отсутствовать в списке — дозапрашиваем
        if price is None:
            price = roblox_get_gamepass_price(p["id"])

        if price is None:
            price_str = "не на продаже/скрыта"
        else:
            price_str = f"{price} R$"
            if expected_price is not None:
                if price == expected_price:
                    price_str += " ✅ цена верна"
                elif abs(price - expected_price) <= 1:
                    price_str += " ⚠️ почти совпадает"
                else:
                    price_str += " ❌ не совпадает"
            # запоминаем цену, ближайшую к ожидаемой (или первую известную)
            if expected_price is None:
                if detected_price is None:
                    detected_price = price
            elif detected_price is None or abs(price - expected_price) < abs(detected_price - expected_price):
                detected_price = price

        lines.append(
            f"• {p['name']}\n"
            f"  💰 Цена: {price_str}\n"
            f"  🔗 https://www.roblox.com/game-pass/{p['id']}"
        )

    if total > MAX_SHOW:
        lines.append(f"… и ещё {total - MAX_SHOW} пасс(ов) (показаны ближайшие к нужной цене)")

    return "\n".join(lines), user_id, detected_price


def build_admin_order_keyboard(uid):
    kb = types.InlineKeyboardMarkup()
    kb.row(
        types.InlineKeyboardButton("✅ Успешно", callback_data=f"gp_ok:{uid}"),
        types.InlineKeyboardButton("❌ Ошибка", callback_data=f"gp_err:{uid}"),
    )
    return kb


def build_reasons_keyboard(uid):
    kb = types.InlineKeyboardMarkup()
    for idx, (short, _) in enumerate(GAMEPASS_REASONS):
        kb.add(types.InlineKeyboardButton(f"❌ {short}", callback_data=f"gp_r:{idx}:{uid}"))
    kb.add(types.InlineKeyboardButton("⬅️ Назад", callback_data=f"gp_back:{uid}"))
    return kb


def notify_admin_new_order(message, data):
    """Отправляет админу заявку с авто-инфо по гейм-пассам и inline-кнопками."""
    uid = message.from_user.id
    username = f"@{message.from_user.username}" if message.from_user.username else "—"
    fullname = f"{message.from_user.first_name or ''} {message.from_user.last_name or ''}".strip()

    robux_amount = parse_robux_amount(data['product'])
    exp_price = expected_gamepass_price(robux_amount)
    gp_text, _, actual_price = build_gamepass_report(
        data['nickname'], expected_price=exp_price, robux_amount=robux_amount
    )

    pending_orders[uid] = {
        'nickname': data['nickname'],
        'product': data['product'],
        'code': data['code'],
        'username': username,
        'fullname': fullname,
        'expected_price': exp_price,
        'actual_price': actual_price,
    }

    admin_msg = (
        f"📢 Новая заявка на проверку!\n\n"
        f"👤 {fullname or '—'}\n"
        f"💬 Username: {username}\n"
        f"🆔 ID: {uid}\n\n"
        f"🔑 Код: {data['code']}\n"
        f"📦 Товар: {data['product']}\n"
        f"🎮 Никнейм: {data['nickname']}\n\n"
        f"{gp_text}"
    )
    # Telegram не принимает сообщения длиннее 4096 символов
    if len(admin_msg) > 4000:
        admin_msg = admin_msg[:3950] + "\n\n… (список обрезан, слишком длинный)"
    try:
        bot.send_message(
            ADMIN_ID, admin_msg,
            reply_markup=build_admin_order_keyboard(uid),
            disable_web_page_preview=True,
        )
    except Exception as e:
        print("Ошибка отправки заявки админу:", e)


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

        user_id = user["id"]
        games = roblox_get_games(user_id)
        inv_places = roblox_get_inventory_places(user_id)

        text = (
            f"👤 Никнейм: {user.get('name')}\n"
            f"🏷 Имя: {user.get('displayName')}\n"
            f"🆔 User ID: {user_id}\n"
        )

        # Place ID из опубликованных игр (Experiences), чтобы не дублировать в инвентаре
        game_place_ids = set()

        if games:
            text += f"\n🎮 Игры (Experiences): {len(games)}\n"
            for g in games:
                place_id = (g.get("rootPlace") or {}).get("id", "—")
                if place_id != "—":
                    game_place_ids.add(place_id)
                text += f"• {g.get('name', 'Без названия')}\n  🔑 Place ID: {place_id}\n"

        # Place ID из инвентаря (в т.ч. Offsale / неопубликованные)
        if inv_places is None:
            text += "\n🔒 Инвентарь скрыт настройками приватности пользователя (Places не проверить)."
        else:
            extra = [p for p in inv_places if p.get("assetId") not in game_place_ids]
            if extra:
                text += f"\n📁 Места из инвентаря: {len(extra)}\n"
                for p in extra:
                    text += f"• {p.get('name', 'Без названия')}\n  🔑 Place ID: {p.get('assetId', '—')}\n"

        if not games and not inv_places:
            text += "\n📭 Публичных игр и мест (Place ID) не найдено."

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


# ---------- Обработка inline-кнопок проверки гейм-пасса ----------
@bot.callback_query_handler(func=lambda c: c.data and c.data.startswith("gp_"))
def gamepass_decision_handler(call):
    if call.from_user.id != ADMIN_ID:
        bot.answer_callback_query(call.id, "❌ Нет доступа.")
        return

    parts = call.data.split(":")
    action = parts[0]

    try:
        if action == "gp_err":
            uid = int(parts[1])
            bot.edit_message_reply_markup(
                call.message.chat.id, call.message.message_id,
                reply_markup=build_reasons_keyboard(uid),
            )
            bot.answer_callback_query(call.id, "Выберите причину отклонения")
            return

        if action == "gp_back":
            uid = int(parts[1])
            bot.edit_message_reply_markup(
                call.message.chat.id, call.message.message_id,
                reply_markup=build_admin_order_keyboard(uid),
            )
            bot.answer_callback_query(call.id)
            return

        if action == "gp_ok":
            uid = int(parts[1])
            try:
                bot.send_message(uid, GAMEPASS_OK_TEXT, disable_web_page_preview=True)
            except Exception as e:
                bot.answer_callback_query(call.id, f"Не отправлено пользователю: {e}")
                return
            pending_orders.pop(uid, None)
            bot.edit_message_text(
                (call.message.text or "") + "\n\n✅ ОДОБРЕНО — пользователю отправлено уведомление.",
                call.message.chat.id, call.message.message_id,
                reply_markup=None, disable_web_page_preview=True,
            )
            bot.answer_callback_query(call.id, "✅ Одобрено")
            return

        if action == "gp_r":
            idx = int(parts[1])
            uid = int(parts[2])
            short, template = GAMEPASS_REASONS[idx]
            order = pending_orders.get(uid, {})

            # Подставляем нужную и фактическую цену (для причины «Неверная цена»)
            expected = order.get('expected_price')
            actual = order.get('actual_price')
            user_text = template.format(
                expected=expected if expected is not None else "нужную",
                actual=actual if actual is not None else "—",
            )

            # Возвращаем код в свободные, чтобы пользователь мог повторить активацию
            code = order.get('code')
            if code:
                cursor.execute("UPDATE codes SET status = 'свободен' WHERE code = ?", (code,))
                conn.commit()

            try:
                bot.send_message(uid, user_text, disable_web_page_preview=True)
            except Exception as e:
                bot.answer_callback_query(call.id, f"Не отправлено пользователю: {e}")
                return
            pending_orders.pop(uid, None)
            freed = " Код возвращён в свободные." if code else ""
            bot.edit_message_text(
                (call.message.text or "") + f"\n\n❌ ОТКЛОНЕНО ({short}).{freed}",
                call.message.chat.id, call.message.message_id,
                reply_markup=None, disable_web_page_preview=True,
            )
            bot.answer_callback_query(call.id, "❌ Отклонено")
            return

    except Exception as e:
        try:
            bot.answer_callback_query(call.id, f"Ошибка: {e}")
        except Exception:
            pass


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

                # Уведомление админу с авто-инфо по пассам и inline-кнопками
                notify_admin_new_order(message, data)

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
