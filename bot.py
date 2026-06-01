#!/usr/bin/env python3
"""
Telegram-бот «Большой день женского здоровья» с Еленой Пшинник
Хранение данных: SQLite (локальный файл bot.db)
"""

import asyncio
import logging
import io
import sqlite3
from datetime import datetime, timedelta
from contextlib import contextmanager

from telegram import (
    Update, InlineKeyboardButton, InlineKeyboardMarkup,
    ReplyKeyboardMarkup, ReplyKeyboardRemove, KeyboardButton
)
from telegram.ext import (
    Application, CommandHandler, MessageHandler, CallbackQueryHandler,
    ConversationHandler, filters, ContextTypes
)
from telegram.constants import ParseMode

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# ═══════════════════════════════════════════════════════════
#  НАСТРОЙКИ — заполни перед запуском
# ═══════════════════════════════════════════════════════════
BOT_TOKEN          = "8757283175:AAEy1joRPQl-QfFJ84QvtgqW1vgXaormSjg"           # токен от @BotFather
ADMIN_IDS          = [334618540]                # твой Telegram user_id

WORKSHOP_DATE_STR  = "01 июня 2025, 12:00"   # дата для показа
WORKSHOP_DATETIME  = datetime(2025, 7, 15, 10, 0)
WORKSHOP_PRICE     = 10000
WORKSHOP_LOCATION  = "Ссылка появится за день до воркшопа"

QR_IMAGE_PATH      = "qr_sbp.png"
DB_PATH            = "bot.db"

# Реквизиты ИП
IP_FIO             = "Пшинник Елена Борисовна"
IP_INN             = "622601705505"             # ← вставь ИНН
IP_OGRNIP          = "1027739609391"          # ← вставь ОГРНИП
IP_EMAIL           = "elena-pshinnik@mail.ru"
IP_ADDRESS         = "107031, Г. МОСКВА, г Москва, г МОСКВА, УЛ РОЖДЕСТВЕНКА, 10/2, СТР 1"

# ═══════════════════════════════════════════════════════════
#  СОСТОЯНИЯ
# ═══════════════════════════════════════════════════════════
(
    S_CONSENT, S_FULLNAME, S_PHONE, S_EMAIL,
    S_PAYMENT, S_SCREENSHOT,
    S_SUPPORT,
    S_BROADCAST_TARGET, S_BROADCAST_TEXT,
) = range(9)

# ═══════════════════════════════════════════════════════════
#  БАЗА ДАННЫХ (SQLite)
# ═══════════════════════════════════════════════════════════
@contextmanager
def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()

def init_db():
    with db() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS participants (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id     INTEGER UNIQUE,
                username    TEXT,
                full_name   TEXT,
                phone       TEXT,
                email       TEXT,
                status      TEXT DEFAULT 'Ожидает подтверждения',
                amount      TEXT DEFAULT '',
                reg_date    TEXT,
                pay_date    TEXT DEFAULT '',
                note        TEXT DEFAULT ''
            )
        """)

def db_add(data: dict) -> bool:
    try:
        with db() as conn:
            conn.execute("""
                INSERT OR IGNORE INTO participants
                (user_id, username, full_name, phone, email, reg_date)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (
                data["user_id"], data.get("username", ""),
                data["full_name"], data["phone"], data["email"],
                datetime.now().strftime("%d.%m.%Y %H:%M")
            ))
        return True
    except Exception as e:
        logging.error(f"DB add error: {e}"); return False

def db_confirm(user_id: int, amount: str) -> bool:
    try:
        with db() as conn:
            conn.execute("""
                UPDATE participants
                SET status=?, amount=?, pay_date=?
                WHERE user_id=?
            """, ("Оплачено", amount, datetime.now().strftime("%d.%m.%Y %H:%M"), user_id))
        return True
    except Exception as e:
        logging.error(f"DB confirm error: {e}"); return False

def db_all() -> list:
    with db() as conn:
        rows = conn.execute("SELECT * FROM participants ORDER BY id").fetchall()
        return [dict(r) for r in rows]

def db_get(user_id: int) -> dict | None:
    with db() as conn:
        row = conn.execute("SELECT * FROM participants WHERE user_id=?", (user_id,)).fetchone()
        return dict(row) if row else None

def db_paid_ids() -> list[int]:
    with db() as conn:
        rows = conn.execute("SELECT user_id FROM participants WHERE status='Оплачено'").fetchall()
        return [r["user_id"] for r in rows]

def db_all_ids() -> list[int]:
    with db() as conn:
        rows = conn.execute("SELECT user_id FROM participants").fetchall()
        return [r["user_id"] for r in rows]

def db_stats() -> dict:
    with db() as conn:
        total  = conn.execute("SELECT COUNT(*) FROM participants").fetchone()[0]
        paid   = conn.execute("SELECT COUNT(*) FROM participants WHERE status='Оплачено'").fetchone()[0]
        amount = conn.execute("SELECT SUM(CAST(amount AS INTEGER)) FROM participants WHERE status='Оплачено'").fetchone()[0] or 0
        return {"total": total, "paid": paid, "pending": total - paid, "amount": amount}

# ═══════════════════════════════════════════════════════════
#  EXCEL ЭКСПОРТ
# ═══════════════════════════════════════════════════════════
def make_excel(records: list) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Участники"

    # Заголовки
    headers = ["№", "Дата регистрации", "ФИО", "Телефон", "Email",
               "Статус", "Сумма, ₽", "Дата оплаты", "Telegram ID", "Username"]
    header_fill = PatternFill("solid", fgColor="C6EFCE")
    header_font = Font(bold=True, color="276221")
    thin = Side(style="thin", color="AAAAAA")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    ws.row_dimensions[1].height = 30

    # Данные
    paid_fill   = PatternFill("solid", fgColor="EBF5E6")
    pend_fill   = PatternFill("solid", fgColor="FFF9E6")

    for i, r in enumerate(records, 1):
        is_paid = r.get("status") == "Оплачено"
        row_fill = paid_fill if is_paid else pend_fill
        vals = [
            i,
            r.get("reg_date", ""),
            r.get("full_name", ""),
            r.get("phone", ""),
            r.get("email", ""),
            r.get("status", ""),
            r.get("amount", ""),
            r.get("pay_date", ""),
            str(r.get("user_id", "")),
            r.get("username", ""),
        ]
        for col, val in enumerate(vals, 1):
            cell = ws.cell(row=i+1, column=col, value=val)
            cell.fill = row_fill
            cell.border = border
            cell.alignment = Alignment(vertical="center")

    # Ширина столбцов
    widths = [5, 18, 30, 18, 28, 22, 10, 18, 14, 16]
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[ws.cell(1, col).column_letter].width = w

    # Итоговая строка
    last = len(records) + 2
    stats = db_stats()
    ws.cell(last, 1, "Итого:")
    ws.cell(last, 1).font = Font(bold=True)
    ws.cell(last, 6, f"Оплатили: {stats['paid']} / {stats['total']}")
    ws.cell(last, 7, stats["amount"])
    ws.cell(last, 7).font = Font(bold=True)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

# ═══════════════════════════════════════════════════════════
#  ТЕКСТЫ
# ═══════════════════════════════════════════════════════════
T_WELCOME = """
🌸 *Добро пожаловать!*

Вы регистрируетесь на воркшоп
*«Большой день женского здоровья»*
с Еленой Пшинник

📅 Дата: *{date}*
💰 Стоимость: *{price} ₽*

Перед регистрацией необходимо ознакомиться с документами.
""".strip()

T_DOCS = """
📄 *Документы*

Ознакомьтесь перед регистрацией:

Нажимая *«✅ Согласна»*, вы подтверждаете:
— согласие с публичной офертой
— согласие на обработку персональных данных (ФЗ №152-ФЗ)
""".strip()

T_FULLNAME = "✏️ *Введите ваше ФИО*\n\nФамилия Имя Отчество — полностью.\n_Пример: Иванова Мария Сергеевна_"
T_PHONE    = "📱 *Введите номер телефона*\n\nФормат: +7XXXXXXXXXX\n\n_Или нажмите кнопку ниже_"
T_EMAIL    = "📧 *Введите электронную почту*\n\n_Пример: name@mail.ru_"

T_PAYMENT = """
💳 *Оплата участия*

Стоимость: *{price} ₽*

Отсканируйте QR-код для оплаты через СБП — подходит любой российский банк.

После оплаты нажмите *«✅ Я оплатила»* и пришлите скриншот из банка.
""".strip()

T_SCREENSHOT = "📸 *Пришлите скриншот оплаты*\n\nСделайте скриншот успешного платежа и отправьте сюда.\n\n_На скриншоте должны быть видны: сумма, дата и статус «Успешно»_"

T_SUCCESS = """
✅ *Заявка принята!*

Спасибо, {first_name}!

📋 *Ваши данные:*
👤 {full_name}
📱 {phone}
📧 {email}

📅 Воркшоп: *{date}*

После проверки оплаты вы получите подтверждение.
_Вопросы — кнопка «💬 Поддержка» в меню._
""".strip()

T_ADMIN_NEW = """
🔔 *Новая заявка!*

👤 {full_name}
📱 {phone}
📧 {email}
🕐 {time}
🆔 `{user_id}` @{username}

✅ Подтвердить оплату:
`/confirm_{user_id}_{price}`
""".strip()

T_RECEIPT = """
🧾 *ЧЕК ОБ ОПЛАТЕ*
━━━━━━━━━━━━━━━━━━━━
ИП {ip_fio}
ИНН: {ip_inn} | ОГРНИП: {ip_ogrnip}

📌 Услуга: участие в онлайн-воркшопе
«Большой день женского здоровья»

👤 {full_name}
📅 Дата воркшопа: {workshop_date}
💰 Сумма: {amount} ₽
📆 Дата оплаты: {pay_date}

✅ *Оплачено*
━━━━━━━━━━━━━━━━━━━━
_Сохраните чек_
""".strip()

T_REMINDER = """
🌸 *Напоминание!*

Завтра — воркшоп
*«Большой день женского здоровья»*
с Еленой Пшинник 💫

📅 *{date}*
📍 {location}

До встречи! 🙏
""".strip()

T_OFFER_FULL = """
📋 *ПУБЛИЧНАЯ ОФЕРТА*
━━━━━━━━━━━━━━━━━━━━

*Индивидуальный предприниматель Пшинник Елена Борисовна*
ИНН: {inn} | ОГРНИП: {ogrnip}
Адрес: {address}
E-mail: {email}

*1. ПРЕДМЕТ ОФЕРТЫ*
Организатор проводит онлайн-воркшоп «Большой день женского здоровья» и предоставляет Участнику доступ к нему.
Дата: {workshop_date}. Формат: онлайн-трансляция в Telegram.

*2. АКЦЕПТ*
Акцептом является заполнение регистрационной формы и оплата. Акцепт = полное согласие с офертой.

*3. СТОИМОСТЬ И ОПЛАТА*
Стоимость: {price} рублей. Оплата через СБП до начала воркшопа.

*4. ОБЯЗАТЕЛЬСТВА СТОРОН*
Организатор: провести воркшоп, предоставить ссылку, вернуть оплату при отмене по вине организатора.
Участник: предоставить достоверные данные, не записывать и не распространять материалы.

*5. ВОЗВРАТ*
Возврат при отказе за 3+ дня до воркшопа. При отказе менее чем за 3 дня — оплата не возвращается. При отмене организатором — полный возврат в течение 10 рабочих дней.

*6. ОТВЕТСТВЕННОСТЬ*
Организатор не несёт ответственности за технические сбои на стороне Участника.

*7. ПЕРСОНАЛЬНЫЕ ДАННЫЕ*
Участник даёт согласие на обработку персональных данных согласно Политике конфиденциальности.

ИП Пшинник Елена Борисовна
━━━━━━━━━━━━━━━━━━━━
""".strip()

T_PRIVACY_FULL = """
🔒 *ПОЛИТИКА ОБРАБОТКИ ПЕРСОНАЛЬНЫХ ДАННЫХ*
━━━━━━━━━━━━━━━━━━━━

*ИП Пшинник Елена Борисовна*
ИНН: {inn} | E-mail: {email}

*1. ОБЩИЕ ПОЛОЖЕНИЯ*
Политика разработана в соответствии с ФЗ №152-ФЗ «О персональных данных».

*2. ЦЕЛИ ОБРАБОТКИ*
— Идентификация участника при регистрации
— Направление информации о воркшопе
— Отправка чека об оплате
— Направление напоминаний и организационных сообщений

*3. ОБРАБАТЫВАЕМЫЕ ДАННЫЕ*
— Фамилия, имя, отчество
— Номер телефона
— Адрес электронной почты
— Идентификатор Telegram

*4. ПРАВОВОЕ ОСНОВАНИЕ*
Согласие субъекта персональных данных (ст. 6 ч. 1 п. 1 ФЗ-152).

*5. ПОРЯДОК ОБРАБОТКИ*
Данные хранятся в защищённой базе данных на сервере. Передача третьим лицам без согласия не осуществляется.

*6. СРОК ХРАНЕНИЯ*
1 год с момента проведения воркшопа.

*7. ПРАВА СУБЪЕКТА*
Вы вправе запросить, уточнить или удалить свои данные. Обращайтесь: {email}

ИП Пшинник Елена Борисовна
━━━━━━━━━━━━━━━━━━━━
""".strip()

# ═══════════════════════════════════════════════════════════
#  ВСПОМОГАТЕЛЬНОЕ
# ═══════════════════════════════════════════════════════════
def is_admin(uid): return uid in ADMIN_IDS

def main_kb():
    return ReplyKeyboardMarkup(
        [["📋 Моя регистрация", "💬 Поддержка"],
         ["📄 Оферта",          "🔒 Политика данных"]],
        resize_keyboard=True
    )

# ═══════════════════════════════════════════════════════════
#  /start
# ═══════════════════════════════════════════════════════════
async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> int:
    arg = ctx.args[0] if ctx.args else ""
    if arg == "offer":   return await _send_offer(update)
    if arg == "privacy": return await _send_privacy(update)
    ctx.user_data.clear()
    kb = [[InlineKeyboardButton("📄 Ознакомиться с документами →", callback_data="show_docs")]]
    await update.message.reply_text(
        T_WELCOME.format(date=WORKSHOP_DATE_STR, price=WORKSHOP_PRICE),
        parse_mode=ParseMode.MARKDOWN, reply_markup=InlineKeyboardMarkup(kb)
    )
    return S_CONSENT

async def _send_offer(update):
    text = T_OFFER_FULL.format(inn=IP_INN, ogrnip=IP_OGRNIP, address=IP_ADDRESS,
                                email=IP_EMAIL, workshop_date=WORKSHOP_DATE_STR, price=WORKSHOP_PRICE)
    for chunk in [text[i:i+4000] for i in range(0, len(text), 4000)]:
        await update.message.reply_text(chunk, parse_mode=ParseMode.MARKDOWN)
    return S_CONSENT

async def _send_privacy(update):
    text = T_PRIVACY_FULL.format(inn=IP_INN, email=IP_EMAIL)
    for chunk in [text[i:i+4000] for i in range(0, len(text), 4000)]:
        await update.message.reply_text(chunk, parse_mode=ParseMode.MARKDOWN)
    return S_CONSENT

# ═══════════════════════════════════════════════════════════
#  РЕГИСТРАЦИЯ
# ═══════════════════════════════════════════════════════════
async def cb_show_docs(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> int:
    q = update.callback_query; await q.answer()
    kb = [
        [InlineKeyboardButton("📋 Публичная оферта",               callback_data="read_offer")],
        [InlineKeyboardButton("🔒 Политика персональных данных",   callback_data="read_privacy")],
        [InlineKeyboardButton("✅ Согласна со всеми условиями",    callback_data="consent_yes")],
        [InlineKeyboardButton("❌ Не согласна",                    callback_data="consent_no")],
    ]
    await q.edit_message_text(T_DOCS, parse_mode=ParseMode.MARKDOWN,
                              reply_markup=InlineKeyboardMarkup(kb))
    return S_CONSENT

async def cb_read_offer(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> int:
    q = update.callback_query; await q.answer()
    text = T_OFFER_FULL.format(inn=IP_INN, ogrnip=IP_OGRNIP, address=IP_ADDRESS,
                                email=IP_EMAIL, workshop_date=WORKSHOP_DATE_STR, price=WORKSHOP_PRICE)
    for chunk in [text[i:i+4000] for i in range(0, len(text), 4000)]:
        await q.message.reply_text(chunk, parse_mode=ParseMode.MARKDOWN)
    kb = [[InlineKeyboardButton("✅ Согласна", callback_data="consent_yes"),
           InlineKeyboardButton("❌ Не согласна", callback_data="consent_no")]]
    await q.message.reply_text("Продолжить?", reply_markup=InlineKeyboardMarkup(kb))
    return S_CONSENT

async def cb_read_privacy(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> int:
    q = update.callback_query; await q.answer()
    text = T_PRIVACY_FULL.format(inn=IP_INN, email=IP_EMAIL)
    for chunk in [text[i:i+4000] for i in range(0, len(text), 4000)]:
        await q.message.reply_text(chunk, parse_mode=ParseMode.MARKDOWN)
    kb = [[InlineKeyboardButton("✅ Согласна", callback_data="consent_yes"),
           InlineKeyboardButton("❌ Не согласна", callback_data="consent_no")]]
    await q.message.reply_text("Продолжить?", reply_markup=InlineKeyboardMarkup(kb))
    return S_CONSENT

async def cb_consent_yes(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> int:
    q = update.callback_query; await q.answer()
    ctx.user_data["user_id"]  = q.from_user.id
    ctx.user_data["username"] = q.from_user.username or ""
    await q.edit_message_text("✅ Согласие получено! Начинаем регистрацию.")
    await q.message.reply_text(T_FULLNAME, parse_mode=ParseMode.MARKDOWN)
    return S_FULLNAME

async def cb_consent_no(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> int:
    q = update.callback_query; await q.answer()
    await q.edit_message_text("❌ Без согласия регистрация невозможна.\n\nЕсли передумаете — /start")
    return ConversationHandler.END

async def rx_fullname(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> int:
    text = update.message.text.strip()
    if len(text.split()) < 2:
        await update.message.reply_text("⚠️ Введите полное ФИО (минимум имя и фамилию).")
        return S_FULLNAME
    ctx.user_data["full_name"] = text
    kb = [[KeyboardButton("📱 Отправить мой номер", request_contact=True)]]
    await update.message.reply_text(
        T_PHONE, parse_mode=ParseMode.MARKDOWN,
        reply_markup=ReplyKeyboardMarkup(kb, resize_keyboard=True, one_time_keyboard=True)
    )
    return S_PHONE

async def rx_phone(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> int:
    if update.message.contact:
        phone = update.message.contact.phone_number
        if not phone.startswith("+"): phone = "+" + phone
    else:
        phone = update.message.text.strip()
        clean = phone.replace("+","").replace("-","").replace(" ","").replace("(","").replace(")","")
        if not clean.isdigit() or len(clean) < 10:
            await update.message.reply_text("⚠️ Введите корректный номер (+79991234567).")
            return S_PHONE
    ctx.user_data["phone"] = phone
    await update.message.reply_text(T_EMAIL, parse_mode=ParseMode.MARKDOWN,
                                    reply_markup=ReplyKeyboardRemove())
    return S_EMAIL

async def rx_email(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> int:
    email = update.message.text.strip().lower()
    if "@" not in email or "." not in email.split("@")[-1]:
        await update.message.reply_text("⚠️ Введите корректный email.")
        return S_EMAIL
    ctx.user_data["email"] = email
    kb = [[InlineKeyboardButton("✅ Я оплатила", callback_data="payment_done")]]
    await update.message.reply_text(
        T_PAYMENT.format(price=WORKSHOP_PRICE),
        parse_mode=ParseMode.MARKDOWN, reply_markup=InlineKeyboardMarkup(kb)
    )
    try:
        with open(QR_IMAGE_PATH, "rb") as f:
            await update.message.reply_photo(photo=f,
                caption=f"📲 QR-код СБП для оплаты {WORKSHOP_PRICE} ₽")
    except FileNotFoundError:
        await update.message.reply_text("⚠️ QR-код не найден — свяжитесь с организатором.")
    return S_PAYMENT

async def cb_payment_done(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> int:
    q = update.callback_query; await q.answer()
    await q.message.reply_text(T_SCREENSHOT, parse_mode=ParseMode.MARKDOWN)
    return S_SCREENSHOT

async def rx_screenshot(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> int:
    if not update.message.photo and not update.message.document:
        await update.message.reply_text(
            "⚠️ Пришлите *скриншот* (фото из галереи).", parse_mode=ParseMode.MARKDOWN)
        return S_SCREENSHOT

    file_id = (update.message.photo[-1].file_id if update.message.photo
               else update.message.document.file_id)
    data = ctx.user_data
    user = update.message.from_user
    parts = data["full_name"].split()
    first = parts[1] if len(parts) > 1 else parts[0]

    db_add({"user_id": user.id, "username": data.get("username",""),
            "full_name": data["full_name"], "phone": data["phone"], "email": data["email"]})

    await update.message.reply_text(
        T_SUCCESS.format(first_name=first, full_name=data["full_name"],
                         phone=data["phone"], email=data["email"], date=WORKSHOP_DATE_STR),
        parse_mode=ParseMode.MARKDOWN, reply_markup=main_kb()
    )

    # Уведомляем администратора
    for aid in ADMIN_IDS:
        try:
            await ctx.bot.send_message(
                chat_id=aid,
                text=T_ADMIN_NEW.format(
                    full_name=data["full_name"], phone=data["phone"], email=data["email"],
                    time=datetime.now().strftime("%d.%m.%Y %H:%M"),
                    user_id=user.id, username=data.get("username","—"), price=WORKSHOP_PRICE
                ), parse_mode=ParseMode.MARKDOWN
            )
            caption = f"💳 Скриншот от {data['full_name']}"
            if update.message.photo:
                await ctx.bot.send_photo(chat_id=aid, photo=file_id, caption=caption)
            else:
                await ctx.bot.send_document(chat_id=aid, document=file_id, caption=caption)
        except Exception as e:
            logging.error(f"Admin notify {aid}: {e}")

    # Напоминание за день
    remind_at = WORKSHOP_DATETIME - timedelta(days=1)
    if remind_at > datetime.now():
        ctx.job_queue.run_once(
            job_reminder, when=remind_at,
            data={"user_id": user.id},
            name=f"reminder_{user.id}"
        )
    return ConversationHandler.END

async def job_reminder(ctx: ContextTypes.DEFAULT_TYPE):
    uid = ctx.job.data["user_id"]
    try:
        await ctx.bot.send_message(
            chat_id=uid,
            text=T_REMINDER.format(date=WORKSHOP_DATE_STR, location=WORKSHOP_LOCATION),
            parse_mode=ParseMode.MARKDOWN
        )
    except Exception as e:
        logging.error(f"Reminder {uid}: {e}")

# ═══════════════════════════════════════════════════════════
#  ПОДДЕРЖКА
# ═══════════════════════════════════════════════════════════
async def start_support(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> int:
    kb = [[InlineKeyboardButton("❌ Отмена", callback_data="support_cancel")]]
    await update.message.reply_text(
        "💬 *Напишите ваш вопрос* — отвечу в ближайшее время.",
        parse_mode=ParseMode.MARKDOWN, reply_markup=InlineKeyboardMarkup(kb)
    )
    return S_SUPPORT

async def rx_support(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> int:
    user = update.effective_user
    text = update.message.text or "[вложение]"
    for aid in ADMIN_IDS:
        try:
            await ctx.bot.send_message(
                chat_id=aid,
                text=f"💬 *Вопрос поддержки*\n\n"
                     f"👤 {user.full_name} (@{user.username or '—'})\n"
                     f"🆔 `{user.id}`\n\n"
                     f"{text}\n\n"
                     f"_Ответить: `/reply_{user.id} текст`_",
                parse_mode=ParseMode.MARKDOWN
            )
        except Exception as e:
            logging.error(f"Support fwd {aid}: {e}")
    await update.message.reply_text("✅ Вопрос отправлен! Отвечу в ближайшее время 🙏",
                                    reply_markup=main_kb())
    return ConversationHandler.END

async def cb_support_cancel(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> int:
    q = update.callback_query; await q.answer()
    await q.edit_message_text("Возврат в меню.")
    return ConversationHandler.END

# ═══════════════════════════════════════════════════════════
#  МЕНЮ
# ═══════════════════════════════════════════════════════════
async def menu_my_reg(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    rec = db_get(update.effective_user.id)
    if not rec:
        await update.message.reply_text(
            "Вы ещё не зарегистрированы.\nНажмите /start чтобы начать.",
            reply_markup=main_kb()); return
    e = "✅" if rec["status"] == "Оплачено" else "⏳"
    await update.message.reply_text(
        f"📋 *Ваша регистрация*\n\n"
        f"👤 {rec['full_name']}\n📱 {rec['phone']}\n📧 {rec['email']}\n\n"
        f"{e} Статус: *{rec['status']}*\n📅 {WORKSHOP_DATE_STR}",
        parse_mode=ParseMode.MARKDOWN, reply_markup=main_kb()
    )

async def menu_offer(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    text = T_OFFER_FULL.format(inn=IP_INN, ogrnip=IP_OGRNIP, address=IP_ADDRESS,
                                email=IP_EMAIL, workshop_date=WORKSHOP_DATE_STR, price=WORKSHOP_PRICE)
    for chunk in [text[i:i+4000] for i in range(0, len(text), 4000)]:
        await update.message.reply_text(chunk, parse_mode=ParseMode.MARKDOWN)

async def menu_privacy(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    text = T_PRIVACY_FULL.format(inn=IP_INN, email=IP_EMAIL)
    for chunk in [text[i:i+4000] for i in range(0, len(text), 4000)]:
        await update.message.reply_text(chunk, parse_mode=ParseMode.MARKDOWN)

# ═══════════════════════════════════════════════════════════
#  ADMIN
# ═══════════════════════════════════════════════════════════
async def cmd_confirm(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id): return
    try:
        _, uid_s, amount = update.message.text.split("_", 2)
        user_id = int(uid_s)
    except ValueError:
        await update.message.reply_text("Формат: `/confirm_123456789_10000`",
                                        parse_mode=ParseMode.MARKDOWN); return
    db_confirm(user_id, amount)
    rec = db_get(user_id)
    full_name = rec["full_name"] if rec else "Участник"
    receipt = T_RECEIPT.format(
        ip_fio=IP_FIO, ip_inn=IP_INN, ip_ogrnip=IP_OGRNIP,
        full_name=full_name, workshop_date=WORKSHOP_DATE_STR,
        amount=amount, pay_date=datetime.now().strftime("%d.%m.%Y %H:%M")
    )
    try:
        await ctx.bot.send_message(chat_id=user_id, text=receipt, parse_mode=ParseMode.MARKDOWN)
        await update.message.reply_text(f"✅ Подтверждено. Чек отправлен — {full_name}.")
    except Exception as e:
        await update.message.reply_text(f"⚠️ Не удалось отправить чек: {e}")

async def cmd_reply(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id): return
    try:
        parts = update.message.text.split(" ", 1)
        user_id = int(parts[0].split("_")[1])
        text = parts[1] if len(parts) > 1 else ""
    except (IndexError, ValueError):
        await update.message.reply_text("Формат: `/reply_123456789 Текст ответа`",
                                        parse_mode=ParseMode.MARKDOWN); return
    try:
        await ctx.bot.send_message(chat_id=user_id,
                                   text=f"💬 *Ответ организатора:*\n\n{text}",
                                   parse_mode=ParseMode.MARKDOWN)
        await update.message.reply_text("✅ Ответ отправлен.")
    except Exception as e:
        await update.message.reply_text(f"⚠️ Ошибка: {e}")

async def cmd_participants(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("⛔ Нет доступа."); return
    records = db_all()
    if not records:
        await update.message.reply_text("📋 Участников пока нет."); return
    st = db_stats()
    lines = [f"📋 *Участники* — {st['total']} чел.\n"
             f"✅ Оплатили: {st['paid']} | ⏳ Ожидают: {st['pending']} | 💰 {st['amount']} ₽\n"
             f"{'━'*20}"]
    for r in records:
        e = "✅" if r["status"] == "Оплачено" else "⏳"
        amt = f" | {r['amount']} ₽" if r.get("amount") else ""
        lines.append(f"{e} *{r['id']}.* {r['full_name']}\n"
                     f"   📱 {r['phone']}  📧 {r['email']}{amt}")
    text = "\n".join(lines)
    for chunk in [text[i:i+4000] for i in range(0, len(text), 4000)]:
        await update.message.reply_text(chunk, parse_mode=ParseMode.MARKDOWN)
    kb = [[InlineKeyboardButton("📥 Скачать Excel", callback_data="export_excel")]]
    await update.message.reply_text("Выгрузить в файл?", reply_markup=InlineKeyboardMarkup(kb))

async def cb_export_excel(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query; await q.answer()
    if not is_admin(q.from_user.id): return
    records = db_all()
    if not records:
        await q.message.reply_text("Нет данных."); return
    if not HAS_OPENPYXL:
        await q.message.reply_text("⚠️ Установи openpyxl: pip install openpyxl"); return
    data = make_excel(records)
    fname = f"participants_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    await q.message.reply_document(
        document=io.BytesIO(data), filename=fname,
        caption="📊 Список участников воркшопа"
    )

async def cmd_export(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Быстрый экспорт командой /export"""
    if not is_admin(update.effective_user.id): return
    records = db_all()
    if not records:
        await update.message.reply_text("Нет данных."); return
    if not HAS_OPENPYXL:
        await update.message.reply_text("⚠️ Установи openpyxl: pip install openpyxl"); return
    data = make_excel(records)
    fname = f"participants_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    await update.message.reply_document(
        document=io.BytesIO(data), filename=fname,
        caption=f"📊 Участники воркшопа — {datetime.now().strftime('%d.%m.%Y')}"
    )

async def cmd_stats(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id): return
    st = db_stats()
    await update.message.reply_text(
        f"📊 *Статистика*\n\n"
        f"👥 Заявок: *{st['total']}*\n"
        f"✅ Оплачено: *{st['paid']}*\n"
        f"⏳ Ожидают: *{st['pending']}*\n"
        f"💰 Собрано: *{st['amount']} ₽*\n\n"
        f"📅 {WORKSHOP_DATE_STR}",
        parse_mode=ParseMode.MARKDOWN
    )

async def cmd_admin_help(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id): return
    await update.message.reply_text(
        "🛠 *Команды администратора*\n\n"
        "/participants — список участников\n"
        "/export — скачать Excel\n"
        "/stats — статистика\n"
        "/broadcast — рассылка сообщения\n"
        "/confirm\\_ID\\_СУММА — подтвердить оплату + отправить чек\n"
        "/reply\\_ID текст — ответить участнице\n\n"
        "_Примеры:_\n"
        "`/confirm_123456789_10000`\n"
        "`/reply_123456789 Добрый день!`",
        parse_mode=ParseMode.MARKDOWN
    )

# ═══════════════════════════════════════════════════════════
#  РАССЫЛКА
# ═══════════════════════════════════════════════════════════
async def cmd_broadcast(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> int:
    if not is_admin(update.effective_user.id): return ConversationHandler.END
    kb = [
        [InlineKeyboardButton("✅ Только оплатившим",           callback_data="bc_paid")],
        [InlineKeyboardButton("👥 Всем зарегистрированным",     callback_data="bc_all")],
        [InlineKeyboardButton("❌ Отмена",                      callback_data="bc_cancel")],
    ]
    await update.message.reply_text("📢 *Рассылка*\n\nКому отправить?",
                                    parse_mode=ParseMode.MARKDOWN,
                                    reply_markup=InlineKeyboardMarkup(kb))
    return S_BROADCAST_TARGET

async def cb_bc_target(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> int:
    q = update.callback_query; await q.answer()
    if q.data == "bc_cancel":
        await q.edit_message_text("Рассылка отменена.")
        return ConversationHandler.END
    ctx.user_data["bc_target"] = q.data
    label = "оплатившим" if q.data == "bc_paid" else "всем зарегистрированным"
    await q.edit_message_text(
        f"📝 Напишите текст рассылки *{label}*.\n\n"
        f"Можно прикрепить фото с подписью.\n/cancel — отмена.",
        parse_mode=ParseMode.MARKDOWN
    )
    return S_BROADCAST_TEXT

async def rx_broadcast_text(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> int:
    target = ctx.user_data.get("bc_target", "bc_all")
    ids = db_paid_ids() if target == "bc_paid" else db_all_ids()
    if not ids:
        await update.message.reply_text("⚠️ Нет получателей.")
        return ConversationHandler.END
    text  = update.message.text or update.message.caption or ""
    photo = update.message.photo[-1].file_id if update.message.photo else None
    await update.message.reply_text(f"⏳ Отправляю {len(ids)} получателям...")
    ok = fail = 0
    for uid in ids:
        try:
            if photo:
                await ctx.bot.send_photo(chat_id=uid, photo=photo,
                                         caption=text, parse_mode=ParseMode.MARKDOWN)
            else:
                await ctx.bot.send_message(chat_id=uid, text=text,
                                           parse_mode=ParseMode.MARKDOWN)
            ok += 1
            await asyncio.sleep(0.05)
        except Exception as e:
            logging.warning(f"Broadcast {uid}: {e}"); fail += 1
    await update.message.reply_text(f"✅ Готово!\n📬 Доставлено: {ok}\n❌ Ошибок: {fail}")
    return ConversationHandler.END

async def cmd_cancel(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.reply_text("Отменено. /start — начать заново.",
                                    reply_markup=ReplyKeyboardRemove())
    return ConversationHandler.END

# ═══════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════
def main():
    logging.basicConfig(format="%(asctime)s %(levelname)s %(message)s", level=logging.INFO)
    init_db()

    app = Application.builder().token(BOT_TOKEN).build()

    reg_conv = ConversationHandler(
        entry_points=[CommandHandler("start", cmd_start)],
        states={
            S_CONSENT:    [CallbackQueryHandler(cb_show_docs,    pattern="^show_docs$"),
                           CallbackQueryHandler(cb_read_offer,   pattern="^read_offer$"),
                           CallbackQueryHandler(cb_read_privacy, pattern="^read_privacy$"),
                           CallbackQueryHandler(cb_consent_yes,  pattern="^consent_yes$"),
                           CallbackQueryHandler(cb_consent_no,   pattern="^consent_no$")],
            S_FULLNAME:   [MessageHandler(filters.TEXT & ~filters.COMMAND, rx_fullname)],
            S_PHONE:      [MessageHandler(filters.CONTACT | (filters.TEXT & ~filters.COMMAND), rx_phone)],
            S_EMAIL:      [MessageHandler(filters.TEXT & ~filters.COMMAND, rx_email)],
            S_PAYMENT:    [CallbackQueryHandler(cb_payment_done, pattern="^payment_done$")],
            S_SCREENSHOT: [MessageHandler(filters.PHOTO | filters.Document.IMAGE |
                                          (filters.TEXT & ~filters.COMMAND), rx_screenshot)],
        },
        fallbacks=[CommandHandler("cancel", cmd_cancel)],
        allow_reentry=True,
    )

    support_conv = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex("^💬 Поддержка$"), start_support)],
        states={S_SUPPORT: [MessageHandler(filters.TEXT & ~filters.COMMAND, rx_support),
                             CallbackQueryHandler(cb_support_cancel, pattern="^support_cancel$")]},
        fallbacks=[CommandHandler("cancel", cmd_cancel)],
    )

    broadcast_conv = ConversationHandler(
        entry_points=[CommandHandler("broadcast", cmd_broadcast)],
        states={
            S_BROADCAST_TARGET: [CallbackQueryHandler(cb_bc_target, pattern="^bc_")],
            S_BROADCAST_TEXT:   [MessageHandler(filters.TEXT & ~filters.COMMAND, rx_broadcast_text),
                                  MessageHandler(filters.PHOTO, rx_broadcast_text)],
        },
        fallbacks=[CommandHandler("cancel", cmd_cancel)],
    )

    app.add_handler(reg_conv)
    app.add_handler(support_conv)
    app.add_handler(broadcast_conv)

    app.add_handler(MessageHandler(filters.Regex("^📋 Моя регистрация$"), menu_my_reg))
    app.add_handler(MessageHandler(filters.Regex("^📄 Оферта$"),          menu_offer))
    app.add_handler(MessageHandler(filters.Regex("^🔒 Политика данных$"), menu_privacy))

    app.add_handler(CommandHandler("participants", cmd_participants))
    app.add_handler(CommandHandler("export",       cmd_export))
    app.add_handler(CommandHandler("stats",        cmd_stats))
    app.add_handler(CommandHandler("admin",        cmd_admin_help))
    app.add_handler(CallbackQueryHandler(cb_export_excel, pattern="^export_excel$"))
    app.add_handler(MessageHandler(
        filters.Regex(r"^/confirm_\d+_\d+") & filters.ChatType.PRIVATE, cmd_confirm))
    app.add_handler(MessageHandler(
        filters.Regex(r"^/reply_\d+") & filters.ChatType.PRIVATE, cmd_reply))

    print("🌸 Бот запущен! База данных: bot.db")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
