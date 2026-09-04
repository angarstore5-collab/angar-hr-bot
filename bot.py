import logging
import os
import sqlite3
from datetime import datetime

from openpyxl import Workbook

from telegram import (
    Update,
    ReplyKeyboardRemove,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
)
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    ConversationHandler,
    MessageHandler,
    filters,
)

# ----------------------------------------------------------------------------
# SOZLAMALAR
# ----------------------------------------------------------------------------
BOT_TOKEN = os.getenv("BOT_TOKEN", "SIZNING_BOT_TOKEN_BU_YERGA")
# HR/Admin guruh yoki shaxsiy chat ID (manfiy son bo'lishi mumkin, agar guruh bo'lsa)
ADMIN_CHAT_ID = int(os.getenv("ADMIN_CHAT_ID", "0"))

DB_PATH = os.path.join(os.path.dirname(__file__), "applications.db")

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

# Suhbat bosqichlari (states)
(
    ISM,
    FAMILIYA,
    YOSH,
    TELEFON,
    FILIAL,
    LAVOZIM,
    LAVOZIM_INFO,
    TAJRIBA,
    REZYUME,
    TASDIQLASH,
) = range(10)

# ----------------------------------------------------------------------------
# FILIALLAR VA LAVOZIMLAR MA'LUMOTLARI
# (Yangi filial yoki lavozim qo'shish uchun shu ro'yxatlarga qator qo'shing)
# ----------------------------------------------------------------------------
FILIALLAR = [
    "Olmazor tumani filiali",
    "Yashnaobod tumani filiali",
    "Qarshi shahar filiali",
]

LAVOZIMLAR = {
    "sotuvchi": {
        "nomi": "Sotuvchi",
        "malumot": (
            "💼 <b>Sotuvchi</b>\n"
            "Ayol-qizlar va yigitlar!\n\n"
            "🕗 <b>Ish grafigi:</b>\n"
            "🔹 1-smena: 08:00 — 18:00\n"
            "🔹 2-smena: 12:00 — 22:00\n\n"
            "✅ <b>Talablar:</b>\n"
            "• 18 yoshdan 25 yoshgacha\n"
            "• Xushmuomala va mijozlar bilan ishlash ko'nikmasiga ega bo'lishi\n"
            "• Mas'uliyatli va halol bo'lishi\n"
            "• Jamoa bilan ishlay olishi\n"
            "• Sotuvchilik sohasida tajribaga ega bo'lishi ma'qul\n\n"
            "💼 <b>Sizdan kutiladigan vazifalar:</b>\n"
            "✔️ Mijozlarni iliq kutib olish va ularning ehtiyojlarini tushunish\n"
            "✔️ Mahsulotlar haqida to'g'ri va aniq ma'lumot berish\n"
            "✔️ Do'kon ichidagi tartib-intizomni saqlash\n"
            "✔️ Rus tili va o'zbek tillarida erkin muloqat qilish\n"
            "✔️ Maishiy texnikalar bo'yicha yetarli bilim va tushunchaga ega bo'lishi lozim\n\n"
            "💰 <b>Oyliq maosh:</b>\n"
            "📌 3 000 000 so'mdan boshlanadi\n"
            "3 oy sinov muddatidan keyin Grade sistemasi bo'yicha oyliklar ko'tariladi + KPI\n\n"
            "🍽 <b>Imtiyozlar:</b>\n"
            "✔️ Ishxona tomonidan tushlik taqdim etiladi\n\n"
            "📌 <b>Eslatma:</b>\n"
            "❗️ Yotoqxona mavjud emas.\n"
            "📍 Do'kon Toshkent shahrida joylashgan."
        ),
    },
    "kassir": {
        "nomi": "Kassir",
        "malumot": (
            "💼 <b>Kassir</b>\n\n"
            "💰 <b>Oylik maosh:</b> 4 000 000 so'mdan boshlanadi\n\n"
            "📌 <b>Talablar:</b>\n"
            "✔️ 18–35 yosh\n"
            "✔️ Kassada ish tajribasi kamida 1 yil\n"
            "✔️ Rus tili va o'zbek tilini bilishi\n"
            "✔️ Xushmuomala va mas'uliyatli bo'lishi kerak"
        ),
    },
    "tozalik": {
        "nomi": "Tozalik xodimi",
        "malumot": (
            "💼 <b>Tozalik xodimi</b>\n\n"
            "🕗 <b>Ish sharoitlari:</b>\n"
            "• Ish vaqti: 08:00 – 14:00\n"
            "• Tushlik ish beruvchi hisobidan\n"
            "• Ish grafigi: 6 kun / 1 dam olish kuni\n"
            "• Mas'uliyatli va mehnatsevar ayollar ishga qabul qilinadi\n"
            "• Yosh: 25 yoshdan 35 yoshgacha\n\n"
            "💰 <b>Oylik maosh:</b> 3 000 000 so'mdan boshlab\n\n"
            "📋 <b>Vazifalari:</b>\n"
            "Do'kon hududida tozalik va tartibni saqlash."
        ),
    },
    "ombor": {
        "nomi": "Ombor ishchisi",
        "malumot": (
            "💼 <b>Ombor ishchisi</b>\n\n"
            "👤 Yosh: 18 yoshdan 25 yoshgacha\n"
            "🕗 Ish vaqti: 09:00 — 18:00\n\n"
            "💰 <b>Maosh:</b> 3 500 000 so'mdan boshlanadi. Ish samarasi va tajribaga qarab oshib boradi.\n\n"
            "📋 <b>Vazifalar:</b>\n"
            "1. Yuklarni tushirish\n"
            "2. Tovarlarni joylashtirish\n"
            "3. Mahsulotlarni joyidan olish va buyurtmalarni yig'ish\n"
            "4. Tovarlarni qadoqlash va chiqarishga tayyorlash\n"
            "5. Hisob yozuvlarini yuritish\n"
            "6. Skladni toza va tartibli saqlash\n"
            "7. Jamoa bilan samarali ishlash\n"
            "8. Xavfsizlik qoidalariga rioya qilish\n\n"
            "✅ <b>Biz taklif qilamiz:</b>\n"
            "✔️ Tushlik ishxona hisobidan\n"
            "✔️ Do'stona va ahil jamoa\n"
            "❗️ Yotoqxona mavjud emas"
        ),
    },
}


# ----------------------------------------------------------------------------
# BAZA BILAN ISHLASH
# ----------------------------------------------------------------------------
def init_db():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    # Asosiy jadval — barcha kelgan anketalar (holatidan qat'i nazar)
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS applications (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            username TEXT,
            ism TEXT,
            familiya TEXT,
            yosh TEXT,
            telefon TEXT,
            filial TEXT,
            lavozim TEXT,
            tajriba TEXT,
            rezyume_file_id TEXT,
            status TEXT DEFAULT 'kutilmoqda',
            created_at TEXT
        )
        """
    )
    # Qabul qilinganlar uchun alohida jadval
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS qabul_qilinganlar (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            original_id INTEGER,
            user_id INTEGER,
            username TEXT,
            ism TEXT,
            familiya TEXT,
            yosh TEXT,
            telefon TEXT,
            filial TEXT,
            lavozim TEXT,
            tajriba TEXT,
            rezyume_file_id TEXT,
            created_at TEXT,
            qabul_vaqti TEXT
        )
        """
    )
    # Arxivlanganlar uchun alohida jadval
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS arxivlanganlar (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            original_id INTEGER,
            user_id INTEGER,
            username TEXT,
            ism TEXT,
            familiya TEXT,
            yosh TEXT,
            telefon TEXT,
            filial TEXT,
            lavozim TEXT,
            tajriba TEXT,
            rezyume_file_id TEXT,
            created_at TEXT,
            arxiv_vaqti TEXT
        )
        """
    )
    conn.commit()
    conn.close()


def save_application(data: dict) -> int:
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO applications
        (user_id, username, ism, familiya, yosh, telefon, filial, lavozim, tajriba, rezyume_file_id, status, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'kutilmoqda', ?)
        """,
        (
            data.get("user_id"),
            data.get("username"),
            data.get("ism"),
            data.get("familiya"),
            data.get("yosh"),
            data.get("telefon"),
            data.get("filial"),
            data.get("lavozim"),
            data.get("tajriba"),
            data.get("rezyume_file_id"),
            datetime.now().isoformat(timespec="seconds"),
        ),
    )
    new_id = cur.lastrowid
    conn.commit()
    conn.close()
    return new_id


def update_status(app_id: int, status: str):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("UPDATE applications SET status = ? WHERE id = ?", (status, app_id))
    conn.commit()
    conn.close()


def move_to_table(app_data: dict, target_table: str, time_column: str):
    """Anketani 'qabul_qilinganlar' yoki 'arxivlanganlar' jadvaliga nusxalaydi."""
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute(
        f"""
        INSERT INTO {target_table}
        (original_id, user_id, username, ism, familiya, yosh, telefon, filial, lavozim, tajriba, rezyume_file_id, created_at, {time_column})
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            app_data.get("id"),
            app_data.get("user_id"),
            app_data.get("username"),
            app_data.get("ism"),
            app_data.get("familiya"),
            app_data.get("yosh"),
            app_data.get("telefon"),
            app_data.get("filial"),
            app_data.get("lavozim"),
            app_data.get("tajriba"),
            app_data.get("rezyume_file_id"),
            app_data.get("created_at"),
            datetime.now().isoformat(timespec="seconds"),
        ),
    )
    conn.commit()
    conn.close()


def get_application(app_id: int):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute("SELECT * FROM applications WHERE id = ?", (app_id,))
    row = cur.fetchone()
    conn.close()
    return dict(row) if row else None


def get_all_rows(table_name: str):
    """Berilgan jadvaldan barcha qatorlarni ro'yxat (dict) shaklida qaytaradi."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute(f"SELECT * FROM {table_name} ORDER BY id")
    rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows]


def build_excel_file(rows: list, time_column: str, file_path: str):
    """Berilgan qatorlar asosida Excel (.xlsx) fayl yaratadi."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Anketalar"

    headers = [
        "ID", "Ism", "Familiya", "Yosh", "Telefon", "Filial",
        "Lavozim", "Tajriba", "Ariza sanasi", "Holat sanasi", "Telegram username",
    ]
    ws.append(headers)

    for row in rows:
        ws.append([
            row.get("original_id") or row.get("id"),
            row.get("ism"),
            row.get("familiya"),
            row.get("yosh"),
            row.get("telefon"),
            row.get("filial"),
            row.get("lavozim"),
            row.get("tajriba"),
            row.get("created_at"),
            row.get(time_column),
            f"@{row.get('username')}" if row.get("username") else "—",
        ])

    # Ustunlar kengligini avtomatik moslashtirish
    for col_cells in ws.columns:
        max_len = max((len(str(c.value)) for c in col_cells if c.value is not None), default=10)
        ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 40)

    wb.save(file_path)


async def admin_only_guard(update: Update) -> bool:
    """Buyruq faqat ADMIN_CHAT_ID'dan kelayotganini tekshiradi."""
    if update.effective_chat.id != ADMIN_CHAT_ID:
        await update.message.reply_text("Bu buyruq faqat admin uchun mavjud.")
        return False
    return True


async def qabullar_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await admin_only_guard(update):
        return
    try:
        rows = get_all_rows("qabul_qilinganlar")
        if not rows:
            await update.message.reply_text("Hozircha qabul qilingan nomzodlar yo'q.")
            return
        file_path = os.path.join(os.path.dirname(__file__), "qabul_qilinganlar.xlsx")
        build_excel_file(rows, "qabul_vaqti", file_path)
        await update.message.reply_document(
            document=open(file_path, "rb"),
            filename="Qabul_qilinganlar.xlsx",
            caption=f"✅ Qabul qilinganlar ro'yxati ({len(rows)} ta nomzod)",
        )
    except Exception as e:
        logger.error("qabullar_command xatosi: %s", e)
        await update.message.reply_text(f"⚠️ Xatolik yuz berdi: {e}")


async def arxiv_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await admin_only_guard(update):
        return
    try:
        rows = get_all_rows("arxivlanganlar")
        if not rows:
            await update.message.reply_text("Hozircha arxivlangan nomzodlar yo'q.")
            return
        file_path = os.path.join(os.path.dirname(__file__), "arxivlanganlar.xlsx")
        build_excel_file(rows, "arxiv_vaqti", file_path)
        await update.message.reply_document(
            document=open(file_path, "rb"),
            filename="Arxivlanganlar.xlsx",
            caption=f"🗄 Arxivlangan nomzodlar ro'yxati ({len(rows)} ta nomzod)",
        )
    except Exception as e:
        logger.error("arxiv_command xatosi: %s", e)
        await update.message.reply_text(f"⚠️ Xatolik yuz berdi: {e}")


# ----------------------------------------------------------------------------
# YORDAMCHI FUNKSIYA: Anketa matnini yig'ish
# ----------------------------------------------------------------------------
def format_application_text(data: dict) -> str:
    return (
        "🆕 <b>Yangi anketa</b>\n\n"
        f"👤 <b>Ism:</b> {data.get('ism')}\n"
        f"👤 <b>Familiya:</b> {data.get('familiya')}\n"
        f"🎂 <b>Yosh:</b> {data.get('yosh')}\n"
        f"📞 <b>Telefon:</b> {data.get('telefon')}\n"
        f"🏢 <b>Filial:</b> {data.get('filial')}\n"
        f"💼 <b>Lavozim:</b> {data.get('lavozim')}\n"
        f"📈 <b>Tajriba:</b> {data.get('tajriba')}\n"
        f"🔗 <b>Username:</b> @{data.get('username') or '—'}\n"
        f"🆔 <b>User ID:</b> {data.get('user_id')}"
    )


# ----------------------------------------------------------------------------
# SUHBAT QADAMLARI
# ----------------------------------------------------------------------------
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.clear()
    await update.message.reply_text(
        "Assalomu alaykum! 👋\n"
        "Bo'sh ish o'rniga ariza topshirish uchun anketani to'ldiramiz.\n\n"
        "Istalgan vaqtda /cancel buyrug'i bilan bekor qilishingiz mumkin.\n\n"
        "Ismingizni kiriting:",
        reply_markup=ReplyKeyboardRemove(),
    )
    return ISM


async def get_ism(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data["ism"] = update.message.text.strip()
    await update.message.reply_text("Familiyangizni kiriting:")
    return FAMILIYA


async def get_familiya(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data["familiya"] = update.message.text.strip()
    await update.message.reply_text("Yoshingizni kiriting (masalan: 25):")
    return YOSH


async def get_yosh(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    text = update.message.text.strip()
    if not text.isdigit() or not (14 <= int(text) <= 80):
        await update.message.reply_text(
            "Iltimos, yoshingizni raqamda va to'g'ri kiriting (masalan: 25):"
        )
        return YOSH
    context.user_data["yosh"] = text
    await update.message.reply_text(
        "Telefon raqamingizni kiriting (masalan: +998901234567)\n"
        "yoki pastdagi tugma orqali yuboring:",
    )
    return TELEFON


async def get_telefon(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    if update.message.contact:
        telefon = update.message.contact.phone_number
    else:
        telefon = update.message.text.strip()
    context.user_data["telefon"] = telefon

    keyboard = [
        [InlineKeyboardButton(filial, callback_data=f"filial:{i}")]
        for i, filial in enumerate(FILIALLAR)
    ]
    await update.message.reply_text(
        "Sizga yaqin filialni tanlang:",
        reply_markup=ReplyKeyboardRemove(),
    )
    await update.message.reply_text(
        "👇",
        reply_markup=InlineKeyboardMarkup(keyboard),
    )
    return FILIAL


async def get_filial(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()

    index = int(query.data.split(":")[1])
    context.user_data["filial"] = FILIALLAR[index]

    keyboard = [
        [InlineKeyboardButton(info["nomi"], callback_data=f"lavozim:{key}")]
        for key, info in LAVOZIMLAR.items()
    ]
    await query.edit_message_text(
        f"✅ Filial: {FILIALLAR[index]}\n\nEndi lavozimni tanlang:",
        reply_markup=InlineKeyboardMarkup(keyboard),
    )
    return LAVOZIM


async def get_lavozim(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()

    key = query.data.split(":")[1]
    lavozim = LAVOZIMLAR[key]
    context.user_data["_pending_lavozim_key"] = key

    keyboard = [
        [
            InlineKeyboardButton("✅ To'g'ri keladi, tasdiqlayman", callback_data="lavozim_confirm"),
            InlineKeyboardButton("⬅️ Orqaga", callback_data="lavozim_back"),
        ]
    ]
    await query.edit_message_text(
        lavozim["malumot"] + "\n\nMazkur lavozim sizga mos keladimi?",
        parse_mode="HTML",
        reply_markup=InlineKeyboardMarkup(keyboard),
    )
    return LAVOZIM_INFO


async def handle_lavozim_decision(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()

    if query.data == "lavozim_back":
        keyboard = [
            [InlineKeyboardButton(info["nomi"], callback_data=f"lavozim:{key}")]
            for key, info in LAVOZIMLAR.items()
        ]
        await query.edit_message_text(
            f"✅ Filial: {context.user_data.get('filial')}\n\nLavozimni tanlang:",
            reply_markup=InlineKeyboardMarkup(keyboard),
        )
        return LAVOZIM

    # lavozim_confirm
    key = context.user_data.pop("_pending_lavozim_key", None)
    context.user_data["lavozim"] = LAVOZIMLAR[key]["nomi"] if key else "Noma'lum"

    await query.edit_message_text(
        f"✅ Lavozim tanlandi: {context.user_data['lavozim']}"
    )
    await query.message.reply_text(
        "Ish tajribangiz haqida qisqacha yozing (yo'q bo'lsa \"yo'q\" deb yozing):"
    )
    return TAJRIBA


async def get_tajriba(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data["tajriba"] = update.message.text.strip()
    await update.message.reply_text(
        "Endi rezyumeingizni (CV) PDF yoki Word fayl sifatida yuboring.\n"
        "Agar rezyumeingiz bo'lmasa, /skip buyrug'ini yuboring."
    )
    return REZYUME


async def get_rezyume(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    document = update.message.document
    if document is None:
        await update.message.reply_text(
            "Iltimos, fayl (hujjat) sifatida yuboring, yoki /skip bosing."
        )
        return REZYUME
    context.user_data["rezyume_file_id"] = document.file_id
    context.user_data["rezyume_file_name"] = document.file_name
    return await show_summary(update, context)


async def skip_rezyume(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data["rezyume_file_id"] = None
    context.user_data["rezyume_file_name"] = None
    return await show_summary(update, context)


async def show_summary(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    data = context.user_data
    user = update.effective_user
    data["user_id"] = user.id
    data["username"] = user.username

    text = format_application_text(data)
    await update.message.reply_text(
        "Ma'lumotlaringizni tekshiring:\n\n" + text.replace("<b>", "").replace("</b>", ""),
        parse_mode=None,
    )
    await update.message.reply_text(
        "Hammasi to'g'rimi? Tasdiqlash uchun /confirm, qaytadan boshlash uchun /restart yuboring."
    )
    return TASDIQLASH


def build_admin_keyboard(app_id: int) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        [
            [
                InlineKeyboardButton("✅ Qabul qilish", callback_data=f"accept:{app_id}"),
                InlineKeyboardButton("🗄 Arxivga yuborish", callback_data=f"archive:{app_id}"),
            ]
        ]
    )


async def confirm(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    data = context.user_data

    # 1) Bazaga saqlash (yangi anketa ID'sini olamiz)
    app_id = save_application(data)

    # 2) Admin/HR chatiga tugmalar bilan yuborish
    text = format_application_text(data)
    keyboard = build_admin_keyboard(app_id)
    if ADMIN_CHAT_ID:
        try:
            if data.get("rezyume_file_id"):
                await context.bot.send_document(
                    chat_id=ADMIN_CHAT_ID,
                    document=data["rezyume_file_id"],
                    caption=text,
                    parse_mode="HTML",
                    reply_markup=keyboard,
                )
            else:
                await context.bot.send_message(
                    chat_id=ADMIN_CHAT_ID,
                    text=text,
                    parse_mode="HTML",
                    reply_markup=keyboard,
                )
        except Exception as e:
            logger.error("Admin chatga yuborishda xato: %s", e)

    await update.message.reply_text(
        "✅ Arizangiz qabul qilindi! Tez orada siz bilan bog'lanamiz.\n"
        "Rahmat!"
    )
    context.user_data.clear()
    return ConversationHandler.END


# ----------------------------------------------------------------------------
# ADMIN TUGMALARI: Qabul qilish / Arxivga yuborish
# ----------------------------------------------------------------------------
async def handle_admin_decision(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query

    action, app_id_str = query.data.split(":")
    app_id = int(app_id_str)

    app_data = get_application(app_id)
    if not app_data:
        await query.answer("Anketa topilmadi.", show_alert=True)
        return

    if app_data.get("status") != "kutilmoqda":
        await query.answer("Bu anketa bo'yicha qaror allaqachon qabul qilingan.", show_alert=True)
        return

    await query.answer()  # oddiy "bosildi" signali

    if action == "accept":
        new_status = "qabul qilindi ✅"
        update_status(app_id, "qabul_qilindi")
        move_to_table(app_data, "qabul_qilinganlar", "qabul_vaqti")
        nomzod_xabari = (
            "🎉 Tabriklaymiz! Sizning arizangiz ko'rib chiqildi va siz keyingi bosqichga "
            "taklif qilindingiz. Tez orada siz bilan bog'lanamiz."
        )
    elif action == "archive":
        new_status = "arxivlandi 🗄"
        update_status(app_id, "arxivlandi")
        move_to_table(app_data, "arxivlanganlar", "arxiv_vaqti")
        nomzod_xabari = (
            "Xabaringiz uchun rahmat. Hozircha sizning nomzodingiz boshqa vakansiyalar "
            "uchun arxivda saqlanadi. Mos lavozim ochilganda albatta bog'lanamiz."
        )
    else:
        return

    # Nomzodga xabar yuborishga urinib ko'ramiz (agar u botni bloklamagan bo'lsa)
    try:
        await context.bot.send_message(chat_id=app_data["user_id"], text=nomzod_xabari)
    except Exception as e:
        logger.warning("Nomzodga xabar yuborib bo'lmadi (user_id=%s): %s", app_data["user_id"], e)

    # Admin chatidagi xabarni yangilaymiz: status ko'rsatamiz, tugmalarni olib tashlaymiz
    status_text = f"\n\n📌 <b>Holat:</b> {new_status}"
    try:
        if query.message.caption:
            await query.edit_message_caption(
                caption=query.message.caption_html + status_text,
                parse_mode="HTML",
            )
        else:
            await query.edit_message_text(
                text=query.message.text_html + status_text,
                parse_mode="HTML",
            )
    except Exception as e:
        logger.error("Admin xabarini yangilashda xato: %s", e)


async def restart(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    return await start(update, context)


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.clear()
    await update.message.reply_text(
        "Anketa bekor qilindi. Qaytadan boshlash uchun /start yuboring.",
        reply_markup=ReplyKeyboardRemove(),
    )
    return ConversationHandler.END


async def unknown(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        "Buyruqni tushunmadim. Boshlash uchun /start yuboring."
    )


# ----------------------------------------------------------------------------
# ASOSIY FUNKSIYA
# ----------------------------------------------------------------------------
def main():
    if BOT_TOKEN == "SIZNING_BOT_TOKEN_BU_YERGA":
        raise RuntimeError(
            "BOT_TOKEN o'rnatilmagan! Environment variable BOT_TOKEN ni sozlang."
        )

    init_db()

    application = Application.builder().token(BOT_TOKEN).build()

    conv_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            ISM: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_ism)],
            FAMILIYA: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_familiya)],
            YOSH: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_yosh)],
            TELEFON: [
                MessageHandler(filters.CONTACT, get_telefon),
                MessageHandler(filters.TEXT & ~filters.COMMAND, get_telefon),
            ],
            FILIAL: [CallbackQueryHandler(get_filial, pattern=r"^filial:\d+$")],
            LAVOZIM: [CallbackQueryHandler(get_lavozim, pattern=r"^lavozim:")],
            LAVOZIM_INFO: [
                CallbackQueryHandler(handle_lavozim_decision, pattern=r"^lavozim_confirm$|^lavozim_back$")
            ],
            TAJRIBA: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_tajriba)],
            REZYUME: [
                MessageHandler(filters.Document.ALL, get_rezyume),
                CommandHandler("skip", skip_rezyume),
            ],
            TASDIQLASH: [
                CommandHandler("confirm", confirm),
                CommandHandler("restart", restart),
            ],
        },
        fallbacks=[
            CommandHandler("cancel", cancel),
            CommandHandler("qabullar", qabullar_command),
            CommandHandler("arxiv", arxiv_command),
        ],
    )

    application.add_handler(conv_handler)
    application.add_handler(CallbackQueryHandler(handle_admin_decision))
    application.add_handler(CommandHandler("qabullar", qabullar_command))
    application.add_handler(CommandHandler("arxiv", arxiv_command))
    application.add_handler(MessageHandler(filters.COMMAND, unknown))

    logger.info("Bot ishga tushdi...")
    application.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()

