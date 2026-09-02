import logging
import os
import sqlite3
from datetime import datetime

from telegram import (
    Update,
    ReplyKeyboardRemove,
)
from telegram.ext import (
    Application,
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
    LAVOZIM,
    TAJRIBA,
    REZYUME,
    TASDIQLASH,
) = range(8)


# ----------------------------------------------------------------------------
# BAZA BILAN ISHLASH
# ----------------------------------------------------------------------------
def init_db():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
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
            lavozim TEXT,
            tajriba TEXT,
            rezyume_file_id TEXT,
            created_at TEXT
        )
        """
    )
    conn.commit()
    conn.close()


def save_application(data: dict):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO applications
        (user_id, username, ism, familiya, yosh, telefon, lavozim, tajriba, rezyume_file_id, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            data.get("user_id"),
            data.get("username"),
            data.get("ism"),
            data.get("familiya"),
            data.get("yosh"),
            data.get("telefon"),
            data.get("lavozim"),
            data.get("tajriba"),
            data.get("rezyume_file_id"),
            datetime.now().isoformat(timespec="seconds"),
        ),
    )
    conn.commit()
    conn.close()


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
    await update.message.reply_text(
        "Qaysi lavozimga nomzod bo'lmoqchisiz?",
        reply_markup=ReplyKeyboardRemove(),
    )
    return LAVOZIM


async def get_lavozim(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data["lavozim"] = update.message.text.strip()
    await update.message.reply_text(
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


async def confirm(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    data = context.user_data

    # 1) Bazaga saqlash
    save_application(data)

    # 2) Admin/HR chatiga yuborish
    text = format_application_text(data)
    if ADMIN_CHAT_ID:
        try:
            if data.get("rezyume_file_id"):
                await context.bot.send_document(
                    chat_id=ADMIN_CHAT_ID,
                    document=data["rezyume_file_id"],
                    caption=text,
                    parse_mode="HTML",
                )
            else:
                await context.bot.send_message(
                    chat_id=ADMIN_CHAT_ID, text=text, parse_mode="HTML"
                )
        except Exception as e:
            logger.error("Admin chatga yuborishda xato: %s", e)

    await update.message.reply_text(
        "✅ Arizangiz qabul qilindi! Tez orada siz bilan bog'lanamiz.\n"
        "Rahmat!"
    )
    context.user_data.clear()
    return ConversationHandler.END


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
            LAVOZIM: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_lavozim)],
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
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    application.add_handler(conv_handler)
    application.add_handler(MessageHandler(filters.COMMAND, unknown))

    logger.info("Bot ishga tushdi...")
    application.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
