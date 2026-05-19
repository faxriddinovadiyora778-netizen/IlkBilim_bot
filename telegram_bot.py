import logging
import os
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, InputFile
from telegram.ext import Application, CommandHandler, MessageHandler, CallbackQueryHandler, ContextTypes, filters
from PIL import Image, ImageEnhance, ImageFilter
import io
from docx import Document
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.environ.get("BOT_TOKEN")

REQUIRED_CHANNELS = [
    {"name": "Ilk Bilim", "link": "https://t.me/+aN43lpwgyEU5NThi"},
    {"name": "Pedagoglar", "username": "@pedagoglar1226", "link": "https://t.me/pedagoglar1226"},
]

def main_menu():
    keyboard = [
        [InlineKeyboardButton("📚 Maktab darsliklari", callback_data="darsliklar")],
        [InlineKeyboardButton("📖 Badiiy kitoblar", callback_data="kitoblar")],
        [InlineKeyboardButton("🖼️ Rasm tahrirlash", callback_data="rasm_tahrirl")],
        [InlineKeyboardButton("📄 Rasm → PDF / Word", callback_data="rasm_convert")],
    ]
    return InlineKeyboardMarkup(keyboard)

async def check_sub(user_id, bot):
    not_subbed = []
    try:
        m = await bot.get_chat_member("@pedagoglar1226", user_id)
        if m.status in ["left", "kicked"]:
            not_subbed.append(REQUIRED_CHANNELS[1])
    except:
        not_subbed.append(REQUIRED_CHANNELS[1])
    return not_subbed

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    not_subbed = await check_sub(user.id, context.bot)
    if not_subbed:
        keyboard = [
            [InlineKeyboardButton("➡️ Ilk Bilim kanali", url="https://t.me/+aN43lpwgyEU5NThi")],
            [InlineKeyboardButton("➡️ Pedagoglar kanali", url="https://t.me/pedagoglar1226")],
            [InlineKeyboardButton("✅ Obuna bo'ldim", callback_data="check_sub")],
        ]
        await update.message.reply_text(
            f"👋 Salom {user.first_name}!\n\nBotdan foydalanish uchun kanallarga obuna bo'ling:",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )
    else:
        await update.message.reply_text(f"👋 Salom {user.first_name}! Xush kelibsiz!", reply_markup=main_menu())

async def check_sub_cb(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    not_subbed = await check_sub(query.from_user.id, context.bot)
    if not_subbed:
        keyboard = [
            [InlineKeyboardButton("➡️ Ilk Bilim kanali", url="https://t.me/+aN43lpwgyEU5NThi")],
            [InlineKeyboardButton("➡️ Pedagoglar kanali", url="https://t.me/pedagoglar1226")],
            [InlineKeyboardButton("✅ Obuna bo'ldim", callback_data="check_sub")],
        ]
        await query.edit_message_text("❌ Hali obuna bo'lmadingiz!", reply_markup=InlineKeyboardMarkup(keyboard))
    else:
        await query.edit_message_text("✅ Rahmat! Xush kelibsiz!", reply_markup=main_menu())

DARSLIKLAR = {
    "1-sinf": ["Ona tili", "Matematika", "Atrofimizdagi olam"],
    "2-sinf": ["Ona tili", "Matematika", "Tabiatshunoslik"],
    "3-sinf": ["Ona tili", "Matematika", "Ingliz tili"],
    "4-sinf": ["Ona tili", "Matematika", "Ingliz tili"],
    "5-sinf": ["Matematika", "Ona tili", "Tarix", "Ingliz tili"],
    "6-sinf": ["Matematika", "Ona tili", "Tarix", "Ingliz tili"],
    "7-sinf": ["Algebra", "Fizika", "Kimyo", "Ingliz tili"],
    "8-sinf": ["Algebra", "Fizika", "Kimyo", "Biologiya"],
    "9-sinf": ["Algebra", "Fizika", "Kimyo", "Biologiya"],
    "10-sinf": ["Algebra", "Fizika", "Kimyo", "Biologiya"],
    "11-sinf": ["Algebra", "Fizika", "Kimyo", "Biologiya"],
}

async def cb_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data

    if data == "check_sub":
        await check_sub_cb(update, context)
    elif data == "darsliklar":
        kb = [[InlineKeyboardButton(s, callback_data=f"sinf_{s}")] for s in DARSLIKLAR]
        kb.append([InlineKeyboardButton("🔙 Orqaga", callback_data="back")])
        await query.edit_message_text("📚 Sinfni tanlang:", reply_markup=InlineKeyboardMarkup(kb))
    elif data.startswith("sinf_"):
        sinf = data.replace("sinf_", "")
        kitoblar = DARSLIKLAR.get(sinf, [])
        text = f"📚 {sinf} darsliklari:\n\n" + "\n".join(f"{i+1}. {k}" for i,k in enumerate(kitoblar))
        text += "\n\n📥 Yuklash uchun: @pedagoglar1226"
        await query.edit_message_text(text, reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("🔙 Orqaga", callback_data="darsliklar")]]))
    elif data == "kitoblar":
        turlar = ["Roman", "Hikoya", "She'riyat", "Dunyo adabiyoti", "O'zbek adabiyoti"]
        kb = [[InlineKeyboardButton(t, callback_data=f"kitob_{t}")] for t in turlar]
        kb.append([InlineKeyboardButton("🔙 Orqaga", callback_data="back")])
        await query.edit_message_text("📖 Kitob turini tanlang:", reply_markup=InlineKeyboardMarkup(kb))
    elif data.startswith("kitob_"):
        tur = data.replace("kitob_", "")
        await query.edit_message_text(f"📖 {tur}\n\n📥 Kitob so'rash: @pedagoglar1226", reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("🔙 Orqaga", callback_data="kitoblar")]]))
    elif data == "rasm_tahrirl":
        context.user_data["mode"] = "tahrirl"
        await query.edit_message_text("🖼️ Rasmni yuboring, tahrirlayман!", reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("🔙 Orqaga", callback_data="back")]]))
    elif data == "rasm_convert":
        kb = [
            [InlineKeyboardButton("📄 PDF", callback_data="fmt_pdf")],
            [InlineKeyboardButton("📝 Word", callback_data="fmt_word")],
            [InlineKeyboardButton("🔙 Orqaga", callback_data="back")],
        ]
        await query.edit_message_text("Format tanlang:", reply_markup=InlineKeyboardMarkup(kb))
    elif data.startswith("fmt_"):
        fmt = data.replace("fmt_", "")
        context.user_data["mode"] = "convert"
        context.user_data["fmt"] = fmt
        await query.edit_message_text(f"✅ {fmt.upper()} tanlandi. Rasmni yuboring:", reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("🔙 Orqaga", callback_data="rasm_convert")]]))
    elif data == "back":
        await query.edit_message_text("Asosiy menyu:", reply_markup=main_menu())

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    mode = context.user_data.get("mode", "tahrirl")
    photo = update.message.photo[-1]
    file = await context.bot.get_file(photo.file_id)
    file_bytes = await file.download_as_bytearray()
    img = Image.open(io.BytesIO(file_bytes)).convert("RGB")

    if mode == "convert":
        fmt = context.user_data.get("fmt", "pdf")
        if fmt == "pdf":
            buf = io.BytesIO()
            c = canvas.Canvas(buf, pagesize=A4)
            w, h = A4
            iw, ih = img.size
            ratio = min(w/iw, h/ih) * 0.9
            tmp = io.BytesIO()
            img.save(tmp, format="JPEG")
            tmp.seek(0)
            c.drawImage(ImageReader(tmp), (w-iw*ratio)/2, (h-ih*ratio)/2, iw*ratio, ih*ratio)
            c.save()
            buf.seek(0)
            await update.message.reply_document(InputFile(buf, "rasm.pdf"), caption="✅ PDF tayyor!")
        else:
            doc = Document()
            doc.add_heading("Rasm", 1)
            tmp = io.BytesIO()
            img.save(tmp, format="PNG")
            tmp.seek(0)
            doc.add_picture(tmp)
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            await update.message.reply_document(InputFile(buf, "rasm.docx"), caption="✅ Word tayyor!")
    else:
        img = ImageEnhance.Brightness(img).enhance(1.2)
        img = ImageEnhance.Contrast(img).enhance(1.3)
        img = img.filter(ImageFilter.SHARPEN)
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=95)
        buf.seek(0)
        await update.message.reply_photo(InputFile(buf, "tahrirlangan.jpg"), caption="✅ Rasm tahrirlandi!")

def main():
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CallbackQueryHandler(cb_handler))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    print("Bot ishga tushdi!")
    app.run_polling()

if __name__ == "__main__":
    main()
