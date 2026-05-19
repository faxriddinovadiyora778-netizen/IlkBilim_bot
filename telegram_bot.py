import logging
import os
from telegram import (
    Update, InlineKeyboardButton, InlineKeyboardMarkup, InputFile
)
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, ContextTypes, filters
)
from PIL import Image, ImageEnhance, ImageFilter
import io
from docx import Document
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

# ===================== SOZLAMALAR =====================
BOT_TOKEN = "BU_YERGA_TOKEN_KIRITING"

# Obuna bo'lishi shart kanallar
REQUIRED_CHANNELS = [
    {"name": "Asosiy kanal", "username": "@sizning_kanal1", "link": "https://t.me/sizning_kanal1"},
    {"name": "Kitoblar kanali", "username": "@sizning_kanal2", "link": "https://t.me/sizning_kanal2"},
]

# ===================== LOGGING =====================
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ===================== KANALGA OBUNA TEKSHIRISH =====================
async def check_subscriptions(user_id: int, bot) -> list:
    """Foydalanuvchi obuna bo'lmagan kanallarni qaytaradi"""
    not_subscribed = []
    for channel in REQUIRED_CHANNELS:
        try:
            member = await bot.get_chat_member(channel["username"], user_id)
            if member.status in ["left", "kicked", "banned"]:
                not_subscribed.append(channel)
        except Exception:
            not_subscribed.append(channel)
    return not_subscribed

# ===================== ASOSIY MENYU =====================
def main_menu_keyboard():
    keyboard = [
        [InlineKeyboardButton("📚 Maktab darsliklari", callback_data="darsliklar")],
        [InlineKeyboardButton("📖 Badiiy kitoblar", callback_data="kitoblar")],
        [InlineKeyboardButton("🖼️ Rasm tahrirlash", callback_data="rasm_tahrirl")],
        [InlineKeyboardButton("📄 Rasm → PDF / Word", callback_data="rasm_convert")],
    ]
    return InlineKeyboardMarkup(keyboard)

# ===================== /start =====================
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    not_subscribed = await check_subscriptions(user.id, context.bot)

    if not_subscribed:
        # Obuna bo'lmagan kanallar bor
        keyboard = []
        for ch in not_subscribed:
            keyboard.append([InlineKeyboardButton(f"➡️ {ch['name']}", url=ch["link"])])
        keyboard.append([InlineKeyboardButton("✅ Obuna bo'ldim, tekshir", callback_data="check_sub")])

        await update.message.reply_text(
            f"👋 Salom, {user.first_name}!\n\n"
            "Botdan foydalanish uchun quyidagi kanallarga obuna bo'ling:\n",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )
    else:
        await update.message.reply_text(
            f"👋 Salom, {user.first_name}! Xush kelibsiz!\n\n"
            "Quyidagi bo'limlardan birini tanlang:",
            reply_markup=main_menu_keyboard()
        )

# ===================== OBUNA TEKSHIRISH TUGMASI =====================
async def check_sub_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user = query.from_user
    not_subscribed = await check_subscriptions(user.id, context.bot)

    if not_subscribed:
        keyboard = []
        for ch in not_subscribed:
            keyboard.append([InlineKeyboardButton(f"➡️ {ch['name']}", url=ch["link"])])
        keyboard.append([InlineKeyboardButton("✅ Obuna bo'ldim, tekshir", callback_data="check_sub")])
        await query.edit_message_text(
            "❌ Siz hali barcha kanallarga obuna bo'lmadingiz!\n\n"
            "Iltimos, quyidagi kanallarga obuna bo'ling:",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )
    else:
        await query.edit_message_text(
            "✅ Rahmat! Obuna tasdiqlandi!\n\nQuyidagi bo'limlardan birini tanlang:",
            reply_markup=main_menu_keyboard()
        )

# ===================== DARSLIKLAR =====================
DARSLIKLAR = {
    "1-sinf": ["Ona tili", "Matematika", "Atrofimizdagi olam"],
    "2-sinf": ["Ona tili", "Matematika", "Tabiatshunoslik"],
    "3-sinf": ["Ona tili", "Matematika", "Ingliz tili"],
    "5-sinf": ["Matematika", "Ona tili va adabiyot", "Tarix", "Ingliz tili"],
    "10-sinf": ["Algebra", "Fizika", "Kimyo", "Biologiya"],
    "11-sinf": ["Algebra", "Fizika", "Kimyo", "Biologiya"],
}

async def darsliklar_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    keyboard = [[InlineKeyboardButton(sinf, callback_data=f"sinf_{sinf}")] for sinf in DARSLIKLAR.keys()]
    keyboard.append([InlineKeyboardButton("🔙 Orqaga", callback_data="back_main")])
    await query.edit_message_text("📚 Sinfni tanlang:", reply_markup=InlineKeyboardMarkup(keyboard))

async def sinf_darsliklar(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    sinf = query.data.replace("sinf_", "")
    kitoblar = DARSLIKLAR.get(sinf, [])
    text = f"📚 {sinf} darsliklari:\n\n"
    for i, kitob in enumerate(kitoblar, 1):
        text += f"{i}. {kitob}\n"
    text += "\n📥 Kerakli darslikni yuklash uchun admin bilan bog'laning: @admin_username"
    keyboard = [[InlineKeyboardButton("🔙 Orqaga", callback_data="darsliklar")]]
    await query.edit_message_text(text, reply_markup=InlineKeyboardMarkup(keyboard))

# ===================== BADIIY KITOBLAR =====================
KITOB_TURLARI = ["Roman", "Hikoya", "She'riyat", "Dunyo adabiyoti", "O'zbek adabiyoti"]

async def kitoblar_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    keyboard = [[InlineKeyboardButton(tur, callback_data=f"kitob_{tur}")] for tur in KITOB_TURLARI]
    keyboard.append([InlineKeyboardButton("🔙 Orqaga", callback_data="back_main")])
    await query.edit_message_text("📖 Kitob turini tanlang:", reply_markup=InlineKeyboardMarkup(keyboard))

async def kitob_turi(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    tur = query.data.replace("kitob_", "")
    text = (
        f"📖 {tur} bo'limi\n\n"
        "Bu bo'limda eng mashhur kitoblarni topasiz.\n"
        "📥 Kitob so'rash yoki yuklash uchun: @admin_username\n\n"
        "🔍 Yoki kitob nomini yozing, qidirib beramiz!"
    )
    keyboard = [[InlineKeyboardButton("🔙 Orqaga", callback_data="kitoblar")]]
    await query.edit_message_text(text, reply_markup=InlineKeyboardMarkup(keyboard))

# ===================== RASM TAHRIRLASH =====================
async def rasm_tahrirl_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    context.user_data["mode"] = "tahrirl"
    await query.edit_message_text(
        "🖼️ Rasmni yuboring, uni tahrirlayман!\n\n"
        "Tahrirlash imkoniyatlari:\n"
        "• Yorqinlikni oshirish\n"
        "• Kontrastni yaxshilash\n"
        "• Sharp (aniqroq) qilish\n"
        "• Grayscale (qora-oq)\n\n"
        "📤 Rasmni hozir yuboring:",
        reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("🔙 Orqaga", callback_data="back_main")]])
    )

async def rasm_convert_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    context.user_data["mode"] = "convert"
    keyboard = [
        [InlineKeyboardButton("📄 PDF ga o'tkazish", callback_data="convert_pdf")],
        [InlineKeyboardButton("📝 Word ga o'tkazish", callback_data="convert_word")],
        [InlineKeyboardButton("🔙 Orqaga", callback_data="back_main")],
    ]
    await query.edit_message_text(
        "📄 Rasm → PDF yoki Word\n\nAvval formatni tanlang:",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )

async def set_convert_format(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    fmt = query.data.replace("convert_", "")
    context.user_data["convert_format"] = fmt
    context.user_data["mode"] = "convert"
    fmt_name = "PDF" if fmt == "pdf" else "Word"
    await query.edit_message_text(
        f"✅ Format tanlandi: {fmt_name}\n\n📤 Endi rasmni yuboring:",
        reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("🔙 Orqaga", callback_data="rasm_convert")]])
    )

# ===================== RASM QABUL QILISH =====================
async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    not_subscribed = await check_subscriptions(user.id, context.bot)
    if not_subscribed:
        await update.message.reply_text("❌ Botdan foydalanish uchun avval kanallarga obuna bo'ling! /start")
        return

    mode = context.user_data.get("mode", "tahrirl")
    photo = update.message.photo[-1]
    file = await context.bot.get_file(photo.file_id)
    file_bytes = await file.download_as_bytearray()
    img = Image.open(io.BytesIO(file_bytes))

    if mode == "convert":
        fmt = context.user_data.get("convert_format", "pdf")
        if fmt == "pdf":
            await convert_to_pdf(update, img)
        else:
            await convert_to_word(update, img)
    else:
        # Rasm tahrirlash
        await edit_image(update, img)

async def edit_image(update: Update, img: Image.Image):
    """Rasmni tahrirlash: yorqinlik, kontrast, sharpness"""
    img = img.convert("RGB")
    enhancer = ImageEnhance.Brightness(img)
    img = enhancer.enhance(1.2)
    enhancer = ImageEnhance.Contrast(img)
    img = enhancer.enhance(1.3)
    img = img.filter(ImageFilter.SHARPEN)

    output = io.BytesIO()
    img.save(output, format="JPEG", quality=95)
    output.seek(0)

    await update.message.reply_photo(
        photo=InputFile(output, filename="tahrirlangan.jpg"),
        caption="✅ Rasm tahrirlandi!\n🌟 Yorqinlik +20%\n🎨 Kontrast +30%\n🔍 Sharpness yaxshilandi"
    )

async def convert_to_pdf(update: Update, img: Image.Image):
    """Rasmni PDF ga o'tkazish"""
    img = img.convert("RGB")
    pdf_buffer = io.BytesIO()
    page_width, page_height = A4
    c = canvas.Canvas(pdf_buffer, pagesize=A4)

    img_w, img_h = img.size
    ratio = min(page_width / img_w, page_height / img_h) * 0.9
    new_w = img_w * ratio
    new_h = img_h * ratio
    x = (page_width - new_w) / 2
    y = (page_height - new_h) / 2

    img_temp = io.BytesIO()
    img.save(img_temp, format="JPEG")
    img_temp.seek(0)

    from reportlab.lib.utils import ImageReader
    c.drawImage(ImageReader(img_temp), x, y, width=new_w, height=new_h)
    c.save()
    pdf_buffer.seek(0)

    await update.message.reply_document(
        document=InputFile(pdf_buffer, filename="rasm.pdf"),
        caption="✅ Rasm PDF ga o'tkazildi!"
    )

async def convert_to_word(update: Update, img: Image.Image):
    """Rasmni Word ga o'tkazish"""
    img = img.convert("RGB")
    img_buffer = io.BytesIO()
    img.save(img_buffer, format="PNG")
    img_buffer.seek(0)

    doc = Document()
    doc.add_heading("Rasm", level=1)
    doc.add_picture(img_buffer, width=doc.sections[0].page_width * 0.8)
    doc.add_paragraph("Bot tomonidan yaratildi.")

    word_buffer = io.BytesIO()
    doc.save(word_buffer)
    word_buffer.seek(0)

    await update.message.reply_document(
        document=InputFile(word_buffer, filename="rasm.docx"),
        caption="✅ Rasm Word (.docx) ga o'tkazildi!"
    )

# ===================== ORQAGA =====================
async def back_main(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    context.user_data.pop("mode", None)
    await query.edit_message_text("Asosiy menyu:", reply_markup=main_menu_keyboard())

# ===================== CALLBACK HANDLER =====================
async def callback_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    data = query.data

    if data == "check_sub":
        await check_sub_callback(update, context)
    elif data == "darsliklar":
        await darsliklar_menu(update, context)
    elif data.startswith("sinf_"):
        await sinf_darsliklar(update, context)
    elif data == "kitoblar":
        await kitoblar_menu(update, context)
    elif data.startswith("kitob_"):
        await kitob_turi(update, context)
    elif data == "rasm_tahrirl":
        await rasm_tahrirl_menu(update, context)
    elif data == "rasm_convert":
        await rasm_convert_menu(update, context)
    elif data.startswith("convert_"):
        await set_convert_format(update, context)
    elif data == "back_main":
        await back_main(update, context)

# ===================== MAIN =====================
def main():
    app = Application.builder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CallbackQueryHandler(callback_handler))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))

    print("✅ Bot ishga tushdi!")
    app.run_polling()

if __name__ == "__main__":
    main()
