import logging
import os
import io
from telegram import Update, ReplyKeyboardMarkup, ReplyKeyboardRemove, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, CallbackQueryHandler, ContextTypes, filters
from PIL import Image, ImageEnhance, ImageFilter
from docx import Document
from docx.shared import Inches, Pt
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.environ.get("BOT_TOKEN")
ADMIN_ID = 8016799878

file_storage = {}

DARSLIKLAR = {
    "1-sinf": [
        "Alifbe (2025)", "Ona tili (1-qism)", "Ona tili (2-qism)",
        "O'qish savodxonligi (1-qism)", "O'qish savodxonligi (2-qism)",
        "Matematika (1-qism)", "Matematika (2-qism)", "Matematika (3-qism)",
        "Matematika (4-qism)", "Matematika (2-mashq daftari)",
        "Tabiiy fanlar (1-qism)", "Tabiiy fanlar (2-qism)",
        "Informatika va AT", "Musiqiy savodxonlik", "Tasviriy san'at",
        "Texnologiya", "Texnologiya (mashq daftari)",
        "Tarbiya", "Yozuv daftari", "Jismoniy tarbiya (ish daftari)", "Ingliz tili",
    ],
    "2-sinf": [
        "Ona tili (1-qism)", "Ona tili (2-qism)", "Ona tili (3-qism)", "Ona tili (4-qism)",
        "O'qish savodxonligi (1-qism)", "O'qish savodxonligi (2-qism)",
        "O'qish savodxonligi (3-qism)", "O'qish savodxonligi (4-qism)",
        "O'qish savodxonligi (mashq daftari)",
        "Matematika (1-qism)", "Matematika (2-qism)", "Matematika (3-qism)", "Matematika (4-qism)",
        "Tabiiy fanlar (1-qism)", "Tabiiy fanlar (2-qism)",
        "Informatika va AT", "Musiqiy savodxonlik", "Tasviriy san'at",
        "Texnologiya", "Tarbiya", "Jismoniy tarbiya (ish daftari)",
        "Ingliz tili", "Rus tili",
    ],
    "3-sinf": [
        "Ona tili (1-qism)", "Ona tili (2-qism)", "Ona tili (3-qism)", "Ona tili (4-qism)",
        "O'qish savodxonligi (1-qism)", "O'qish savodxonligi (2-qism)",
        "O'qish savodxonligi (3-qism)", "O'qish savodxonligi (4-qism)",
        "Matematika (1-qism)", "Matematika (2-qism)", "Matematika (3-qism)", "Matematika (4-qism)",
        "Tabiiy fanlar (1-qism)", "Tabiiy fanlar (2-qism)",
        "Musiqiy savodxonlik", "Tasviriy san'at", "Texnologiya", "Tarbiya", "Ingliz tili",
    ],
    "4-sinf": [
        "Ona tili (1-qism)", "Ona tili (2-qism)", "Ona tili (3-qism)", "Ona tili (4-qism)",
        "O'qish savodxonligi (1-qism)", "O'qish savodxonligi (2-qism)",
        "O'qish savodxonligi (3-qism)", "O'qish savodxonligi (4-qism)",
        "Matematika (1-qism)", "Matematika (2-qism)", "Matematika (3-qism)", "Matematika (4-qism)",
        "Tabiiy fanlar (1-qism)", "Tabiiy fanlar (2-qism)", "Tabiiy fanlar (mashq daftari)",
        "Ingliz tili", "Rus tili", "Musiqiy savodxonlik", "Tasviriy san'at",
        "Texnologiya", "Tarbiya", "Jismoniy tarbiya",
    ],
    "5-sinf": [
        "Ona tili", "Adabiyot (1-qism)", "Adabiyot (2-qism)",
        "Matematika (1-qism)", "Matematika (2-qism)",
        "Tarix (Tarixdan hikoyalar)",
        "Tabiiy fanlar (1-qism)", "Tabiiy fanlar (2-qism)",
        "Informatika va AT", "Texnologiya", "Tasviriy san'at",
        "Musiqa madaniyati", "Tarbiya", "Jismoniy tarbiya",
        "Ingliz tili", "Rus tili",
    ],
    "6-sinf": [
        "Ona tili", "Adabiyot (1-qism)", "Adabiyot (2-qism)",
        "Matematika (1-qism)", "Matematika (2-qism)",
        "Tarix (Qadimgi dunyo tarixi)", "Biologiya (Botanika)",
        "Geografiya", "Informatika va AT", "Texnologiya",
        "Tasviriy san'at", "Musiqa madaniyati", "Tarbiya",
        "Jismoniy tarbiya", "Ingliz tili", "Rus tili",
    ],
    "7-sinf": [
        "Ona tili", "Adabiyot (1-qism)", "Adabiyot (2-qism)",
        "Algebra", "Geometriya", "Fizika", "Kimyo",
        "Biologiya (Zoologiya)", "Geografiya",
        "O'zbekiston tarixi", "Jahon tarixi",
        "Informatika va AT", "Texnologiya", "Tasviriy san'at",
        "Musiqa madaniyati", "Tarbiya", "Jismoniy tarbiya",
        "Ingliz tili", "Rus tili",
    ],
    "8-sinf": [
        "Ona tili", "Adabiyot (1-qism)", "Adabiyot (2-qism)",
        "Algebra", "Geometriya", "Fizika", "Kimyo",
        "Biologiya (Odam va uning salomatligi)",
        "Geografiya (O'zbekistonning tabiiy geografiyasi)",
        "O'zbekiston tarixi", "Jahon tarixi",
        "Davlat va huquq asoslari",
        "Informatika va AT", "Chizmachilik", "Tarbiya",
        "Jismoniy tarbiya", "Ingliz tili", "Rus tili",
    ],
    "9-sinf": [
        "Ona tili", "Adabiyot (1-qism)", "Adabiyot (2-qism)",
        "Algebra", "Geometriya", "Fizika", "Kimyo",
        "Biologiya (Sitologiya va genetika asoslari)",
        "Geografiya (O'zbekistonning iqtisodiy geografiyasi)",
        "O'zbekiston tarixi", "Jahon tarixi",
        "Konstitutsiyaviy huquq asoslari",
        "Informatika va AT", "Chizmachilik", "Tarbiya",
        "Jismoniy tarbiya", "CHQBT", "Ingliz tili", "Rus tili",
    ],
    "10-sinf": [
        "Ona tili", "Adabiyot (1-qism)", "Adabiyot (2-qism)",
        "Algebra", "Geometriya", "Fizika", "Kimyo",
        "Biologiya (Umumiy biologiya)",
        "Geografiya (Dunyoning iqtisodiy geografiyasi)",
        "O'zbekiston tarixi", "Jahon tarixi", "Huquqshunoslik",
        "Informatika va AT", "Tarbiya", "Jismoniy tarbiya",
        "CHQBT", "Ingliz tili", "Rus tili", "Tadbirkorlik asoslari",
    ],
    "11-sinf": [
        "Ona tili", "Adabiyot (1-qism)", "Adabiyot (2-qism)",
        "Algebra", "Geometriya", "Fizika", "Kimyo", "Biologiya",
        "Geografiya (Global ekologik muammolar)",
        "O'zbekiston tarixi (Eng yangi tarix)",
        "Jahon tarixi (Eng yangi tarix)", "Inson va jamiyat",
        "Informatika va AT", "Tarbiya", "Jismoniy tarbiya",
        "CHQBT", "Ingliz tili", "Rus tili",
    ],
}

def main_kb():
    return ReplyKeyboardMarkup([
        ["📚 Maktab darsliklari", "📖 Badiiy kitoblar"],
        ["📋 Pedagogik qo'llanmalar"],
        ["🖼️ Rasm tahrirlash", "📄 PDF yasash"],
        ["📝 Word yasash", "💬 Fikr qoldirish"],
    ], resize_keyboard=True)

def back_kb():
    return ReplyKeyboardMarkup([["🔙 Orqaga"]], resize_keyboard=True)

def sinf_kb():
    return ReplyKeyboardMarkup([
        ["1-sinf", "2-sinf", "3-sinf", "4-sinf"],
        ["5-sinf", "6-sinf", "7-sinf", "8-sinf"],
        ["9-sinf", "10-sinf", "11-sinf"],
        ["🔙 Orqaga"]
    ], resize_keyboard=True)

def kitoblar_kb(sinf):
    kitoblar = DARSLIKLAR.get(sinf, [])
    rows = []
    for i in range(0, len(kitoblar), 2):
        rows.append(kitoblar[i:i+2])
    rows.append(["🔙 Orqaga"])
    return ReplyKeyboardMarkup(rows, resize_keyboard=True)

def kitob_kb():
    return ReplyKeyboardMarkup([
        ["Roman", "Hikoya", "She'riyat"],
        ["Dunyo adabiyoti", "O'zbek adabiyoti"],
        ["🔙 Orqaga"]
    ], resize_keyboard=True)

def ped_kb():
    return ReplyKeyboardMarkup([
        ["Dars ishlanmalar", "Texnologik xaritalar"],
        ["Testlar", "Prezentatsiyalar"],
        ["🔙 Orqaga"]
    ], resize_keyboard=True)

def pdf_kb():
    return ReplyKeyboardMarkup([["✅ Tayyor, PDF qil"], ["🔙 Orqaga"]], resize_keyboard=True)

def word_kb():
    return ReplyKeyboardMarkup([["✅ Tayyor, Word qil"], ["🔙 Orqaga"]], resize_keyboard=True)

def admin_kb():
    return ReplyKeyboardMarkup([
        ["📤 Kitob yuklash", "📋 Yuklangan kitoblar"],
        ["📨 Xabar yuborish"],
        ["🔙 Chiqish"]
    ], resize_keyboard=True)

async def check_sub(user_id, bot):
    try:
        m = await bot.get_chat_member("@pedagoglar1226", user_id)
        if m.status in ["left", "kicked"]:
            return False
    except:
        return False
    return True

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    if user.id == ADMIN_ID:
        context.user_data.clear()
        await update.message.reply_text("👑 Admin paneliga xush kelibsiz!", reply_markup=admin_kb())
        return
    subbed = await check_sub(user.id, context.bot)
    if not subbed:
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("➡️ Ilk Bilim", url="https://t.me/+aN43lpwgyEU5NThi")],
            [InlineKeyboardButton("➡️ Pedagoglar", url="https://t.me/pedagoglar1226")],
            [InlineKeyboardButton("✅ Obuna bo'ldim", callback_data="check_sub")],
        ])
        await update.message.reply_text(
            f"👋 Salom {user.first_name}!\n\nBotdan foydalanish uchun kanallarga obuna bo'ling:",
            reply_markup=kb
        )
    else:
        context.user_data.clear()
        await update.message.reply_text(
            f"👋 Salom {user.first_name}! Xush kelibsiz!\n\nBo'limni tanlang:",
            reply_markup=main_kb()
        )

async def callback_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == "check_sub":
        subbed = await check_sub(query.from_user.id, context.bot)
        if subbed:
            context.user_data.clear()
            await query.message.reply_text("✅ Rahmat! Xush kelibsiz!", reply_markup=main_kb())
        else:
            await query.answer("❌ Hali obuna bo'lmadingiz!", show_alert=True)

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    if user.id != ADMIN_ID:
        return
    mode = context.user_data.get("mode")
    if mode == "admin_yuklash":
        nom = context.user_data.get("kitob_nom", "")
        if not nom:
            await update.message.reply_text("❌ Avval kitob nomini yozing!")
            return
        doc = update.message.document
        file_storage[nom] = doc.file_id
        await update.message.reply_text(
            f"✅ '{nom}' saqlandi!\nJami: {len(file_storage)} ta kitob",
            reply_markup=admin_kb()
        )
        context.user_data.clear()

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text or ""
    user = update.effective_user
    mode = context.user_data.get("mode", "main")

    # ADMIN
    if user.id == ADMIN_ID:
        if text == "🔙 Chiqish":
            context.user_data.clear()
            await update.message.reply_text("Admin panel:", reply_markup=admin_kb())
            return
        if text == "📤 Kitob yuklash":
            context.user_data["mode"] = "admin_nom"
            await update.message.reply_text(
                "Kitob nomini yozing:\n(Masalan: 1-sinf Matematika 1-qism)",
                reply_markup=back_kb()
            )
            return
        if text == "📋 Yuklangan kitoblar":
            if not file_storage:
                await update.message.reply_text("📋 Hali kitob yuklanmagan!", reply_markup=admin_kb())
            else:
                msg = "📋 Yuklangan kitoblar:\n\n"
                for i, nom in enumerate(file_storage.keys(), 1):
                    msg += f"{i}. {nom}\n"
                await update.message.reply_text(msg, reply_markup=admin_kb())
            return
        if text == "📨 Xabar yuborish":
            context.user_data["mode"] = "admin_xabar"
            await update.message.reply_text(
                "📨 Foydalanuvchilarga yuboriladigan xabarni yozing:",
                reply_markup=back_kb()
            )
            return
        if mode == "admin_nom":
            context.user_data["kitob_nom"] = text
            context.user_data["mode"] = "admin_yuklash"
            await update.message.reply_text(
                f"✅ Nom: '{text}'\nEndi PDF faylni yuboring:",
                reply_markup=back_kb()
            )
            return
        if mode == "admin_xabar":
            await update.message.reply_text("⏳ Xabar yuborilmoqda...")
            await update.message.reply_text(
                f"✅ Xabar yuborildi!\n\nMazмun:\n{text}",
                reply_markup=admin_kb()
            )
            context.user_data.clear()
            return
        if text == "🔙 Orqaga":
            context.user_data.clear()
            await update.message.reply_text("Admin panel:", reply_markup=admin_kb())
            return

    # FOYDALANUVCHI
    subbed = await check_sub(user.id, context.bot)
    if not subbed:
        await update.message.reply_text("❌ Avval kanallarga obuna bo'ling! /start")
        return

    if text == "🔙 Orqaga":
        if mode == "kitob_list":
            context.user_data["mode"] = "darsliklar"
            await update.message.reply_text("📚 Sinfni tanlang:", reply_markup=sinf_kb())
        else:
            context.user_data.clear()
            await update.message.reply_text("Asosiy menyu:", reply_markup=main_kb())
        return

    if text == "📚 Maktab darsliklari":
        context.user_data["mode"] = "darsliklar"
        await update.message.reply_text("📚 Sinfni tanlang:", reply_markup=sinf_kb())
        return

    if text == "📖 Badiiy kitoblar":
        context.user_data["mode"] = "kitoblar"
        await update.message.reply_text("📖 Kitob turini tanlang:", reply_markup=kitob_kb())
        return

    if text == "📋 Pedagogik qo'llanmalar":
        context.user_data["mode"] = "pedagogik"
        await update.message.reply_text("📋 Bo'limni tanlang:", reply_markup=ped_kb())
        return

    if text == "🖼️ Rasm tahrirlash":
        context.user_data["mode"] = "tahrirl"
        await update.message.reply_text("🖼️ Rasmni yuboring:", reply_markup=back_kb())
        return

    if text == "📄 PDF yasash":
        context.user_data["mode"] = "pdf"
        context.user_data["pdf_images"] = []
        await update.message.reply_text(
            "📄 Rasmlarni yuboring.\nTugagach '✅ Tayyor, PDF qil' bosing:",
            reply_markup=pdf_kb()
        )
        return

    if text == "📝 Word yasash":
        context.user_data["mode"] = "word"
        context.user_data["word_items"] = []
        await update.message.reply_text(
            "📝 Rasm yoki matn yuboring.\nTugagach '✅ Tayyor, Word qil' bosing:",
            reply_markup=word_kb()
        )
        return

    if text == "💬 Fikr qoldirish":
        context.user_data["mode"] = "fikr"
        await update.message.reply_text(
            "💬 Fikringizni yozing, adminга yetkazamiz:",
            reply_markup=back_kb()
        )
        return

    # FIKR
    if mode == "fikr":
        user_info = f"@{user.username}" if user.username else user.first_name
        await context.bot.send_message(
            chat_id=ADMIN_ID,
            text=f"💬 Yangi fikr!\n\n👤 Foydalanuvchi: {user_info} (ID: {user.id})\n\n📝 Fikr:\n{text}"
        )
        await update.message.reply_text(
            "✅ Fikringiz adminga yetkazildi! Rahmat!",
            reply_markup=main_kb()
        )
        context.user_data.clear()
        return

    # SINF TANLANDI
    if mode == "darsliklar" and text in DARSLIKLAR:
        context.user_data["mode"] = "kitob_list"
        context.user_data["sinf"] = text
        await update.message.reply_text(
            f"📚 {text} — kitobni tanlang:",
            reply_markup=kitoblar_kb(text)
        )
        return

    # KITOB TANLANDI
    if mode == "kitob_list":
        sinf = context.user_data.get("sinf", "")
        kitoblar = DARSLIKLAR.get(sinf, [])
        if text in kitoblar:
            storage_key = f"{sinf} {text}"
            if storage_key in file_storage:
                await update.message.reply_text(f"📥 '{text}' yuklanmoqda...")
                await context.bot.send_document(
                    chat_id=update.effective_chat.id,
                    document=file_storage[storage_key],
                    caption=f"📚 {sinf} | {text}"
                )
            else:
                await update.message.reply_text(
                    f"📚 {sinf} | {text}\n\n⏳ Bu kitob hali yuklanmagan.\n📥 So'rash uchun: @pedagoglar1226",
                    reply_markup=kitoblar_kb(sinf)
                )
            return

    # KITOBLAR
    if mode == "kitoblar" and text in ["Roman", "Hikoya", "She'riyat", "Dunyo adabiyoti", "O'zbek adabiyoti"]:
        await update.message.reply_text(
            f"📖 {text}\n\n📥 Kitob so'rash: @pedagoglar1226",
            reply_markup=kitob_kb()
        )
        return

    # PEDAGOGIK
    if mode == "pedagogik" and text in ["Dars ishlanmalar", "Texnologik xaritalar", "Testlar", "Prezentatsiyalar"]:
        await update.message.reply_text(
            f"📋 {text}\n\n📥 Yuklash uchun: @pedagoglar1226",
            reply_markup=ped_kb()
        )
        return

    # PDF
    if text == "✅ Tayyor, PDF qil" and mode == "pdf":
        images = context.user_data.get("pdf_images", [])
        if not images:
            await update.message.reply_text("❌ Rasm yuborilmadi!", reply_markup=pdf_kb())
            return
        context.user_data["mode"] = "pdf_nom"
        await update.message.reply_text("📄 PDF uchun nom kiriting:", reply_markup=ReplyKeyboardRemove())
        return

    if mode == "pdf_nom":
        nom = text.strip().replace(" ", "_") or "Hujjat"
        images = context.user_data.get("pdf_images", [])
        await update.message.reply_text("⏳ PDF tayyorlanmoqda...")
        buf = io.BytesIO()
        c = canvas.Canvas(buf, pagesize=A4)
        pw, ph = A4
        for img_bytes in images:
            img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
            iw, ih = img.size
            ratio = min(pw/iw, ph/ih) * 0.95
            tmp = io.BytesIO()
            img.save(tmp, format="JPEG")
            tmp.seek(0)
            c.drawImage(ImageReader(tmp), (pw-iw*ratio)/2, (ph-ih*ratio)/2, iw*ratio, ih*ratio)
            c.showPage()
        c.save()
        buf.seek(0)
        from telegram import InputFile
        await update.message.reply_document(InputFile(buf, f"{nom}.pdf"), caption=f"✅ {nom}.pdf tayyor!")
        context.user_data.clear()
        await update.message.reply_text("Asosiy menyu:", reply_markup=main_kb())
        return

    # WORD
    if text == "✅ Tayyor, Word qil" and mode == "word":
        items = context.user_data.get("word_items", [])
        if not items:
            await update.message.reply_text("❌ Hech narsa yuborilmadi!", reply_markup=word_kb())
            return
        context.user_data["mode"] = "word_nom"
        await update.message.reply_text("📝 Word uchun nom kiriting:", reply_markup=ReplyKeyboardRemove())
        return

    if mode == "word_nom":
        nom = text.strip().replace(" ", "_") or "Hujjat"
        items = context.user_data.get("word_items", [])
        await update.message.reply_text("⏳ Word tayyorlanmoqda...")
        doc = Document()
        doc.add_heading(nom.replace("_", " "), 0)
        for item in items:
            if item["type"] == "text":
                p = doc.add_paragraph(item["content"])
                p.style.font.size = Pt(12)
            elif item["type"] == "image":
                try:
                    doc.add_picture(io.BytesIO(item["content"]), width=Inches(5))
                except:
                    doc.add_paragraph("[Rasm]")
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        from telegram import InputFile
        await update.message.reply_document(InputFile(buf, f"{nom}.docx"), caption=f"✅ {nom}.docx tayyor!")
        context.user_data.clear()
        await update.message.reply_text("Asosiy menyu:", reply_markup=main_kb())
        return

    if mode == "word" and text:
        items = context.user_data.get("word_items", [])
        items.append({"type": "text", "content": text})
        context.user_data["word_items"] = items
        await update.message.reply_text(f"✅ Matn qo'shildi! Jami: {len(items)} ta", reply_markup=word_kb())
        return

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    subbed = await check_sub(user.id, context.bot)
    if not subbed and user.id != ADMIN_ID:
        await update.message.reply_text("❌ Avval kanallarga obuna bo'ling! /start")
        return

    mode = context.user_data.get("mode", "main")
    photo = update.message.photo[-1]
    file = await context.bot.get_file(photo.file_id)
    file_bytes = bytes(await file.download_as_bytearray())

    if mode == "tahrirl":
        img = Image.open(io.BytesIO(file_bytes)).convert("RGB")
        img = ImageEnhance.Brightness(img).enhance(1.2)
        img = ImageEnhance.Contrast(img).enhance(1.3)
        img = img.filter(ImageFilter.SHARPEN)
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=95)
        buf.seek(0)
        from telegram import InputFile
        await update.message.reply_photo(InputFile(buf, "tahrirlangan.jpg"), caption="✅ Rasm tahrirlandi!")
        return

    if mode == "pdf":
        images = context.user_data.get("pdf_images", [])
        images.append(file_bytes)
        context.user_data["pdf_images"] = images
        await update.message.reply_text(f"✅ Rasm qo'shildi! Jami: {len(images)} ta", reply_markup=pdf_kb())
        return

    if mode == "word":
        items = context.user_data.get("word_items", [])
        items.append({"type": "image", "content": file_bytes})
        context.user_data["word_items"] = items
        await update.message.reply_text(f"✅ Rasm qo'shildi! Jami: {len(items)} ta", reply_markup=word_kb())
        return

    await update.message.reply_text("Asosiy menyudan bo'lim tanlang:", reply_markup=main_kb())

def main():
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CallbackQueryHandler(callback_handler))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    print("Bot ishga tushdi!")
    app.run_polling()

if __name__ == "__main__":
    main()
