"""
Rasmni PDF ga o'giruvchi Telegram bot
--------------------------------------
Foydalanuvchi botga rasm yuborsa, bot uni PDF formatiga o'girib qaytaradi.
Render.com da 24/7 ishlashi uchun Flask keep-alive serveri qo'shilgan.
"""

import os
import io
import img2pdf
import telebot
from threading import Thread
from flask import Flask

# --- Render portini tinglash uchun Flask server (Keep-Alive) ---
app = Flask('')

@app.route('/')
def home():
    return "Bot is running 24/7!"

def run():
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 8080)))

def keep_alive():
    t = Thread(target=run)
    t.start()
# -------------------------------------------------------------

# Render Environment Variables ichidan BOT_TOKEN ni oladi
BOT_TOKEN = os.environ.get("BOT_TOKEN")

if not BOT_TOKEN:
    print("DIQQAT: BOT_TOKEN topilmadi! Render muhitida BOT_TOKEN sozlanganini tekshiring.")

bot = telebot.TeleBot(BOT_TOKEN)


@bot.message_handler(commands=["start", "help"])
def send_welcome(message):
    bot.reply_to(
        message,
        "Salom! Menga bitta rasm yuboring — men uni sizga PDF fayl "
        "qilib qaytarib beraman. 📄"
    )


def convert_and_send(message, image_bytes: bytes, filename_hint: str = "rasm"):
    try:
        pdf_bytes = img2pdf.convert(image_bytes)
        pdf_file = io.BytesIO(pdf_bytes)
        pdf_file.name = f"{filename_hint}.pdf"
        bot.send_document(message.chat.id, pdf_file, visible_file_name=pdf_file.name)
    except Exception as e:
        bot.reply_to(message, f"Kechirasiz, xatolik yuz berdi: {e}")


@bot.message_handler(content_types=["photo"])
def handle_photo(message):
    file_info = bot.get_file(message.photo[-1].file_id)
    downloaded = bot.download_file(file_info.file_path)
    convert_and_send(message, downloaded)


@bot.message_handler(content_types=["document"])
def handle_document(message):
    mime = message.document.mime_type or ""
    if not mime.startswith("image/"):
        bot.reply_to(message, "Iltimos, rasm fayl yuboring (jpg, png va h.k.).")
        return

    file_info = bot.get_file(message.document.file_id)
    downloaded = bot.download_file(file_info.file_path)
    name_hint = os.path.splitext(message.document.file_name or "rasm")[0]
    convert_and_send(message, downloaded, name_hint)


if __name__ == "__main__":
    print("Bot ishga tushmoqda...")
    keep_alive()  # Flask serverini fon rejimida yoqish
    bot.infinity_polling()
