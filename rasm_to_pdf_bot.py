import os
import requests
import telebot

# Token va kalitlarni muhitdan o'qish
BOT_TOKEN = os.getenv("BOT_TOKEN")
GROQ_API_KEY = os.getenv("GROQ_API_KEY")
GROQ_MODEL = "llama-3.3-70b-versatile"
GROQ_API_URL = "https://api.groq.com/openai/v1/chat/completions"

bot = telebot.TeleBot(BOT_TOKEN)

def generate_outline_with_ai(topic, slide_count=5):
    headers = {
        "Authorization": f"Bearer {GROQ_API_KEY}",
        "Content-Type": "application/json",
    }
    
    user_prompt = (
        f"Mavzu: {topic}\n"
        f"Aynan {slide_count} ta kontent slaydi bo'lsin (title slayddan tashqari). "
        "Har bir slaydda 3-5 ta qisqa va aniq fikr (bullet) bo'lsin. "
        "Javob o'zbek tilida bo'lsin."
    )
    
    payload = {
        "model": GROQ_MODEL,
        "messages": [
            {"role": "system", "content": "Siz professional taqdimot tayyorlovchi yordamchisiz."},
            {"role": "user", "content": user_prompt}
        ],
        "temperature": 0.7
    }
    
    response = requests.post(GROQ_API_URL, headers=headers, json=payload)
    response.raise_for_status()
    
    data = response.json()
    return data["choices"][0]["message"]["content"]


# Telegram 400 xatosini oldini oluvchi admin tugmasi funksiyasi namunasi
@bot.callback_query_handler(func=lambda call: True)
def handle_admin_decision(call):
    try:
        bot.edit_message_caption(
            caption=call.message.caption + "\n\n✅ TASDIQLANDI",
            chat_id=call.message.chat.id,
            message_id=call.message.message_id,
            reply_markup=None
        )
    except Exception as e:
        print(f"E'tiborsiz xato (Telegram 400): {e}")

# Botni ishga tushirish
if __name__ == "__main__":
    print("Bot ishga tushmoqda...")
    bot.infinity_polling()
