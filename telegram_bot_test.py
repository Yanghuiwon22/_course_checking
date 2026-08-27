import asyncio
import telegram

token = "8779438981:AAHs3LWbKTSeb3tFbJY9dupMLkGKX_R-4fY"
chat_id = 5522728409

import asyncio
from telegram import Update
from telegram.ext import Application, MessageHandler, filters, ContextTypes


async def handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_text = update.message.text

    if user_text == "안녕":
        await update.message.reply_text("어 그래 안녕")
    elif user_text == "뭐해":
        await update.message.reply_text("그냥 있어")

async def post_init(application):
    await application.bot.send_message(chat_id=chat_id, text="테스트 중입니다.")

app = Application.builder().token(token).post_init(post_init).build()
app.add_handler(MessageHandler(filters.TEXT, handler))
app.run_polling()  # ✅ asyncio.run() 없이 직접 호출