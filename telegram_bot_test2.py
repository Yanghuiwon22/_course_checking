import asyncio
import telegram

token = "8779438981:AAHs3LWbKTSeb3tFbJY9dupMLkGKX_R-4fY"
chat_id = 5522728409


async def send_daily_message(token, chat_id):
    # 실행시킬 함수명 임의지정
    bot = telegram.Bot(token=token)

    message = "당신이 오늘 진행한 업무 또는 진행 해야할 업무는?"
    await bot.send_message(chat_id, message)

async def main():
    # ✅ async 추가
    await send_daily_message(token, chat_id)
    # ✅ await 추가
if __name__ == '__main__':
    asyncio.run(main( ))
