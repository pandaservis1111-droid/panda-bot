import asyncio
import logging
import os
from aiogram import Bot, Dispatcher
from aiogram.enums import ParseMode
from aiogram.client.default import DefaultBotProperties
from dotenv import load_dotenv
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from aiogram.types import FSInputFile
from aiohttp import web

# 🔹 Load .env
load_dotenv()
BOT_TOKEN = os.getenv("BOT_TOKEN")
ADMIN_IDS = [int(x) for x in os.getenv("ADMIN_IDS", "").split(",") if x.strip().isdigit()]
GROUP_IDS = [int(x) for x in os.getenv("GROUP_IDS", "").split(",") if x.strip().isdigit()]
EXCEL_FILE = "users.xlsx"
WEBHOOK_PATH = f"/webhook/{BOT_TOKEN}"
WEBHOOK_URL = os.getenv("WEBHOOK_URL")  # твой URL на Render, например https://yourapp.onrender.com

if not BOT_TOKEN or not WEBHOOK_URL:
    raise ValueError("❌ BOT_TOKEN yoki WEBHOOK_URL topilmadi .env faylda!")

# 🔹 Import handlers
from handlers import start

# 📊 --- Send report ---
async def send_excel_report(bot: Bot):
    if not os.path.exists(EXCEL_FILE):
        print("⚠️ Excel fayl topilmadi!")
        return

    # Объединяем админов и группы
    recipients = ADMIN_IDS + GROUP_IDS

    for recipient_id in recipients:
        try:
            file = FSInputFile(EXCEL_FILE)
            await bot.send_document(
                recipient_id,
                file,
                caption="📊 Bugungi foydalanuvchilar hisobot fayli"
            )
            print(f"✅ Hisobot yuborildi -> {recipient_id}")
        except Exception as e:
            print(f"⚠️ Ошибка отправки {recipient_id}: {e}")

# 🔁 --- Scheduler ---
def setup_scheduler(bot: Bot):
    scheduler = AsyncIOScheduler(timezone="Asia/Tashkent")
    scheduler.add_job(send_excel_report, "cron", hour=8, minute=0, args=[bot])
    scheduler.add_job(send_excel_report, "cron", hour=20, minute=0, args=[bot])
    scheduler.start()
    print("⏰ Avtomatik hisobot yuborish yo‘lga qo‘yildi.")

# 🌐 --- Webhook server ---
async def handle(request):
    bot = request.app['bot']
    update = await request.json()
    try:
        from aiogram.types import Update
        telegram_update = Update(**update)
        await bot.dispatch(bot.dp, telegram_update)
    except Exception as e:
        logging.error(f"Webhook xato: {e}")
    return web.Response(text="ok")

# 🔹 Keep-alive endpoint
async def keep_alive(request):
    return web.Response(text="✅ Bot is running!")

async def on_startup(app):
    bot = app['bot']
    await bot.delete_webhook()
    await bot.set_webhook(WEBHOOK_URL + WEBHOOK_PATH)
    setup_scheduler(bot)
    print("🤖 Bot ishga tushdi va webhook o‘rnatildi")

async def create_app():
    logging.basicConfig(level=logging.INFO)
    bot = Bot(token=BOT_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    dp = Dispatcher()
    dp.include_router(start.router)
    bot.dp = dp  # save dispatcher for webhook handling

    app = web.Application()
    app['bot'] = bot
    # Webhook endpoint
    app.router.add_post(WEBHOOK_PATH, handle)
    # Keep-alive endpoint
    app.router.add_get("/", keep_alive)
    app.on_startup.append(on_startup)
    return app

if __name__ == "__main__":
    web.run_app(create_app(), host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))
