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

# 🔹 Fayldan token va adminlarni olish
load_dotenv()
BOT_TOKEN = os.getenv("BOT_TOKEN")
ADMIN_IDS = [int(x) for x in os.getenv("ADMIN_IDS", "").split(",") if x.strip().isdigit()]
EXCEL_FILE = "users.xlsx"

if not BOT_TOKEN:
    raise ValueError("❌ BOT_TOKEN topilmadi .env faylda!")

# 🔹 Import handlers
from handlers import start


# 📊 --- Hisobot yuborish funksiyasi ---
async def send_excel_report(bot: Bot):
    if not os.path.exists(EXCEL_FILE):
        print("⚠️ Excel fayl topilmadi!")
        return

    for admin_id in ADMIN_IDS:
        try:
            file = FSInputFile(EXCEL_FILE)
            await bot.send_document(
                admin_id,
                file,
                caption="📊 Bugungi foydalanuvchilar hisobot fayli"
            )
            print(f"✅ Hisobot yuborildi -> {admin_id}")
        except Exception as e:
            print(f"⚠️ Admin {admin_id} ga yuborishda xato: {e}")


# 🔁 --- Jadval qo‘shish (08:00 va 20:00) ---
def setup_scheduler(bot: Bot):
    scheduler = AsyncIOScheduler(timezone="Asia/Tashkent")
    scheduler.add_job(send_excel_report, "cron", hour=8, minute=0, args=[bot])
    scheduler.add_job(send_excel_report, "cron", hour=20, minute=0, args=[bot])
    scheduler.start()
    print("⏰ Avtomatik hisobot yuborish yo‘lga qo‘yildi.")


# 🌐 --- Mini web-server (keep-alive uchun) ---
async def handle(request):
    return web.Response(text="✅ Bot is running!")

async def start_keep_alive():
    app = web.Application()
    app.router.add_get("/", handle)
    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", 8080)
    await site.start()
    print("🌍 Keep-alive web-server ishga tushdi (port 8080)")


# 🚀 --- Botni ishga tushirish ---
async def run_bot():
    logging.basicConfig(level=logging.INFO)
    bot = Bot(token=BOT_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    dp = Dispatcher()
    dp.include_router(start.router)
    setup_scheduler(bot)
    print("🤖 Bot ishga tushdi...")

    # Polling’ni avtomatik qayta ishga tushirish bilan
    while True:
        try:
            await dp.start_polling(bot)
        except Exception as e:
            logging.error(f"❌ Polling xatosi: {e}")
            await asyncio.sleep(10)  # qayta urinishdan oldin kutish


async def main():
    await asyncio.gather(
        start_keep_alive(),  # keep-alive server
        run_bot()            # bot
    )


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except (KeyboardInterrupt, SystemExit):
        print("❌ Bot to‘xtatildi.")
