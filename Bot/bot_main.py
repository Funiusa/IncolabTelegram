import asyncio
import logging
import os
from aiogram import Bot, Dispatcher, types
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.types import BotCommand, chat, InputMedia

from handlers.helper import register_help_handlers
from handlers.start_handler import register_start_handler
from handlers.files_handler import register_files_handler

from dotenv import load_dotenv

load_dotenv()


# Включаем логирование, чтобы не пропустить важные сообщения
logging.basicConfig(level=logging.INFO)


# Регистрация команд, отображаемых в интерфейсе Telegram. Menu
async def set_commands(bot: Bot):
    commands = [
        BotCommand(command="/start", description="Вступление"),
        BotCommand(command="/help", description="Что делать и кто виноват"),
        BotCommand(command="/run", description="Начать обработку"),
        BotCommand(command="/file", description="Скачать готовый файл")
    ]
    await bot.set_my_commands(commands)


# Запуск процесса поллинга новых апдейтов
async def main():
    # Объект бота
    storage = MemoryStorage()
    wago_bot = Bot(token=os.getenv('API_TOKEN'))  # Token in .env
    dp = Dispatcher(wago_bot, storage=storage)

    register_help_handlers(dp)
    register_start_handler(dp)
    register_files_handler(dp)

    await set_commands(wago_bot)  # Set commands in the menu
    await wago_bot.delete_webhook(drop_pending_updates=True)
    await dp.start_polling(wago_bot)


if __name__ == "__main__":
    asyncio.run(main())
