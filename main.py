import asyncio
import logging
import os
from aiogram import Bot, Dispatcher, types
from aiogram.filters import Command
from aiogram.types import FSInputFile
from config import TOKEN, ADMIN_ID
from excel_to_xml import excel_to_xml  # Импорт вашей функции конвертации


bot = Bot(token=TOKEN)
dp = Dispatcher()

# Пути к файлам
TEMPLATE_PATH = "template.xlsx"
DOWNLOAD_DIR = "downloads"
OUTPUT_DIR = "output"

os.makedirs(DOWNLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


@dp.message(Command("start"))
async def start_cmd(message: types.Message):
    # Отправляем сообщение пользователю
    await message.answer(
        "Здравствуйте. Это бот для формирования xml-файла. "
        "Заполните пжл шаблон и отправьте мне заполненный."
    )

    # Отправляем шаблон
    template = FSInputFile(TEMPLATE_PATH)
    await message.answer_document(template)

    # Уведомление админу
    await bot.send_message(
        ADMIN_ID,
        f"Новый пользователь:\n"
        f"ID: {message.from_user.id}\n"
        f"Username: @{message.from_user.username}"
    )


@dp.message()
async def handle_docs(message: types.Message):
    if not message.document:
        await message.answer("Пожалуйста, отправьте заполненный Excel-файл")
        return

    # Проверка расширения файла
    if not message.document.file_name.endswith(('.xlsx', '.xls')):
        await message.answer("Пожалуйста, отправьте файл в формате Excel (.xlsx)")
        return

    try:
    # Скачивание файла
        file_id = message.document.file_id
        file = await bot.get_file(file_id)
        file_path = file.file_path

        download_path = os.path.join(DOWNLOAD_DIR, message.document.file_name)
        await bot.download_file(file_path, destination=download_path)

        # Обработка файла
        output_filename = os.path.splitext(message.document.file_name)[0] + ".xml"
        output_path = os.path.join(OUTPUT_DIR, output_filename)

        excel_to_xml(download_path, output_path)

        # Отправка результата
        result_file = FSInputFile(output_path, filename=output_filename)
        await message.answer_document(result_file)

    except Exception as e:
        await message.answer(f"Ошибка обработки файла: {str(e)}")
    finally:
        # Очистка временных файлов
        if os.path.exists(download_path):
            os.remove(download_path)
        if os.path.exists(output_path):
            os.remove(output_path)

logger = logging.getLogger(__name__)


async def main():
    logging.basicConfig(level=logging.INFO,
                        format='%(filename)s:%(lineno)d %(levelname)-8s [%(asctime)s] - %(name)s - %(message)s')
    logging.info('Starting bot')
    bot = Bot(token=TOKEN)
    await bot.delete_webhook(drop_pending_updates=True)
    await dp.start_polling(bot)


if __name__ == '__main__':
    asyncio.run(main())

