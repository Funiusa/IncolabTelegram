from aiogram import types, Dispatcher
from asyncio import sleep
from os import path


async def get_file_from_user(message: types.Message):
    if document := message.document:
        await document.download(destination_file=document.file_name)
        await message.answer(f"<b>{message.from_user.first_name}</b>, файл <u>{document.file_name}</u> получен.",
                             parse_mode="HTML")
    else:
        await message.answer("Файл не найден. Попробуй еще раз.")


async def upload_new_file(message: types.Message):
    """ Button for upload the compleat file"""
    try:
        path_to_new_file = path.abspath('./new.xlsx')  # TODO create renamer for new file
        new_file = open(path_to_new_file, 'rb')
        await message.answer("Вот ваш файл")
        await sleep(1)
        await message.answer_document(document=new_file)
        new_file.close()
    except FileNotFoundError:
        await message.answer("Новый файл еще <u>не создан</u>.", parse_mode="HTML")
    except Exception as e:
        await message.answer(f"<b>Ошибка:</b> {e}.", parse_mode="HTML")


def register_files_handler(dp: Dispatcher):
    # dp.register_message_handler(get_file_from_user, content_types=ContentTypes.DOCUMENT, state='*')
    dp.register_message_handler(upload_new_file, commands='file', state='*')
