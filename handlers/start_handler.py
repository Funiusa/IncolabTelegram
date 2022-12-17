from os import path

from aiogram import Dispatcher, types
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters.state import StatesGroup, State
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton
import aiogram.utils.markdown as md
from asyncio import sleep
from .xlsx_handler.wagons_handler import wagons_xlsx


async def cmd_start(message: types.Message):
    kb = [
        [
            KeyboardButton(text="Run"),
        ],
    ]
    keyboard = ReplyKeyboardMarkup(
        keyboard=kb,
        resize_keyboard=True,
        input_field_placeholder="help для помощи",
    )
    await message.answer("Нажмите кнопку <b>Run</b>!", reply_markup=keyboard, parse_mode="HTML")



async def get_wagons(message: types.Message):
    try:
        wagons = [int(wagon) for wagon in message.text.split('\n') if wagon.isdigit()]
        if wagons:
            await message.answer(f"Вагоны {wagons} добавлены!")
    except Exception as e:
        await message.answer(f"<b>Ошибка: </b>{e}.")


class SearchWagons(StatesGroup):
    waiting_for_file_upload = State()
    waiting_for_wagons_list = State()
    waiting_for_name = State()


async def start_calculation(message: types.Message, state: FSMContext):
    await message.answer("Добавьте основной файл."
                         "\n<i>При вводе нескольких основных файлов работа будет произведена с <b>последним</b>.</i>",
                         reply_markup=types.ReplyKeyboardRemove(), parse_mode="HTML")
    await state.set_state(SearchWagons.waiting_for_file_upload.state)


async def upload_file(message: types.Message, state: FSMContext):
    if document := message.document:
        await state.update_data(main=document)
        user_data = await state.get_data()
        if user_data['main']['mime_type'] == 'application/wps-office.xlsx':
            await document.download(destination_file=document.file_name)  # download main file
            await state.update_data(main=document.file_name)
        else:
            await message.answer(f"Не верный формат файла. Необходим .xlsx")
            return

    await state.set_state(SearchWagons.waiting_for_wagons_list.state)
    await message.answer(f"Файл <b>{document.file_name}</b>, получен."
                         f"\n\nТеперь Введите список вагонов.", parse_mode="HTML")


async def wagons_list(message: types.Message, state: FSMContext):
    wagons = [int(wagon) for wagon in message.text.split('\n') if wagon.isdigit()]
    if len(wagons) == 0:
        await message.answer(md.text("Список вагонов должен быть числовой.",
                                     "\nВ столбик, без запятых и других знаков.",
                                     "\n\n<b>Например так:</b>"
                                     "\n452345324",
                                     "\n85678678",
                                     "\n27856675  -  это 3 вагона",
                                     "\n<b>Или даже так:</b>",
                                     "\ndfsgfgdsgf",
                                     "\n345354663"
                                     "\ngsdfg5735 - это 1 вагон."), parse_mode="HTML")
        await sleep(1)
        return
    await state.update_data(wagons=wagons)
    await message.answer(f"Вагоны <b>{wagons}</b> добавлены!"
                         "\n\nТеперь укажите имя нового файла.",
                         parse_mode="HTML")
    await state.set_state(SearchWagons.waiting_for_name.state)


async def new_file_and_calculate(message: types.Message, state: FSMContext):
    await state.update_data(new=message.text.lower() + '.xlsx')
    user_data = await state.get_data()
    await message.answer(f"Новый файл будет называться - <b>{user_data.get('new')}</b>.",
                         parse_mode='HTML')
    await message.answer(f"{user_data.get('wagons')}")
    out = wagons_xlsx(user_data.get('main'), user_data.get('new'), user_data.get('wagons'))

    if out:
        for key, val in out.items():
            if key == 'not_found':
                await message.answer(f"Вагоны с номерами {val} не найдены в основном файле.")
            else:
                await message.answer(f"Вагоны {val} уже записаны в {user_data.get('new')}.")
    await message.answer(f"Готово")
    try:
        path_to_new_file = path.abspath(user_data.get('new'))  # TODO create renamer for new file
        new_file = open(path_to_new_file, 'rb')
        await sleep(1)
        await message.answer_document(document=new_file)
        new_file.close()
    except FileNotFoundError:
        await message.answer("Новый файл еще <u>не создан</u>.", parse_mode="HTML")
    except Exception as e:
        await message.answer(f"<b>Ошибка:</b> {e}.", parse_mode="HTML")
    await state.finish()


def register_start_handler(dp: Dispatcher):
    # dp.register_message_handler(cmd_start, commands='start', state='*')
    # dp.register_message_handler(answer_on_start, Text('Run'), state='*')
    # dp.register_message_handler(get_wagons, state='*')

    dp.register_message_handler(start_calculation, commands='run', state="*")
    dp.register_message_handler(upload_file,
                                content_types='document', state=SearchWagons.waiting_for_file_upload)
    dp.register_message_handler(wagons_list, state=SearchWagons.waiting_for_wagons_list)
    dp.register_message_handler(new_file_and_calculate, state=SearchWagons.waiting_for_name)
