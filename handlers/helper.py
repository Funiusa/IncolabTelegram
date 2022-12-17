from aiogram import Dispatcher, types
import aiogram.utils.markdown as md

commands_list = md.text(
    "<b>Доступные команды:</b>\n",
    "/help - что делать и кто виноват",
    "/run - начать обработку",
    "/file - скачать готовый файл",
    sep="\n"
)
url = md.hlink("Кортни Кокс", "https://ru.wikipedia.org/wiki/%D0%9A%D0%BE%D0%BA%D1%81")

help_text = md.text("\n<b>Это описание процесса обработки файла.</b>",
                    "\nПосле нажатия кнопки <u>run</u>",
                    "тебе нужно загрузить в бот основной файл с вагонами (название_файла.xlsx).",
                    "Затем внеси список вагонов, в столбик, для поиска и расчета.",
                    "Придумай и внеси название нового файла или пропусти этот пункт, ",
                    "но тогда получившийся файл будет называться  просто <b>new</b>.",
                    "<b>Готово.</b> \nТвой файл появится в боте и ты сможешь скачать его.",
                    "\nЕсли потребуется еще раз скачать файл, используй комманду /file.",
                    "Удачи!",
                    sep="\n"
                    )


async def help_message(message: types.Message):
    await message.answer(f"И так, {message.from_user.first_name}! {help_text}", parse_mode="HTML")


welcome = md.text("\nМеня зовут <b>WagonBot</b>",
                  "\nОчень рад помочь.",
                  "Надеюсь, наше взаимодействие будет интересным и полезным.",
                  "Если ты <s>лузер</s> <u>впервые</u> пользуешься ботом,",
                  "то советую тебе, воспользоваться коммандой <u>/help</u>.",

                  "\nНо, если ты <s>отчаянная домохозяйка</s>, <s>тертый калач</s>,",
                  "знойный инспектор и знаешь, не по наслышке,",
                  f"<tg-spoiler>{url}</tg-spoiler>? Cмело жми /run!",
                  sep="\n"
                  )


async def welcome_message(message: types.Message):
    await message.answer(f"Привет, {message.from_user.first_name}! {welcome}", parse_mode="HTML")


async def all_commands(message: types.Message):
    await message.answer(commands_list, parse_mode="HTML")


def register_help_handlers(dp: Dispatcher):
    dp.register_message_handler(welcome_message, commands='start', state='*')
    dp.register_message_handler(help_message, commands='help', state='*')
    dp.register_message_handler(all_commands, commands='all', state='*')
