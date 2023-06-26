from aiogram import Bot, Dispatcher, executor, types
from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton
from aiogram.contrib.fsm_storage.memory import MemoryStorage
import emoji
import time
from docx2pdf import convert
import os

API_TOKEN = 'YOUR_TOKEN'
 
bot = Bot(token=API_TOKEN)
storage = MemoryStorage()
dp = Dispatcher(bot)
Logging = False

# Время последней отправки файла
last_file_sent_time = {}

@dp.message_handler(commands=['start'])
async def send_welcome(message: types.Message):
   inline_keyboard = types.InlineKeyboardMarkup()
   help_button = types.InlineKeyboardButton("Help", callback_data="Help")
   inline_keyboard.add(help_button)

   await message.answer('Приветствую! Я бот, который конвертирует docx в pdf - загрузите файл для начала работы', reply_markup = inline_keyboard)

# Обработчик нажатия на кнопку
@dp.callback_query_handler(lambda callback_query: True)
async def button_click_handler(callback_query: types.CallbackQuery):
   if callback_query.data == "Help":
      await bot.send_message(callback_query.from_user.id,'Данный бот конвертирует файлы формата docx в pdf. \nДля начала работы загрузите файл, используя кнопку "Прикрепить" или путем перетаскивания файла в поле ввода сообщения.')

@dp.message_handler(commands=['help'])
async def send_welcome(message: types.Message):
   await message.reply('Данный бот конвертирует файлы формата docx в pdf. \nДля начала работы загрузите файл, используя кнопку "Прикрепить" или путем перетаскивания файла в поле ввода сообщения.')

@dp.message_handler()
async def message_reply(message: types.Message):
   # Обнаружение эмоджи в тексте
   if emoji.is_emoji(message.text):
      # Если найдены эмоджи, отправляем их обратно
      await message.answer(message.text)
   else:
      await message.reply('Для начала работы отправьте файл с расширением docx')

@dp.message_handler(content_types=[types.ContentType.DOCUMENT, types.ContentType.PHOTO])
async def handle_files(message: types.Message):

   user_id = message.from_user.id
   current_time = time.time()
   
   # Проверяем, сколько времени прошло с момента последней отправки файла
   if user_id in last_file_sent_time and (current_time - last_file_sent_time[user_id]) < 10:
      await message.answer('Пожалуйста, дождитесь 10 секунд перед отправкой следующего файла')
      return

   # Проверяем, является ли сообщение документом
   if message.document:
      # Обновляем время последней отправки файла
      last_file_sent_time[user_id] = current_time

      file_id = message.document.file_id
      file_name = message.document.file_name.rsplit('.', 1)[0]
      file_info = await bot.get_file(file_id)
      file_path = file_info.file_path

      # Получаем расширение файла
      file_extension = file_path.split('.')[-1]

      # Проверяем, что расширение файла соответствует допустимым значениям
      if file_extension.lower() == 'docx':
            # Получаем информацию о размере файла
            file_size = file_info.file_size
            # Проверяем размер файла (до 100 МБ)
            if file_size <= 100 * 1024 * 1024:
               # Сохраняем файл на сервере
               dir = 'Source\\'
               await bot.download_file(file_path, (dir + f'{user_id}.{file_extension}'))
               await message.answer('Файл успешно загружен! Процесс займет несколько секунд')

               # Конвертируем DOCX в PDF
               docx_file = dir + f'{user_id}.{file_extension}'
               pdf_file = dir + f'{file_name}.pdf'
               convert(docx_file, pdf_file)

               # Отправка файла пользователю
               with open(pdf_file, 'rb') as file:
                  await bot.send_document(message.chat.id, file)

               if not Logging:
                  for f in os.listdir(dir):
                     os.remove(os.path.join(dir, f))
            else:
               await message.answer('Размер файла превышает 100 МБ. Пожалуйста, отправьте файл размером до 100 МБ')
      else:
         await message.answer(f'Не удалось обработать файл с расширением {file_extension}. Пожалуйста, отправьте файл с расширением docx')

   # Проверяем, является ли сообщение фотографией
   elif message.photo:
      # Обновляем время последней отправки файла
      last_file_sent_time[user_id] = current_time

      photo = message.photo[-1]
      file_id = photo.file_id
      file_info = await bot.get_file(file_id)
      file_path = file_info.file_path

      # Получаем расширение файла
      file_extension = file_path.split('.')[-1]
      await message.answer(f'Не удалось обработать файл с расширением {file_extension}. Пожалуйста, отправьте файл с расширением docx')
 
if __name__ == '__main__':
   executor.start_polling(dp, skip_updates=True)