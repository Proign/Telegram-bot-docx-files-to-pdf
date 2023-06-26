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

# The time the file was last sent
last_file_sent_time = {}

@dp.message_handler(commands=['start'])
async def send_welcome(message: types.Message):
   inline_keyboard = types.InlineKeyboardMarkup()
   help_button = types.InlineKeyboardButton("Help", callback_data="Help")
   inline_keyboard.add(help_button)

   await message.answer('Greetings! I\'m a bot that converts docx to pdf - download the file to get started', reply_markup = inline_keyboard)

# Button click handler
@dp.callback_query_handler(lambda callback_query: True)
async def button_click_handler(callback_query: types.CallbackQuery):
   if callback_query.data == "Help":
      await bot.send_message(callback_query.from_user.id,'This bot converts docx format files to pdf. \nto get started, upload the file using the "Attach" button or by dragging the file into the message input field.')

@dp.message_handler(commands=['help'])
async def send_welcome(message: types.Message):
   await message.reply('This bot converts docx format files to pdf. \nto get started, upload the file using the "Attach" button or by dragging the file into the message input field.')

@dp.message_handler()
async def message_reply(message: types.Message):
   # Emoji detection in text
   if emoji.is_emoji(message.text):
      # If emojis are found, we send them back
      await message.answer(message.text)
   else:
      await message.reply('To get started, send a file with the docx extension')

@dp.message_handler(content_types=[types.ContentType.DOCUMENT, types.ContentType.PHOTO])
async def handle_files(message: types.Message):

   user_id = message.from_user.id
   current_time = time.time()
   
   # Check how much time has passed since the last file was sent
   if user_id in last_file_sent_time and (current_time - last_file_sent_time[user_id]) < 10:
      await message.answer('Please wait 10 seconds before sending the next file')
      return

   # Check if the message is a document
   if message.document:
      # Updating the time of the last file submission
      last_file_sent_time[user_id] = current_time

      file_id = message.document.file_id
      file_name = message.document.file_name.rsplit('.', 1)[0]
      file_info = await bot.get_file(file_id)
      file_path = file_info.file_path

      # Getting the file extension
      file_extension = file_path.split('.')[-1]

      # Check that the file extension matches the valid values
      if file_extension.lower() == 'docx':
            # Getting information about the file size
            file_size = file_info.file_size
            # Checking the file size (up to 100 MB)
            if file_size <= 100 * 1024 * 1024:
               # Saving the file on the server
               dir = 'Source\\'
               await bot.download_file(file_path, (dir + f'{user_id}.{file_extension}'))
               await message.answer('The file has been uploaded successfully! The process will take a few seconds')

               # Convert DOCX to PDF
               docx_file = dir + f'{user_id}.{file_extension}'
               pdf_file = dir + f'{file_name}.pdf'
               convert(docx_file, pdf_file)

               # Sending a file to a user
               with open(pdf_file, 'rb') as file:
                  await bot.send_document(message.chat.id, file)

               if not Logging:
                  for f in os.listdir(dir):
                     os.remove(os.path.join(dir, f))
            else:
               await message.answer('The file size exceeds 100 MB. Please send a file up to 100 MB in size')
      else:
         await message.answer(f'Failed to process file with {file_extension} extension. Please send the file with the docx extension')

   # Check if the message is a photo
   elif message.photo:
      # Updating the time of the last file submission
      last_file_sent_time[user_id] = current_time

      photo = message.photo[-1]
      file_id = photo.file_id
      file_info = await bot.get_file(file_id)
      file_path = file_info.file_path

      # Getting the file extension
      file_extension = file_path.split('.')[-1]
      await message.answer(f'Failed to process file with {file_extension} extension. Please send the file with the docx extension')
 
if __name__ == '__main__':
   executor.start_polling(dp, skip_updates=True)