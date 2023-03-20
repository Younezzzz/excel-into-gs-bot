import gspread
import pandas as pd
import openpyxl
import time
import telebot
import os

bot = telebot.TeleBot('')

gc = gspread.service_account(filename='')
sh = gc.open_by_url('')
worksheet = sh.worksheet('')




sheet = 0
def add_value(row):
    values = []
    for cellObj in sheet[f'A{row}':f'AW{row}']:
        for cell in cellObj:
            values.append(cell.value)
    return values

@bot.message_handler(content_types=['text'])
def start(message):
    if message.text == '/start':
        bot.send_message(message.chat.id,'Начало')
@bot.message_handler(content_types=['document'])
def load_document(message):
    file_name = message.document.file_name
    file_info = bot.get_file(message.document.file_id)
    download = bot.download_file(file_info.file_path)
    src = file_name
    with open(src,'wb') as new_file:
        new_file.write(download)

    all_data = []
    work_book = openpyxl.open(file_name)
    global sheet
    sheet = work_book['Sheet1']

    with open(src,'wb') as new_file:
        new_file.write(download)
    values_list = worksheet.col_values(1)
    start_row = len(values_list) + 1
    for i in range(1, 301):
        all_data.append(add_value(i))
    if start_row > 1:
        all_data.remove(all_data[0])
    worksheet.update(f'A{start_row}:AW{start_row + 300}', all_data)
    os.remove(file_name)
    bot.send_message(message.chat.id,'Файл загружен')

bot.polling(none_stop=True)



