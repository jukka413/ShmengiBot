from aiogram import Bot, Dispatcher, types, executor
from aiogram.contrib.fsm_storage.memory import MemoryStorage
import os
import requests
import json
import datetime
import httplib2
import apiclient.discovery
from aiogram.dispatcher.handler import SkipHandler
from oauth2client.service_account import ServiceAccountCredentials

from config import TOKEN, SPREAD_SHEET_ID, G_KEY, SHEET_NAME, SHEET_FOR_COPY, MONTHS, USERS_ID
from states import States


bot = Bot(TOKEN)
dp = Dispatcher(bot, storage=MemoryStorage())


@dp.message_handler(content_types=types.ContentTypes.TEXT | types.ContentTypes.DOCUMENT)
async def check_user(message):
    if message.chat.id in USERS_ID:
        raise SkipHandler


@dp.message_handler(commands='start')
async def hello(message):
    await bot.send_message(message.chat.id, 'Привет, друг!')


@dp.message_handler(commands='copy')
async def copy(message):
    await bot.send_message(message.chat.id, 'Input month')
    await States.month.set()


@dp.message_handler(state=States.month)
async def input_month(message, state):
    answer = message.text
    copy_sheet(answer)
    await state.reset_state()


@dp.message_handler(content_types=types.ContentTypes.TEXT)
async def text_input(message):
    text = message.text
    if text not in MONTHS:
        name = text[:text.find(',')].strip()
        summ = text[text.find(',') + 1:].strip()
        empty = ""
        write_gs(name, empty, empty, summ, empty)
        await bot.send_message(message.chat.id, 'Запись произведена')
    else:
        await bot.send_message(message.chat.id, 'Введено название месяца')


@dp.message_handler(content_types=types.ContentTypes.DOCUMENT)
async def cheque_input(message):
    await bot.send_message(message.from_user.id, message.document.file_id)
    tuple_file_id = message.from_user.id, message.document.file_id
    file_id = tuple_file_id[1]
    url = f'https://api.telegram.org/bot{TOKEN}/getFile?file_id={file_id}'
    rq = requests.get(url).text
    path_start = rq.find('file_path":"') + len('file_path":"')
    path_finish = rq.find('"', path_start)
    path = rq[path_start:path_finish]
    file_url = f'https://api.telegram.org/file/bot{TOKEN}/{path}'
    js = requests.get(file_url)
    cheque = js.text
    info = json.loads(cheque)[0]
    # print(type(info))
    # print(info)

    info = info["ticket"]["document"]["receipt"]
    try:
        shop = info["user"]
    except KeyError:
        shop = "Unknown shop"
    items = info["items"]
    date = info["dateTime"]
    date = date[:date.find('T')]
    date = date[8:] + '.' + date[5:7] + '.' + date[:4]
    #print(date)
    print(items)
    for i in range(len(items)):
        name = items[i]["name"]
        price = items[i]["price"]/100
        quantity = items[i]["quantity"]
        summ = items[i]["sum"]/100
        write_gs(date, name, quantity, price, summ, shop)
    await bot.send_message(message.chat.id, 'Запись произведена')


def write_gs(date, name, count, price, summ, shop_name):
    service = connect_sheet()

    # date = datetime.date.today().strftime("%d.%m.%Y")
    date = date
    range = SHEET_NAME + "!A:F"
    list = [[date, name, count, price, summ, shop_name]]
    resource = {
        "majorDimension": "ROWS",
        "values": list
    }

    result = service.spreadsheets().values().append(
        spreadsheetId=SPREAD_SHEET_ID, range=range,
        valueInputOption="USER_ENTERED", body=resource).execute()
    # print('{0} cells appended.'.format(result.get('updates').get('updatedCells')))


def copy_sheet(month):
    service = connect_sheet()
    resource = {
        'destination_spreadsheet_id': SPREAD_SHEET_ID
    }
    request = service.spreadsheets().sheets().copyTo(spreadsheetId=SPREAD_SHEET_ID, sheetId=SHEET_FOR_COPY,
                                                     body=resource)
    response = request.execute()
    new_sheet_id = response["sheetId"]
    print(response, new_sheet_id)
    body = {
        'requests': [{'updateSheetProperties': {'properties': {'sheetId': new_sheet_id, 'title': month}, 'fields': 'title'}}]
    }
    print(requests)
    request = service.spreadsheets().batchUpdate(
        spreadsheetId=SPREAD_SHEET_ID,
        body=body).execute()

    with open('config.py', 'r', encoding='UTF8') as f:
        old_data = f.read()
        chunk = old_data[old_data.find('SHEET_NAME'):]
        old_sheet_name = chunk[:chunk.find("'", chunk.find("'")+1)+1]
    new_data = old_data.replace(old_sheet_name, "SHEET_NAME = '"+month+"'")
    with open('config.py', 'w', encoding='UTF8') as f:
        f.write(new_data)
    os.system('python "main.py"')


def connect_sheet():
    credentials = ServiceAccountCredentials.from_json_keyfile_name(G_KEY,
                                                                   ['https://www.googleapis.com/auth/spreadsheets',
                                                                    'https://www.googleapis.com/auth/drive'])
    http_auth = credentials.authorize(httplib2.Http())
    service = apiclient.discovery.build('sheets', 'v4', http=http_auth)
    return service


if __name__ == "__main__":
    executor.start_polling(dp)

