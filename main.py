import pandas as pd
import simplekml
import telebot
from telebot import types
from pathlib import Path
from tok import token
import datetime
import re


bot = telebot.TeleBot(token)

log = pd.read_excel('log.xlsx',usecols=['time', 'new messege','id user','total messeges','unique users'])
print (log)
def analitics (func: callable):
    total_messeges = 0
    users = set()
    total_users = 0
    def analitics_wrapper(messege):
        nonlocal total_messeges, total_users
        total_messeges += 1
        if messege.chat.id not in users:
            users.add(messege.chat.id)
            total_users += 1
        log.loc[len(log.index)] = [datetime.datetime.now(), messege.text, messege.chat.id, total_messeges,total_users]
        print (log)
        log.to_excel('log.xlsx')
        return func(messege)
    return analitics_wrapper

@bot.message_handler(commands=['start'])
@analitics
def start(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=1)
    btn_dots = types.KeyboardButton("Точки")
    btn_poly = types.KeyboardButton("Полигон")
    btn_line = types.KeyboardButton("Линия")
    markup.add(btn_dots, btn_poly, btn_line)
    bot.send_message(message.chat.id,
                     text="Привет, {0.first_name}! Выбери тип объекта для нанесения на карту".format(message.from_user),
                     reply_markup=markup)

tochki_list = []
plygon_list = []
line_list = []
@bot.message_handler()
@analitics
def get_user_text(message):
    if message.text == "Точки":
        tochki = 1
        tochki_list.clear()
        tochki_list.append(tochki)
        plygon_list.clear()
        line_list.clear()
        bot.send_document(message.chat.id, open(r'file/Форма заполнения для точек.xlsx', 'rb'))
        bot.send_message(message.chat.id,
                         text="Внеси в этот файл точки и отправь мне. Или скопируй таблицу с точками прямо в сообщение. Формат - Имя точки, X, Y. (Координаты должны быть десятичные с точкой. Разделитель - пробел, табуляция или запятая)".format(message.from_user))

    elif message.text == "Полигон":
        polygon = 1
        plygon_list.clear()
        plygon_list.append(polygon)
        tochki_list.clear()
        line_list.clear()
        bot.send_document(message.chat.id, open(r'file/Форма заполнения для полигона.xlsx', 'rb'))
        bot.send_message(message.chat.id,
                         text="Внеси в этот файл точки и отправь мне. Или скопируй таблицу с точками прямо в сообщение. Формат - Имя точки, X, Y. (Координаты должны быть десятичные с точкой. Разделитель - пробел, табуляция или запятая)".format(message.from_user))
    elif message.text == "Линия":
        line = 1
        line_list.clear()
        line_list.append(line)
        plygon_list.clear()
        tochki_list.clear()
        bot.send_document(message.chat.id, open(r'file/Форма заполнения для линии.xlsx', 'rb'))
        bot.send_message(message.chat.id,
                         text="Внеси в этот файл точки и отправь мне. Или скопируй таблицу с точками прямо в сообщение. Формат - Имя точки, X, Y. (Координаты должны быть десятичные с точкой. Разделитель - пробел, табуляция или запятая)".format(message.from_user))
    elif message.text == "стат1":
        bot.send_document(message.chat.id, open(r'log.xlsx', 'rb'))




    else:
        print (tochki_list,plygon_list)
        data1 = message.text
        data1 = data1.replace(" ", "")
        print (re.search(r"\d", data1))
        if not re.search(r"\d", data1):
            if re.match(r".*привет", data1, re.I):
                bot.reply_to(message, 'привет)')
            else:
                bot.reply_to(message, 'Используй цифры')
        else:
            try:
                data1 = message.text
                data = re.split(" |,|\n|\t", data1)
                data = list(filter(len, data))
                K = 3
                data = [data[i:i + K] for i in range(0, len(data), K)]
                data = pd.DataFrame(data,columns=['№ точки','X', 'Y'])
                print (data)

                df0 = data['№ точки'].to_list()
                df1 = data['X'].to_list()
                df2 = data['Y'].to_list()
                dff = data[['Y', 'X']]
                coord = dff.values.tolist()
                kml = simplekml.Kml()

                if sum(tochki_list) == 1 and sum(plygon_list) == 0 and sum(line_list) == 0:
                    coord = list(zip(df0, df1, df2))
                    for city, X, Y in coord:
                        pnt = kml.newpoint(name=city)
                        coord = [(Y, X)]
                        pnt.coords = coord
                    kml.save(r"file/Точки.kml")
                    bot.send_document(message.chat.id, open(r'file/Точки.kml', 'rb'))
                elif sum(tochki_list) == 0 and sum(plygon_list) == 1 and sum(line_list) == 0:
                    pol = kml.newpolygon(name='Полигон')
                    pol.outerboundaryis = coord
                    kml.save(r"file/Полигон.kml")
                    bot.send_document(message.chat.id, open(r'file/Полигон.kml', 'rb'))
                elif sum(line_list) == 1 and sum(plygon_list) == 0 and sum(tochki_list) == 0:
                    ls = kml.newlinestring(name='Линия')
                    ls.coords = coord
                    ls.extrude = 1
                    ls.altitudemode = simplekml.AltitudeMode.relativetoground
                    kml.save(r"file/Линия.kml")
                    bot.send_document(message.chat.id, open(r'file/Линия.kml', 'rb'))

                elif sum(tochki_list) == 0 and sum(plygon_list) == 0 and sum(line_list) == 0:
                    bot.reply_to(message, 'Выбери вид KML (напиши /start)')
                else:
                    bot.reply_to(message, 'Нехватает данных')
            except:
                data1 = message.text
                data1.replace(' ', '')
                if re.match (r".привет", data1):
                    bot.reply_to(message, 'привет)')
                elif not data1.isnumeric():
                    bot.reply_to(message, 'Нехватает данных (возможно нет данных имени точки)')
                else:
                    bot.reply_to(message, 'Нехватает данных')



@bot.message_handler(content_types=['document'])
@analitics
def handle_docs(message):
    try:
        chat_id = message.chat.id

        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        print (Path(message.document.file_name).stem, Path(message.document.file_name).suffix)
        print (Path(message.document.file_name).stem == 'Форма заполнения для полигона' and Path(message.document.file_name).suffix == '.xlsx')
        print (Path(message.document.file_name).stem == 'Форма заполнения для точек' and Path(message.document.file_name).suffix == '.xlsx')
        df = pd.read_excel(downloaded_file)
        print(df)
        df0 = df['№ точки'].to_list()
        df1 = df['X'].to_list()
        df2 = df['Y'].to_list()

        dff = df[['Y','X']]
        # dff.rename(columns={'X': 'Y', 'Y': 'X'}, inplace=True)
        coord = dff.values.tolist()
        print (coord)

        kml = simplekml.Kml()

        if Path(message.document.file_name).stem == 'Форма заполнения для точек' and Path(message.document.file_name).suffix == '.xlsx':
            coord = list(zip(df0, df1, df2))
            for city, X, Y in coord:
                pnt = kml.newpoint(name=city)
                coord = [(Y, X)]
                pnt.coords = coord
            kml.save(r"file/Точки.kml")
            bot.send_document(message.chat.id, open(r'file/Точки.kml', 'rb'))

        elif Path(message.document.file_name).stem == 'Форма заполнения для полигона' and Path(message.document.file_name).suffix == '.xlsx':
            pol = kml.newpolygon(name='Полигон')
            pol.outerboundaryis = coord
            kml.save(r"file/Полигон.kml")
            bot.send_document(message.chat.id, open(r'file/Полигон.kml', 'rb'))

        elif Path(message.document.file_name).stem == 'Форма заполнения для линии' and Path(message.document.file_name).suffix == '.xlsx':
            ls = kml.newlinestring(name='Линия')
            ls.coords = coord
            ls.extrude = 1
            ls.altitudemode = simplekml.AltitudeMode.relativetoground
            kml.save(r"file/Линия.kml")
            bot.send_document(message.chat.id, open(r'file/Линия.kml', 'rb'))

        else:
            bot.reply_to(message, 'Видимо это не тот файл.. необходимо использовать файл, который я отправил тебе')


    except Exception as e:
        print(e)
        bot.reply_to(message, 'Видимо это не тот файл.. необходимо использовать файл, который я отправил тебе')

bot.polling(none_stop=True)