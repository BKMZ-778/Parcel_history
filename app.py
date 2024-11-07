import ssl
from calendar import monthrange

import xlsxwriter
from flask import Flask, jsonify, request, render_template, redirect, url_for
from flask import abort
import pandas as pd
from sqlalchemy.orm import sessionmaker, scoped_session
from sqlalchemy import create_engine
from sqlalchemy.ext.declarative import declarative_base
import sqlite3 as sl
from flask_jwt_extended import JWTManager, jwt_required, get_jwt_identity
from config import Config
from apispec.ext.marshmallow import MarshmallowPlugin
from apispec import APISpec
from flask_apispec.extension import FlaskApiSpec
from schemas import VideoSchema, UserSchema, AuthSchema
from flask_apispec import use_kwargs, marshal_with
import zipfile
from zipfile import ZipFile
import logging
from flask import Response
import sqlalchemy as db
import pytz
import datetime
import os
from imbox import Imbox
import pyexcel as p
import schedule
import telebot
import traceback
import time
import atexit
from apscheduler.schedulers.background import BackgroundScheduler
import rarfile
import os.path
import glob
import requests
from flask_restful import Resource,  Api
from flask_jwt_extended import JWTManager
from flask_jwt_extended import create_access_token, jwt_required
#from flask_cors import CORS
from flask import make_response, send_file
from flask_jwt_extended import (
    JWTManager, jwt_required, create_access_token,
    create_refresh_token,
    get_jwt_identity, set_access_cookies,
    set_refresh_cookies, unset_jwt_cookies
)
import base64
from base64 import b64encode
import hashlib
import json
import smtplib

bot = telebot.TeleBot('5555513345:AAFzGfbHd4rUzLHh2m4q5kWEFtp7IUx_UNU')
rarfile.UNRAR_TOOL = r'C:/Program Files/WinRAR/UnRAR.exe'
pd.set_option('display.max_columns', None)

now = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
delta = datetime.timedelta(hours=-10, minutes=0)
now_date = datetime.datetime.now() + delta
app = Flask(__name__)
app.config.from_object(Config)
app.config['BASE_URL'] = 'http://127.0.0.1:5000'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
app.config['UPLOAD_EXTENSIONS'] = ['.xls', '.xlsx']
app.config['JWT_CSRF_CHECK_FORM'] = True
app.config['JWT_TOKEN_LOCATION'] = ['cookies']
client = app.test_client()
engine = create_engine('sqlite:///db.sqlite')
session = scoped_session(sessionmaker(autocommit=False, autoflush=False, bind=engine))

Base = declarative_base()
Base.query = session.query_property()

jwt = JWTManager(app)
app.config['JWT_SECRET_KEY'] = 'my_cool_secret'
# CORS(app)
api = Api(app)


OZON_keys = {
            'OZON-AIR-260': {
              "clientId": "78d96cd9-909c-4e17-ac73-1b6669cbbd43",
              "clientSecret": "fyfJmgyftyYB",
              "providerId": 260,
            },
            'OZON-LAND-862': {
              "clientId": "58d65bb4-f7c5-4605-885f-6a9c54e4df7c",
              "clientSecret": "tMgBeETAXMJy",
              "providerId": 862,
            },
            'OZON-AIR-963': {
              "clientId": "39cbcb39-ccbe-415c-8ab7-bfb9ff41c9d7",
              "clientSecret": "OuYihXlWHZaQ",
              "providerId": 963,
            },
            'OZON-LAND-964': {
              "clientId": "a31f2584-94b7-48ac-bf9f-72a526fc4947",
              "clientSecret": "DeEbfGsSdnRd",
              "providerId": 964,
            },
            'OZON-LAND-989': {
              "clientId": "bbfcd88a-9b01-4f59-95b8-c56822e8b7ea",
              "clientSecret": "tNQYXddVwWpM",
              "providerId": 989,
            },
            'OZON-AIR-995': {
              "clientId": "1cc37816-9c1a-466c-bf61-8267df5d6342",
              "clientSecret": "ZyMbKAFIxyhy",
              "providerId": 995,
            },
            'OZON-LAND-1045': {
              "clientId": "13be002c-c54d-447c-9a25-396cabc6e1b1",
              "clientSecret": "XKAMTEKMQftH",
              "providerId": 1045,
            },
            'FBP-AIR-1102': {
              "clientId": "04904a46-4c9f-4f7e-b2e9-019917d491de",
              "clientSecret": "YyBnRTMQaahK",
              "providerId": 1102,
            },
            '7D-AIR-987': {
              "clientId": "ffa8c005-4f91-4a47-b2cf-6db0fe9ee97c",
              "clientSecret": "emtyBaeJtgKX",
              "providerId": 987,
            },
            '7D-LAND-1068': {
              "clientId": "10d0ceaa-267b-4e46-a464-be6074c9c6fd",
              "clientSecret": "GTmnYatEXhJE",
              "providerId": 1068,
            },
            'FBP-LAND-1108': {
              "clientId": "a129b657-2ba4-4ea6-9fa5-704a9b30937c",
              "clientSecret": "gLhAeXLahtYL",
              "providerId": 1108,
            },
            'FBP-AIR-1108': {
                "clientId": "a129b657-2ba4-4ea6-9fa5-704a9b30937c",
                "clientSecret": "gLhAeXLahtYL",
                "providerId": 1108,
            },
            '7D-LAND-1110': {
              "clientId": "2d3bf6b8-9f26-493a-a4d6-48e853405cd6",
              "clientSecret": "MYrevesoOyRE",
              "providerId": 1110,
            },
            '7D-AIR-1111': {
              "clientId": "86b43b9f-3dfb-41b5-96b6-19792f08ff90",
              "clientSecret": "tntwpAModoSt",
              "providerId": 1111,
            },
            'OZON-AIR-1160': {
              "clientId": "a813ba8f-d607-4e68-83f5-a2ca76e75f88",
              "clientSecret": "uhLvfyDyuzDY",
              "providerId": 1160,
            },
            'OZON-AIR-1163': {
              "clientId": "7a68022c-5701-4a95-a210-3451beb85cf2",
              "clientSecret": "YyZFgSKfXaQV",
              "providerId": 1163,
            },
            'OZON-AIR-1164': {
              "clientId": "c9733e62-1df2-44a6-8540-4bab89ee6876",
              "clientSecret": "ZZJwkxWVDVpW",
              "providerId": 1164,
            },
            'OZON-LAND-1165': {
              "clientId": "fa4f1d2d-4522-4b85-bc52-da1b1b55b9f7",
              "clientSecret": "TEKuTubKYmaQ",
              "providerId": 1165,
            },
            'OZON-AIR-1166': {
              "clientId": "1e807247-7aa2-41d6-873e-c186b8386cc9",
              "clientSecret": "GFJvMyPQZauR",
              "providerId": 1166,
            },
            'OZON-AIR-1167': {
              "clientId": "1c1cee8a-46c3-4d04-b432-bbcf8fb4b938",
              "clientSecret": "ymUJGMwmvFcy",
              "providerId": 1167,
            },
            'OZON-LAND-1168': {
              "clientId": "062e6800-e66c-4fa2-90f5-8dcbcdde6992",
              "clientSecret": "LLthZcDtnFRZ",
              "providerId": 1168,
            },
            'OZON-AIR-1169': {
              "clientId": "89e7ad9e-ddd5-4e5d-9da1-cfb84e2f79e9",
              "clientSecret": "vWYNbWjzQdUJ",
              "providerId": 1169,
            },
            'OZON-AIR-1170': {
              "clientId": "ede0d39d-8b39-49a7-b581-91d1f9f3d4c0",
              "clientSecret": "mThSncRxCFFb",
              "providerId": 1170,
            },
            'OZON-LAND-1171': {
              "clientId": "bec77a1f-93bd-4895-9794-006b7e2a0783",
              "clientSecret": "QmxjpMfxJkFf",
              "providerId": 1171,
            },
            'OZON-AIR-1172': {
              "clientId": "1571707c-faf3-4704-9d45-37102775c088",
              "clientSecret": "RYCurLxxvukg",
              "providerId": 1172,
            },
            'OZON-AIR-1173': {
              "clientId": "33fd83a0-8db3-4e7f-b130-5b16865fad56",
              "clientSecret": "UUArrJCmCYYC",
              "providerId": 1173,
            },
            'OZON-LAND-1174': {
              "clientId": "14356255-28e8-4ea1-bd63-a60e292ea90a",
              "clientSecret": "zEednHwsQehX",
              "providerId": 1174,
            },
            'OZON-AIR-1175': {
              "clientId": "74bd8449-a3cb-4aa5-9725-5047305b7c07",
              "clientSecret": "UycNZFQQkTRG",
              "providerId": 1175,
            },
            'OZON-AIR-1176': {
              "clientId": "1c1c74c8-dd53-4597-8158-ffc2c3948815",
              "clientSecret": "QPxLFkbNyPTg",
              "providerId": 1176,
            },
            'OZON-LAND-1177': {
              "clientId": "063a27c7-2ff2-4302-ac65-d27da5635b67",
              "clientSecret": "TTGnLHyuxghQ",
              "providerId": 1177,
            },
            'FBP-LAND-1178': {
              "clientId": "4ce9623d-cfc7-4cf4-9b1f-3dacbef4f69c",
              "clientSecret": "DRYfKNhbuTXK",
              "providerId": 1178,
            },
            '7D-AIR-1209': {
              "clientId": "500477ef-1e92-43e8-972f-4d0ce690774b",
              "clientSecret": "DyxMsxTpAcgk",
              "providerId": 1209,
            }

            }

client_id_agreg = '78d96cd9-909c-4e17-ac73-1b6669cbbd43'
client_secret_agreg = 'fyfJmgyftyYB'

class UserLogin(Resource):
    def post(self):
        username = request.get_json()['username']
        password = request.get_json()['password']
        if username == 'admin' and password == 'habr':
            access_token = create_access_token(identity={
                'role': 'admin',
            }, expires_delta=False)
            result = {'token': access_token}
            return result
        return {'error': 'Invalid username and password'}


class ProtectArea(Resource):
    @jwt_required
    def get(self):
        return {'answer': 42}


api.add_resource(UserLogin, '/api/login/')
api.add_resource(ProtectArea, '/api/protect-area/')

docs = FlaskApiSpec()
docs.init_app(app)
app.config.update({
    'APICPEC_SPEC': APISpec(
        title='videoblog',
        version='v1',
        openapi_version='2.0',
        plugins=[MarshmallowPlugin()]
    ),
    'APISPEC_SWAGGER_URL': '/swagger/'
})

from models import *
Base.metadata.create_all(bind=engine)


def setup_logger(name, log_file, level=logging.INFO):
    logging.basicConfig(format=u'%(levelname)-8s [%(asctime)s] %(message)s')  # filename=u'mylog.log'
    handler = logging.FileHandler(log_file)
    logger = logging.getLogger(name)
    logger.setLevel(level)
    logger.addHandler(handler)
    return logger

logger = setup_logger('logger', 'mylog.log')


logger_API_insert = setup_logger('logger_API_insert', 'API_insert.log')


logger_API_get_info = setup_logger('logger_API_get_info', 'API_info.log')

logger_API_chunks = setup_logger('logger_API_chunks', 'logger_API_chunks.log')

logger_customs_pay = setup_logger('logger_customs_pay', 'logger_customs_pay.log')
logger_customs_paya_all = setup_logger('customs_paya_all', 'customs_paya_all.log')

logger_tax_documents = setup_logger('logger_tax_documents', 'logger_tax_documents.log')

logger_pay_errors = setup_logger('logger_pay_errors', 'logger_pay_errors.log')

con_gps = sl.connect("GPS.db")
with con_gps:
    data = con_gps.execute("select count(*) from sqlite_master where type='table' and name='gps_parcels'")
    for row in data:
        # если таких таблиц нет
        if row[0] == 0:
            # создаём таблицу
            con_gps.execute("""
                                                                            CREATE TABLE gps_parcels (
                                                                            ID INTEGER PRIMARY KEY AUTOINCREMENT,
                                                                            parcel_numb VARCHAR(25) NOT NULL UNIQUE ON CONFLICT REPLACE,
                                                                            gps_numb VARCHAR(25) NOT NULL
                                                                            );
                                                                            """)


con_pay = sl.connect("Pay.db")
with con_pay:
    data = con_pay.execute("select count(*) from sqlite_master where type='table' and name='pay_customs'")
    for row in data:
        # если таких таблиц нет
        if row[0] == 0:
            # создаём таблицу
            con_pay.execute("""
                                                                            CREATE TABLE pay_customs (
                                                                            ID INTEGER PRIMARY KEY AUTOINCREMENT,
                                                                            PostingNumber VARCHAR(25) NOT NULL,
                                                                            TrackingNumber VARCHAR(25) NOT NULL,
                                                                            TaxPayment FLOAT,
                                                                            CustomsDuty FLOAT,
                                                                            Total FLOAT,
                                                                            InvoiceNumber VARCHAR(50),
                                                                            RegisterNumber VARCHAR(25),
                                                                            Currency VARCHAR(3),
                                                                            DateOfPayment,
                                                                            Provider,
                                                                            send_time,
                                                                            json_info,
                                                                            ozone_response_status_code,
                                                                            ozone_response_text
                                                                            );
                                                                            """)

with con_pay:
    data = con_pay.execute("select count(*) from sqlite_master where type='table' and name='cel_tokens'")
    for row in data:
        # если таких таблиц нет
        if row[0] == 0:
            # создаём таблицу
            con_pay.execute("""
                                                                            CREATE TABLE cel_tokens (
                                                                            ID INTEGER PRIMARY KEY AUTOINCREMENT,
                                                                            provider VARCHAR,
                                                                            token VARCHAR,
                                                                            updateTime VARCHAR

                                                                            );
                                                                            """)

def send_email(body_text, subject):
    username = "cel-python-automatization@yandex.ru"
    password = "vmpqeopkfptvejiz"
    context = ssl.create_default_context()
    from_addr = 'cel-python-automatization@yandex.ru'

    to_addr = ["transpriemka@mail.ru"]  # must be a list

    # Prepare actual message
    message = "\r\n".join((
        "From: %s" % from_addr,
        "To: %s" % to_addr,
        "Subject: %s" % subject,
        "",
        body_text
    ))

    with smtplib.SMTP_SSL("smtp.yandex.com", context=context) as server:
        server.login(username, password)
        server.sendmail(username, ["transpriemka@mail.ru"], message)
        server.quit()
        print('ok')

def api_track718_add_track(gps_numb):
    cel_api_key = "e0fca820-c3dc-11ee-b960-bdfb353c94dc"

    url = "https://apigetway.track718.net/v2/tracks"
    headers = {"Content-Type": "application/json",
    "Track718-API-Key": f"{cel_api_key}"}

    params = [{"trackNum": gps_numb, "code": "gps-truck"}]
    respons = requests.post(url=url, headers=headers, json=params)


    print(respons.status_code)
    print(respons)
    print(respons.json())


def transpriemka_scan():
    now = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
    now_date = datetime.datetime.now() + delta
    host = "imap.yandex.ru"
    username = "cel-python-automatization@yandex.ru"
    password = "vmpqeopkfptvejiz"
    main_folder = "/РЕЕСТРЫ для статусов"
    download_folder = main_folder
    logging.info(download_folder)
    if not os.path.isdir(download_folder):
        os.makedirs(download_folder, exist_ok=True)
    start_date = datetime.date(2023, 3, 1)
    end_date = datetime.date(2023, 2, 24)
    mail = Imbox(host, username=username, password=password, ssl=True, ssl_context=None, starttls=False)
    quont = 0
    try:
        messages_all = mail.messages(date__on=datetime.date(now_date.year, now_date.month, now_date.day))  # date__on=datetime.date(now_date.year, now_date.month, now_date.day-1)
        logging.info(f'{now_date.year}, {now_date.month}, {now_date.day}-2')
        try:
            logging.info(len(messages_all))
            for (uid, message) in messages_all:
                try:
                    message_subject = message.subject
                    #print(message.date)
                    #rcv_tm1 = parser.parse(message.date)
                    #print(rcv_tm1)
                    date_of_message = message.date
                    rcv_tm = datetime.datetime.strptime(date_of_message[0:31], '%a, %d %b %Y %H:%M:%S %z')
                    # print((rcv_tm))
                    # print(start_date)
                    try:
                        mail.mark_seen(uid)  # optional, mark message as read
                        for message in message.attachments:
                            try:
                                substring_list = ['LD', 'JD', 'CNC', 'форма', 'Форма', 'клиент', 'CHZ', 'илья', '珲春仓到', 'OZON', 'ETS', 'WB']
                                substring_party_list = ['CEL', 'CL', 'OZON', 'SUI']
                                att_fn = message.get('filename')
                                print(att_fn)
                                file = message.get('content')
                                for substring in substring_list:
                                    if substring in att_fn:
                                        try:
                                            download_path = f"{download_folder}/{att_fn}"
                                            with open(download_path, "wb") as fp:
                                                fp.write(message.get('content').read())
                                                logging.info(download_path)
                                                logging.info(rcv_tm)
                                                time.sleep(3)
                                            try:
                                                df = pd.read_excel(file, sheet_name=0, engine='openpyxl')
                                            except:
                                                dest_file_name = f'{download_path}.xlsx'
                                                p.save_book_as(file_name=download_path,
                                                            dest_file_name=dest_file_name)
                                                time.sleep(3)
                                                df = pd.read_excel(f'{dest_file_name}', sheet_name=0, engine='openpyxl')
                                            print(df)
                                            for col in df.columns:
                                                if 'Номернакладной' in col or 'Номер накладной' in col or 'Номер накладной' in col:
                                                    df = df[col].to_frame()
                                                    df = df.drop_duplicates(keep='first')
                                                    df = df.rename(columns={df.columns[0]: 'parcel_numb'})
                                                    df['Event'] = 'Сформировано к отгрузке со склада'
                                                    df['Event_comment'] = 'Склад отправителя в Китае'
                                                    df['Event_date'] = rcv_tm
                                                    con = sl.connect('CEL.db')
                                                    with con:
                                                        df.to_sql('events2', con=con, if_exists='append', index=False)
                                                    con.commit()
                                                    quont += 1
                                                    break

                                        except Exception as e:
                                            message = f'Реестр {att_fn} ({substring}) от {rcv_tm}\n\n не был заружен в базу данных. Добавьте статусы вручную. Описание ошибки:\n\n{e}'
                                            #bot = telebot.TeleBot('5555513345:AAFzGfbHd4rUzLHh2m4q5kWEFtp7IUx_UNU')
                                            #bot.send_message(1634121947, message)  # 1285743017
                                            #bot.send_message(422263274, message)  # 1285743017
                                            #bot.send_message(1285743017, message)  # 1285743017
                                for substring_party in substring_party_list:
                                    if substring_party in att_fn:
                                        try:
                                            download_path = f"{download_folder}/{att_fn}"
                                            with open(download_path, "wb") as fp:
                                                fp.write(message.get('content').read())
                                                time.sleep(3)
                                            try:
                                                df = pd.read_excel(file, sheet_name=0, engine='openpyxl')
                                            except:
                                                dest_file_name = f'{download_path}.xlsx'
                                                p.save_book_as(file_name=download_path,
                                                            dest_file_name=dest_file_name)
                                                df = pd.read_excel(f'{dest_file_name}', sheet_name=0, engine='openpyxl')
                                            columns_names = 'Order number'
                                            for col in df.columns:
                                                if columns_names in col:
                                                    df = df[col].to_frame()
                                                    df = df.drop_duplicates(keep='first')
                                                    df = df.rename(columns={df.columns[0]: 'parcel_numb'})
                                                    df['Event'] = 'Сформированно в партию к отгрузке в РФ'
                                                    df['Event_comment'] = 'Хунчунь (КНР)'
                                                    df['Event_date'] = rcv_tm
                                                    con = sl.connect('CEL.db')
                                                    with con:
                                                        df.to_sql('events2', con=con, if_exists='append', index=False)
                                                    con.commit()
                                                    quont += 1
                                                    break
                                                else:
                                                    pass
                                        except Exception as e:
                                            att_fn = message.get('filename')
                                            message = f'Партия {att_fn} от {rcv_tm}\n\n не была заружен в базу данных. Добавьте статусы вручную. Описание ошибки:\n\n{e}\n\n{traceback.print_exc()}'
                                            bot = telebot.TeleBot(
                                                '5555513345:AAFzGfbHd4rUzLHh2m4q5kWEFtp7IUx_UNU')
                                            #bot.send_message(1634121947, message)  # 1285743017
                                            #bot.send_message(422263274, message)  # 1285743017
                                            #bot.send_message(1285743017, message)  # 1285743017
                                            #logger.warning(f'error {e}')
                                        break
                                if 'ocean' in att_fn.lower() or 'ocean' in message_subject.lower():
                                    try:
                                        download_path = f"{download_folder}/{att_fn}"
                                        with open(download_path, "wb") as fp:
                                            fp.write(message.get('content').read())
                                            time.sleep(3)
                                        try:
                                            df = pd.read_excel(file, sheet_name=0, engine='openpyxl')
                                        except:
                                            dest_file_name = f'{download_path}.xlsx'
                                            p.save_book_as(file_name=download_path,
                                                           dest_file_name=dest_file_name)
                                            df = pd.read_excel(f'{dest_file_name}', sheet_name=0, engine='openpyxl')
                                        columns_names = 'Номер-посылки'
                                        for col in df.columns:
                                            if columns_names in col:
                                                df = df[col].to_frame()
                                                df = df.drop_duplicates(keep='first')
                                                df = df.rename(columns={df.columns[0]: 'parcel_numb'})
                                                df['Event'] = 'Сформированно в партию к отгрузке в РФ'
                                                df['Event_comment'] = 'Тояма (Япония)'
                                                df['Event_date'] = rcv_tm
                                                print(df)
                                                con = sl.connect('CEL.db')
                                                with con:
                                                    df.to_sql('events2', con=con, if_exists='append', index=False)
                                                con.commit()
                                                break
                                            else:
                                                pass
                                    except Exception as e:
                                            message = f'Реестр {att_fn} от {rcv_tm}\n\n: {e}'
                            except Exception as e:
                                att_fn = message.get('filename')
                                message = f'Реестр {att_fn} от {rcv_tm}\n\n не был заружен в базу данных. Добавьте статусы вручную. Описание ошибки:\n\n{e}\n\n{traceback.print_exc()}'
                                #bot = telebot.TeleBot('5555513345:AAFzGfbHd4rUzLHh2m4q5kWEFtp7IUx_UNU')
                                #bot.send_message(1634121947, message)  # 1285743017
                                #bot.send_message(422263274, message)  # 1285743017
                                #bot.send_message(1285743017, message)  # 1285743017
                                #logger.warning(f'error {e}')
                    except Exception as e:
                        logger.info(f'error {traceback.print_exc()}')
                except Exception as e:
                    logger.info(f'error {traceback.print_exc()}')
            else:
                """print('Well done')
                message = f'Реестры и партии от {start_date}\n\n заружены в базу данных. \n\nСтатусы проставлены.'
                bot = telebot.TeleBot('5555513345:AAFzGfbHd4rUzLHh2m4q5kWEFtp7IUx_UNU')
                bot.send_message(1634121947, message)  # 1285743017
                bot.send_message(422263274, message)  # 1285743017
                bot.send_message(1285743017, message)  # 1285743017"""
                pass
        except Exception as e:
            logger.info(f'error {traceback.print_exc()}')
    except Exception as e:
        logger.info(f'error {traceback.print_exc()}')
    #message = f'Реестры и партии за {now_date} в кол-ве {quont} заружены в базу данных'
    #bot = telebot.TeleBot('5555513345:AAFzGfbHd4rUzLHh2m4q5kWEFtp7IUx_UNU')
    #bot.send_message(1634121947, message)  # 1285743017
    #bot.send_message(422263274, message)  # 1285743017
    #bot.send_message(1285743017, message)  # 1285743017


def logistick_scan():
    now = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
    now_date = datetime.datetime.now()
    host = "imap.yandex.ru"
    username = "logistick.dv@yandex.ru"
    password = "rwlefgatbfpewlmt"
    main_folder = "C:/Users/User/РЕЕСТРЫ для статусов"
    download_folder = main_folder
    logging.info(download_folder)
    if not os.path.isdir(download_folder):
        os.makedirs(download_folder, exist_ok=True)
    start_date = datetime.date(2023, 2, 20)
    end_date = datetime.date(2023, 2, 24)

    mail = Imbox(host, username=username, password=password, ssl=True, ssl_context=None, starttls=False)
    quont = 0
    try:
        messages_all = mail.messages(date__on=datetime.date(now_date.year, 3, 31), folder='&BB4EQgQ,BEAEMAQyBDsENQQ9BD0ESwQ1-')
        try:
            print(len(messages_all))
            for (uid, message) in messages_all:
                try:
                    # print(message.date)
                    # rcv_tm1 = parser.parse(message.date)
                    # print(rcv_tm1)
                    date_of_message = message.date
                    rcv_tm = datetime.datetime.strptime(date_of_message[0:31], '%a, %d %b %Y %H:%M:%S %z')
                    # print(start_date)
                    message_subject = message.subject
                    substring_subject_list = ['заявка', 'обеспечение']
                    for substring_subject in substring_subject_list:
                        if substring_subject in message_subject.lower():
                            logging.info('заявка')
                            try:
                                mail.mark_seen(uid)  # optional, mark message as read
                                for message in message.attachments:
                                    substring_party_list = ['CL', 'CEL', 'OL-', 'OZON', 'SUI']
                                    pac_string_trigger = 'PAC'
                                    att_fn = message.get('filename')
                                    print(att_fn)
                                    file = message.get('content')
                                    for substring_party in substring_party_list:
                                        rar_string_trigger = 'rar'
                                        if substring_party in att_fn and pac_string_trigger in att_fn:
                                            print(substring_party, att_fn)
                                            try:
                                                download_path = f"{download_folder}/{att_fn}"
                                                with open(download_path, "wb") as fp:
                                                    fp.write(message.get('content').read())
                                                    logging.info(download_path)
                                                    logging.info(rcv_tm)
                                                    time.sleep(2)
                                                df = pd.read_excel(file, sheet_name=0, engine='openpyxl', skiprows=9)
                                                print(df)
                                                df = df[df.columns[3]].to_frame()
                                                df = df.drop_duplicates(keep='first')
                                                df = df.rename(
                                                    columns={df.columns[0]: 'parcel_numb'})
                                                df['Event'] = 'Таможенный транзит'
                                                if substring_party == substring_party_list[2]:
                                                    print(substring_party)
                                                    df['Event_comment'] = 'порт Владивосток'
                                                    df['Event_date'] = rcv_tm
                                                    print(df)
                                                else:
                                                    df['Event_comment'] = 'МАПП Краскино'
                                                    df['Event_date'] = rcv_tm
                                                    print(df)
                                                con = sl.connect('CEL.db')
                                                with con:
                                                    df.to_sql('events2', con=con,
                                                              if_exists='append',
                                                              index=False)
                                                con.commit()
                                                quont +=1
                                            except:
                                                pass
                                        # maybe rar?
                                        elif substring_party in att_fn and rar_string_trigger in att_fn or substring_party in att_fn and 'zip' in att_fn:
                                            print(substring_party, att_fn)
                                            download_path = f"{download_folder}/{att_fn}"
                                            with open(download_path, "wb") as fp:
                                                fp.write(message.get('content').read())
                                                logging.info(download_path)
                                                logging.info(rcv_tm)
                                                time.sleep(3)
                                                try:
                                                    folder_rar = f'{main_folder}/{att_fn}-inculds'  # {str(att_fn).replace(".rar", "")}-new'
                                                    if not os.path.isdir(folder_rar):
                                                        os.makedirs(folder_rar, exist_ok=True)
                                                    try:
                                                        rf = rarfile.RarFile(download_path)
                                                        rf.extractall(path=folder_rar) #members=filename
                                                        time.sleep(5)
                                                    except:
                                                        with zipfile.ZipFile(download_path, 'r') as zip_file:
                                                            zip_file.extractall(folder_rar)
                                                            time.sleep(5)
                                                except Exception as e:
                                                    print(f'file read error: {e}')
                                                    #patoolib.extract_archive(download_path, outdir=folder_rar)

                                            files_in_glob = glob.glob(f'{folder_rar}/**/*.xlsx', recursive=True)
                                            print(files_in_glob)
                                            for file_glob in files_in_glob:
                                                print(file_glob)
                                                if pac_string_trigger in file_glob:
                                                    print(pac_string_trigger)
                                                    df = pd.read_excel(file_glob, sheet_name=0, engine='openpyxl',
                                                                       skiprows=8)
                                                    df = df[df.columns[3]].to_frame()
                                                    df = df.drop_duplicates(keep='first')
                                                    df = df.rename(
                                                        columns={df.columns[0]: 'parcel_numb'})
                                                    df['Event'] = 'Таможенный транзит'
                                                    if substring_party == substring_party_list[2]:
                                                        print(substring_party)
                                                        df['Event_comment'] = 'порт Владивосток'
                                                        df['Event_date'] = rcv_tm
                                                        print(df)
                                                    else:
                                                        df['Event_comment'] = 'МАПП Краскино'
                                                        df['Event_date'] = rcv_tm
                                                        print(df)
                                                    con = sl.connect('CEL.db')
                                                    with con:
                                                        df.to_sql('events2', con=con,
                                                                  if_exists='append',
                                                                  index=False)
                                                    con.commit()
                                                    quont += 1
                                        else:
                                            pass
                            except Exception as e:
                                print(traceback.print_exc())
                                logger.warning(f'error {e}')
                except Exception as e:
                    print(traceback.print_exc())
                    logger.warning(f'error {e}')
        except Exception as e:
            print(traceback.print_exc())
            logger.warning(f'error {e}')
    except Exception as e:
        print(traceback.print_exc())
        logger.warning(f'error {e}')
    message = f'Статусы Таможенный транзит за {now_date} в кол-ве {quont} заружены в базу данных'
    bot = telebot.TeleBot('5555513345:AAFzGfbHd4rUzLHh2m4q5kWEFtp7IUx_UNU')
    #bot.send_message(1634121947, message)  # 1285743017
    #bot.send_message(422263274, message)  # 1285743017
    #bot.send_message(1285743017, message)  # 1285743017

def logistick_scan_manifest():
    now = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
    now_date = datetime.datetime.now()
    host = "imap.yandex.ru"
    username = "cel-python-automatization@yandex.ru"
    password = "vmpqeopkfptvejiz"
    main_folder = "/РЕЕСТРЫ для статусов2"
    download_folder = main_folder
    logging.info(download_folder)
    if not os.path.isdir(download_folder):
        os.makedirs(download_folder, exist_ok=True)

    mail = Imbox(host, username=username, password=password, ssl=True, ssl_context=None, starttls=False)
    #print(mail.folders())
    quont = 0
    try:
        messages_all = mail.messages(date__on=datetime.date(now_date.year, now_date.month, now_date.day)) # sent_to='transpriemka@mail.ru', folder='&BB8EIAQYBBMEIAQQBB0EGAQnBCwEFQ-"
    except:
        current_year = now_date.year
        #print(current_year)
        #print(now_date.month - 1)
        before_month_numb = now_date.month - 1  # int(input())
        before_month_days = monthrange(current_year, before_month_numb)[1]
        #print(before_month_days)
        #print(datetime.date(now_date.year, now_date.month - 1, before_month_days))
        messages_all = mail.messages(date__on=datetime.date(now_date.year, before_month_numb, before_month_days))
    try:
        #print(len(messages_all))
        for (uid, message) in messages_all:
            try:
                #print(message.date)
                # rcv_tm1 = parser.parse(message.date)
                # print(rcv_tm1)
                reciever = message.sent_to[0]['email']
                #print(reciever)
                date_of_message = message.date
                rcv_tm = datetime.datetime.strptime(date_of_message[0:31], '%a, %d %b %Y %H:%M:%S %z')
                # print(start_date)
                try:
                    mail.mark_seen(uid)  # optional, mark message as read
                    for message in message.attachments:
                        att_fn = message.get('filename')
                        print(att_fn)
                        file = message.get('content')
                        if 'Manifest' in att_fn:
                            print('Manifest')
                            try:
                                download_path = f"{download_folder}/{att_fn}"
                                with open(download_path, "wb") as fp:
                                    fp.write(message.get('content').read())
                                    logging.info(download_path)
                                    logging.info(rcv_tm)
                                    time.sleep(2)
                                df = pd.read_excel(file, sheet_name=0, engine='openpyxl', skiprows=2)
                                #print(df)
                                df = df[df.columns[1]].to_frame()
                                df = df.drop_duplicates(keep='first')
                                df = df.rename(
                                    columns={df.columns[0]: 'parcel_numb'})
                                df['Event'] = 'Отгружен с Таможенного склада для доставки по последней миле'
                                df['Event_comment'] = 'Уссурийск'
                                df['Event_date'] = rcv_tm
                                con = sl.connect('CEL.db')
                                with con:
                                    df.to_sql('events2', con=con,
                                              if_exists='append',
                                              index=False)
                                con.commit()
                                custom_status = 'Отгружен с Таможенного склада для доставки по последней миле'
                                refuse_reason = 'Уссурийск'
                                decision_date = rcv_tm
                                # for parcel_numb in df['parcel_numb']:
                                #     #print(parcel_numb)
                                #     try:
                                #         Django_send_status(parcel_numb, custom_status, refuse_reason, decision_date)
                                #     except ValueError as e:
                                #         pass
                                try:
                                    with con_gps:
                                        df_gps = pd.read_excel(file, sheet_name=0, engine='openpyxl', header=None)
                                        gps_numb = str(df_gps.loc[0, 5])

                                        api_track718_add_track(gps_numb)
                                        df['gps_numb'] = gps_numb
                                        df = df[['parcel_numb', 'gps_numb']][:-1]
                                        df.to_sql('gps_parcels', con=con_gps,
                                                  if_exists='append',
                                                  index=False)
                                except Exception as e:
                                    logger.info(f'error {traceback.print_exc()}')
                                quont += 1
                            except Exception as e:
                                logger.info(f'error {traceback.print_exc()}')
                except Exception as e:
                    logger.info(f'error {traceback.print_exc()}')
            except Exception as e:
                logger.info(f'error {traceback.print_exc()}')
    except Exception as e:
        logger.info(f'error {traceback.print_exc()}')


map_eng_to_rus = {'parcel_numb': 'Накладная',
                        'parc_create': 'Заказ Создан',
                        'parc_hunch': 'Отгружен в РФ из Хунчунь',
                        'parc_svh': 'Прибыл на СВХ',
                        'parc_start_custm': 'Начато таможенное оформление',
                        'parc_finish_custm': 'Завершено таможенное оформление'}
map_includs_eng_to_rus = {'parcel_numb': 'Номер накладной', 'second_name': 'Фамилия', 'first_name': 'Имя', 'middle_name': 'Отчество',
                  'reciver_adress': 'Адрес', 'reciver_city': 'Город', 'reciver_state': 'Область', 'reciver_index': 'Индекс', 'phone_numb': 'Телефон',
                  'goods_quantity': 'Кол-во', 'goods_name': 'Описание товара', 'goods_price': 'Стоимость позиции',
                  'goods_link': 'Ссылка', 'pasport_seria': 'Серия пасп.', 'pasport_numb': 'Номер пасп.', 'pasport_date': 'Дата выдачи пасп.',
                          'reciver_birthday_date': 'Дата рождения',
                  'INN': 'ИНН', 'goods_weight': 'Вес товара', 'manifest_numb': 'Номер партии', 'manifest_date': 'Дата партии'}


json_1 = {'name': 'test', 'email': 'test@gmail.com', 'password': '12345'}


@app.route('/export/<string:file_name>', methods=['POST', 'GET'])
def export_excel(file_name):
    now = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
    #file_name = request.form['file_name']
    df = pd.read_excel(f'{file_name}.xlsx')
    # Create a new workbook and add a worksheet
    df.to_excel(f'{file_name} от {now}.xlsx', index=False)

    # Return the Excel file to the client
    return send_file(f'{file_name}.xlsx', as_attachment=True)

def get_db_connection():
    conn = sl.connect('CEL.db')
    conn.row_factory = sl.Row
    return conn

@app.route('/')
def index():
    conn = get_db_connection()
    events = conn.execute('SELECT * FROM events2').fetchmany(20)
    conn.close()
    return render_template('index.html', events=events)


@app.route('/search', methods=['GET'])
def parc_searh():
    try:
        parcel_numb = request.args.get('parcel_numb')
        return render_template('parc_search.html', search=parcel_numb)
    except Exception as e:
        logger.warning(f'read action faild with error: {e}')
        return {'message': str(e)}, 400

@app.route('/add/event', methods=['GET', 'POST'])
def parc_add_event():
    now = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
    events_list = ["Заказ Создан",
                    "Сформирован к отгрузке в РФ",
                    "Отгружен в РФ",
                    "Таможенный транзит",
                    "Прибыл на СВХ",
                    "Начато таможенное оформление",
                    "Таможенные платежи оплачены",
                   "Требуется уплата таможенных платежей",
                    "Завершено таможенное оформление",
                   "Выпуск товаров разрешён, таможенные пошлины уплачены",
                   'Отгружен с Таможенного склада для доставки по последней миле',
                    "Прибыл на коммерческий склад",
                    "Выдано получателю",
                   "г. Хабаровск",
                   "г. Чита",
                   "г. Иркутск",
                   "г. Красноярск",
                   "г. Новосибирск",
                   "г. Екатеринбург",
                   "г. Казань",
                   "г. Нижний Новгород",
                   "г. Москва"
                   ]
    try:
        if request.method == 'POST':
            uploaded_file = request.files['file']
            filename = uploaded_file.filename
            if filename != '':
                file_ext = os.path.splitext(filename)[1]
                if file_ext not in app.config['UPLOAD_EXTENSIONS']:
                    abort(400)
                uploaded_file.save(uploaded_file.filename)
                con = sl.connect('CEL.db')
                # открываем базу
                with con:
                    data_events = con.execute(
                        "select count(*) from sqlite_master where type='table' and name='events2'")
                    for row in data_events:
                        # если таких таблиц нет
                        if row[0] == 0:
                            # создаём таблицу
                            with con:
                                con.execute("""
                                                    CREATE TABLE events2 (
                                                        ID INTEGER PRIMARY KEY AUTOINCREMENT,
                                                        parcel_numb VARCHAR(36),
                                                        Event VARCHAR(40),
                                                        Event_comment VARCHAR(60),
                                                        Event_date DATETIME,
                                                        UNIQUE(parcel_numb, Event) ON CONFLICT REPLACE
                                                    );
                                                """)
                df = pd.read_excel(filename,
                                   engine='openpyxl', usecols='A')
                df_load = df.rename(columns={df.columns[0]: 'parcel_numb'})
                print(df_load)
                Event = request.form['event']
                Event_date = request.form['event_date']
                Event_comment = request.form['event_comment']
                df_load['Event'] = Event
                Event_date = datetime.datetime.strptime(Event_date, "%d.%m.%Y %H:%M").replace(tzinfo=pytz.UTC).astimezone(pytz.timezone("Europe/London"))
                df_load['Event_date'] = Event_date
                df_load['Event_comment'] = Event_comment
                # подготавливаем множественный запрос
                sql_insert = 'INSERT INTO events2 (parcel_numb, Event, Event_date, Event_comment) values(?, ?, ?, ?)'
                # указываем данные для запроса
                print(df_load['Event_date'])
                # df_load_tobase = df_load.rename(columns={0: 'parcel_numb', f'{}': f'{}'})
                # print(df_load_tobase)
                # list_of_load_to_base = list(df_load_tobase.itertuples(index=False, name=None))
                # Insert добавляем с помощью множественного запроса все данные сразу
                with con:
                    try:
                        df_load.to_sql('events2', con=con, if_exists='append', index=False)
                    except sl.DatabaseError as e:
                        logger.warning(f'error add event: {e}')
                    else:
                        con.commit()
                # закрытие соединения
                con.close()
    except Exception as e:
        logger.warning(f' parc_add_event - read action faild with error: {e}')
        return {'message': str(e)}, 400
    return render_template('add_event.html', now=now, events_list=events_list)
@app.after_request
def after_request(response):
    response.headers["Access-Control-Allow-Origin"] = "*" # <- You can change "*" for a domain for example "http://localhost"
    response.headers["Access-Control-Allow-Credentials"] = "true"
    response.headers["Access-Control-Allow-Methods"] = "POST, GET, OPTIONS, PUT, DELETE"
    response.headers["Access-Control-Allow-Headers"] = "Accept, Content-Type, Content-Length, Accept-Encoding, X-CSRF-Token, Authorization"
    return response


def insert_event(parcel_numb, Event, Event_comment, Event_date):
    conn = sl.connect('CEL.db')
    cur = conn.cursor()
    statement = "INSERT INTO events2 (parcel_numb, Event, Event_comment, Event_date) VALUES (?, ?, ?, ?)"
    cur.execute(statement, [parcel_numb, Event, Event_comment, Event_date])
    conn.commit()
    conn.close()

    return True

def send_json_to_SVH(event_details):
    response = requests.post('http://164.132.182.145:5001/api/add_decision', json=event_details,
                             headers={'accept': 'application/json'})
    try:
        return response.json()
    except ValueError:
        pass

# upload decisions in SVH BAZA for sorting
def update_decision_API(parcel_numb, custom_status, custom_status_short, refuse_reason, decision_date, regnumber):
    try:
        con = sl.connect('BAZA.db')
        cur = con.cursor()
        row_isalready_in = pd.read_sql(f"Select * from baza where parcel_numb = '{parcel_numb}'", con)
        print(row_isalready_in)
        registration_numb = regnumber
        print(registration_numb)
        if row_isalready_in.empty:
            statement = "INSERT INTO baza (registration_numb, parcel_numb, custom_status, custom_status_short, refuse_reason, decision_date) VALUES (?, ?, ?, ?, ?, ?)"
            cur.execute(statement, [registration_numb, parcel_numb, custom_status, custom_status_short, refuse_reason, decision_date])
            con.commit()
            con.close()
        else:
            con.execute(f"Update baza set"
                        f" registration_numb = '{regnumber}',"
                        f" custom_status = '{custom_status}',"
                        f" custom_status_short = '{custom_status_short}',"
                        f" decision_date = '{decision_date}',"
                        f" refuse_reason = '{refuse_reason}' where parcel_numb = '{parcel_numb}'")
            con.commit()
            con.close()
            logger_API_insert.info(f'baza updated: {parcel_numb}')
            print('execute OK')
    except Exception as e:
        print(e)
        logger_API_insert.info(f'insert_event_API action: {parcel_numb} fale: {e}')
        time.sleep(2)
        update_decision_API(parcel_numb, custom_status, custom_status_short, refuse_reason, decision_date, regnumber)
    return True

@app.route('/webhook', methods= ['POST'])
def get_webhook():
    if request.method == 'POST':
        response = request.json
        print("received data: ", response)
        with open("web_hook_response.txt", 'w') as web_hook_response:
            web_hook_response.write(str(response))
        return 'success', 200
    else:
        abort(400)


def tochina_prepare(parcel_numb, custom_status, refuse_reason, Event_date_chin):
    track_codes = {'платеж': 'SHND', 'паспорт': 'HCCR', 'пояснение': 'HCSM', 'техническ': 'HCCP', 'перечень категории товаров': 'HCGR',
                   'входит в перечень': 'HCGR', 'товары не относятся': 'HCGS', 'для личного пользования': 'HCGS',
                    'квитанц': 'HCFK', 'ссылк': 'HCRU', 'скриншот': 'HCRU'}
    track_chinees = {'HCPR': '延期放行', 'SHND': '需缴纳关税', 'HCCR': '护照无效', 'HCSM': '产品个人使用说明',
                     'HCCP': '需提供产品说明书', 'HCGR': 'B2B（商品',
                     'HCGS': 'B2B（数量', 'HCFK': '需提供付款凭证', 'HCRU': '需提供网址链接'}
    try:
        track_code = ''
        decision_date = datetime.datetime.strftime(Event_date_chin, "%Y/%#m/%#d %H:%M")
        if refuse_reason == 'nan':
            refuse_reason = ''
        else:
            for key, item in track_codes.items():
                if key in str.lower(refuse_reason):
                    track_code = item
                    break

        if 'выпуск товаров' in str.lower(custom_status):
            track_code = 'RC'
        elif 'продление' in str.lower(custom_status):
            track_code = 'HCPR'


        for key, item in track_chinees.items():
            if track_code != 'RC' and track_code == key:
                refuse_reason = item + refuse_reason
                break
            if track_code == 'RC':
                custom_status = custom_status + '放行'
                break
        print(track_code)
        print(f"{custom_status + '. ' + refuse_reason}")
        data = {"PostingNumber": f"{parcel_numb}", "TrackingNumber": f"{parcel_numb}",
                "Data": [{"track_code": f"{track_code}", "datetime": f"{decision_date}", "location": "Россия",
                          "description": f"{custom_status + '. ' + refuse_reason}"}]}
        return data
    except Exception as e:
        logger_API_insert.info(f'insert_event_API action faled: {parcel_numb}: {e}')


def send_to_china(data):
    try:
        print(data)
        data_str = str(data).replace("'", '"').replace(", ", ",")
        m = hashlib.md5()
        m.update(data_str.encode('utf-8'))
        result = base64.urlsafe_b64encode(m.hexdigest().encode('utf-8')).decode(
            'utf-8')  # b64encode(m.hexdigest().encode('utf-8'))
        url = ("http://hccd.rtb56.com/webservice/edi/TrackService.ashx?code=ADDCUSTOMSCLEARANCETRACK"
               + f'&data={data_str}' + f'&sign={str(result)}')
        response = requests.get(url)
        logger_API_insert.info(f'insert_event_API action: {response.text}')
        print(response.text)
        print('ok')
    except Exception as e:
        logger_API_insert.info(f'insert_event_API action faled: {data}: {e}')
        pass


def Django_send_status(parcel_numb, custom_status, refuse_reason, decision_date):
    url = 'http://127.0.0.1:8000/api_insert_decisions/'
    decision_date = datetime.datetime.strftime(decision_date, "%Y-%m-%d %H:%M:%S")
    body = {'parcel_numb': parcel_numb, 'time': decision_date,
            'status_name': custom_status, 'place': '', 'comment': refuse_reason}
    print(body)
    response = requests.post(url, json=body)
    print(response.status_code)
    print(response.json())


def svh_server_send_status(body):
    url = 'http://127.0.0.1:5001/api_insert_decisions_chunk/'
    print(body)
    response = requests.post(url, json=body)
    print(response.status_code)
    resp = response.json()
    print(resp)
    return resp


def insert_event_API(parcel_numb, Event, Event_comment, Event_date, internal_event, regnumber):
    try:
        Event_date_chin = datetime.datetime.strptime(Event_date, "%Y-%m-%d %H:%M:%S")
        Event_date = datetime.datetime.strptime(Event_date, "%Y-%m-%d %H:%M:%S").astimezone(pytz.timezone("Europe/London")) #.replace(tzinfo=pytz.UTC)
        if Event_comment == 'nan':
            Event_comment = ''
        #result = insert_event(parcel_numb, Event, Event_comment, Event_date)
        custom_status = Event
        if 'Выпуск' in str(custom_status):
            custom_status_short = 'ВЫПУСК'
        else:
            custom_status_short = 'ИЗЪЯТИЕ'
        refuse_reason = Event_comment
        decision_date = Event_date
        resp = True #svh_server_send_status(parcel_numb, custom_status, custom_status_short, refuse_reason, decision_date, regnumber)
        #tochina_send_status(parcel_numb, custom_status, refuse_reason, Event_date_chin)
        #Django_send_status(parcel_numb, custom_status, refuse_reason, decision_date, internal_event, regnumber)
        #update_decision_API(parcel_numb, custom_status, custom_status_short, refuse_reason, decision_date, regnumber)
        logger_API_insert.info(f'insert_event_API action: {parcel_numb}')
        return jsonify(resp)
    except Exception as e:
        print(e)
        logger_API_insert.info(f'insert_event_API action faled: {parcel_numb}: {e}')


@app.route('/api/payresult', methods=['GET', 'POST'])
def payresult():
    now = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
    try:
        details = request.get_json()
        print(details)
        posting_number = details['orderId']
        pay_status_code = details['paymentStatus']
        is_paid = "N"
        if pay_status_code == 5:
            is_paid = "Y"
            url = 'http://hccd.rtb56.com/webservice/Ozon/OzonUpdatePayTaxData.ashx'
            data = [
                {
                    "tracking_number": f"{posting_number}",
                    "is_paid": f"{is_paid}"
                }
            ]
            response = requests.post(url=url, json=data)
        logger_API_insert.info(f'{now} pay result OK: {posting_number}: {details}')
        return "OK"
    except Exception:
        logger_API_insert.info(f'{now} pay_result faled: {str(traceback.format_exc())}')
        return (str(traceback.format_exc()))


@app.route('/api/creating_pay_info', methods=['GET', 'POST'])
def creating_pay_info():
    now = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
    try:
        details = request.get_json()
        print(details)
        parcel_numb = details['parcel_numb']
        pay_sum = details['pay_sum']
        phone = details['phone']
        is_paid = "N"
        expired_date = datetime.datetime.now() + delta
        expired_date = expired_date.strftime("%Y-%m-%d")
        print(expired_date)
        url = "http://hccd.rtb56.com/webservice/Ozon/OzonSavePayTaxData.ashx"
        data = [
            {
                "posting_number": "",
                "tracking_number": parcel_numb,
                "pay_tax_end_time": expired_date,
                "pay_tax_link": f"https://cellog.ed22.ru",
                "tax_amount": pay_sum,
                "is_paid": f"{is_paid}"
            }
        ]
        response = requests.post(url=url, json=data)
        print(response.text)
        logger_API_insert.info(f'{now} creating_pay_info OK: {parcel_numb}: {details}')
        return "OK"
    except Exception:
        logger_API_insert.info(f'{now} creating_pay_info faled: {str(traceback.format_exc())}')
        return (str(traceback.format_exc()))



@app.route('/api/add/new_event_chanks', methods=['POST'])
def insert_event_API_chanks():
    try:
        event_details = request.get_json()
        df = pd.read_json(event_details)
        print(event_details)
        print(df)
        for index, row in df.iterrows():
            (regnumber, parcel_numb, Event, Event_date, Event_comment) = row
            Event_date_chin = datetime.datetime.strptime(Event_date, "%Y-%m-%d %H:%M:%S")
            Event_date = datetime.datetime.strptime(Event_date, "%Y-%m-%d %H:%M:%S").astimezone(pytz.timezone("Europe/London"))
            if Event_comment == 'nan':
                Event_comment = ''
            insert_event(parcel_numb, Event, Event_comment, Event_date)
            custom_status = Event
            refuse_reason = Event_comment
            decision_date = Event_date
            if 'Выпуск' in str(custom_status):
                custom_status_short = 'ВЫПУСК'
            else:
                custom_status_short = 'ИЗЪЯТИЕ'
            update_decision_API(regnumber, parcel_numb, custom_status, custom_status_short, refuse_reason, decision_date)
            tochina_prepare(parcel_numb, custom_status, refuse_reason, Event_date_chin)
        return jsonify(True)
    except Exception as e:
        print(e)


@app.route('/api/add/new_event_chunks2', methods=['POST'])
def insert_event_API_chunks2():
    now = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
    try:
        event_details = request.get_json()
        #df = pd.read_json(event_details)
        print(event_details)
        parcels_list = []
        data_list = []
        body = []
        for parcel in event_details:
            regnumber = parcel["regnumber"]
            parcel_numb = parcel["parcel_numb"]
            print(parcel_numb)
            custom_status = parcel["Event"]
            refuse_reason = parcel["Event_comment"]
            decision_date = parcel["Event_date"]
            Event_date_chin = datetime.datetime.strptime(decision_date, "%Y-%m-%d %H:%M:%S")
            decision_date = datetime.datetime.strptime(decision_date, "%Y-%m-%d %H:%M:%S").astimezone(pytz.timezone("Australia/Sydney"))
            Event = custom_status
            Event_comment = refuse_reason
            Event_date = decision_date
            insert_event(parcel_numb, Event, Event_comment, Event_date)
            if 'Выпуск' in str(custom_status):
                custom_status_short = 'ВЫПУСК'
            else:
                custom_status_short = 'ИЗЪЯТИЕ'
            #update_decision_API(regnumber, parcel_numb, custom_status, custom_status_short, refuse_reason, decision_date)
            parcels_list.append(parcel_numb)
            data = tochina_prepare(parcel_numb, custom_status, refuse_reason, Event_date_chin)
            data_list.append(data)
            decision_date = datetime.datetime.strftime(decision_date, "%Y-%m-%d %H:%M:%S")
            for_svh_serv_data = {'registration_numb': regnumber, 'parcel_numb': parcel_numb, 'decision_date': decision_date,
            'custom_status': custom_status, 'custom_status_short': custom_status_short, 'place': '', 'refuse_reason': refuse_reason}
            body.append(for_svh_serv_data)
        svh_server_send_status(body)
        logger_API_chunks.info(f'{now} insert_event_API_chunks2: {str(parcels_list)}')
        print(data_list)
        send_to_china(data_list)
        return jsonify(True)
    except Exception as e:
        logger.info(f'{now} chunks2 faled: {e}')
        logger.info(f'insert_event_API_chunks2 faled: {traceback.print_exc()}')
        return jsonify(f'insert_event_API_chunks2 faled: {traceback.print_exc()}')



def get_token_agreg(client_id_agreg, client_secret_agreg):
    url = 'https://api-logistic-platform.ozon.ru/GetAuthToken'
    headers = {'content-type': 'application/json'}
    data = {
    "ClientId": client_id_agreg,
    "ClientSecret": client_secret_agreg
    }
    response = requests.post(url=url, json=data, headers=headers)
    token = response.text
    return token


def pars_and_send_pay(TrackingNumber, parcel, token, url, provider, event_details):
    PostingNumber = parcel['PostingNumber']
    TaxPayment = parcel['TaxPayment']
    CustomsDuty = parcel['CustomsDuty']
    Total = parcel['Total']
    Currency = parcel['Currency']
    InvoiceNumber = parcel['InvoiceNumber']
    DateOfPayment = parcel['DateOfPayment']
    RegisterNumber = parcel['RegisterNumber']

    body = {'TaxReports':
                [{'PostingNumber': PostingNumber,
                  'TrackingNumber': TrackingNumber,
                  'TaxPayment': TaxPayment,
                  'CustomsDuty': CustomsDuty,
                  'Total': Total,
                  'Currency': Currency,
                  'InvoiceNumber': InvoiceNumber,
                  'DateOfPayment': DateOfPayment,
                  'RegisterNumber': RegisterNumber
                  }]}
    headers = {'Authorization': f'Bearer {token}'}
    response = requests.post(url=url, json=body,
                             headers=headers)
    status_code = response.status_code
    print(status_code)
    print(response.text)
    ozone_response_text = response.text
    with con_pay:
        print('start insertion')
        query = """INSERT INTO pay_customs (PostingNumber, TrackingNumber, TaxPayment, CustomsDuty, 
                            Total, InvoiceNumber, RegisterNumber, Currency, DateOfPayment, Provider, send_time, json_info, 
                            ozone_response_status_code, ozone_response_text)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"""
        con_pay.execute(query, [PostingNumber, TrackingNumber, TaxPayment, CustomsDuty,
                                Total, InvoiceNumber, RegisterNumber, Currency, DateOfPayment,
                                provider, now, str(event_details), status_code, ozone_response_text])
        print('insert ok')
    return body, response, status_code
@app.route('/api/add/pay_customs_info', methods=['POST'])
def pay_customs_info():
    con_pay = sl.connect('Pay.db')
    now = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
    url = 'https://api-logistic-platform.ozon.ru/v1/TaxReport'
    try:
        event_details = request.get_json()

        print(event_details)
        logger_customs_paya_all.info(f'{now} pay_customs_info : {event_details}')
        parcels = event_details['TaxReports']
        for parcel in parcels:
            TrackingNumber = parcel['TrackingNumber']
            provider = parcel['Provider']
            if provider is not None:
                try:
                    token = ''
                    body, response, status_code = pars_and_send_pay(TrackingNumber, parcel, token, url, provider, event_details)
                    if status_code == 200:
                        logger_customs_pay.info(f'{now, provider, response.text}')
                        print(f'{now, provider, body}')
                    elif status_code == 401:
                        client_id_agreg = OZON_keys[provider]['clientId']
                        client_secret_agreg = OZON_keys[provider]['clientSecret']
                        token = json.loads(get_token_agreg(client_id_agreg, client_secret_agreg))
                        token = token["Data"]
                        with con_pay:
                            con_pay.execute(f"UPDATE cel_tokens SET token = '{token}' WHERE provider = '{provider}'")
                        body, response, status_code = pars_and_send_pay(TrackingNumber, parcel, token, url, provider,
                                                                        event_details)
                        logger_customs_pay.info(f'{now, provider, response.text}')
                        print(f'{now, provider, body}')
                    else:
                        error_info = f'{now} pay_customs_info_faled: {response.text}'
                        logger_pay_errors.info(error_info)
                except Exception as e:
                    print(e)
                    error_info = f'{now} pay_customs_info_faled: {TrackingNumber} {e}'
                    logger_pay_errors.info(error_info)
                    send_email(body_text=f'Error {error_info}', subject=f'{TrackingNumber} pay_customs error')
            else:
                msg = 'ok'
                logger_customs_pay.info(f'{now} provider is None, return {msg}')
        return jsonify('ok')
    except Exception as e:
        try:
            event_details = request.get_json()
            error_info = f'{now} pay_customs_info faled: {e} details: {event_details}'
            logger_pay_errors.info(error_info)
            send_email(body_text=f'{error_info} for: {event_details}', subject=f'pay_customs error')
            return jsonify(error_info)
        except:
            error_info = f'{now} pay_customs_info faled: {e}'
            logger_pay_errors.info(error_info)
            send_email(body_text=f'{error_info}', subject=f'pay_customs error')
            return jsonify(error_info)


@app.route('/api/add/new_event', methods=['POST'])
def json_request_evetnts():
    event_details = request.get_json()
    parcel_numb = event_details["parcel_numb"]
    try:
        internal_event = event_details["internal_event"]
        regnumber = event_details["regnumber"]
    except:
        internal_event = 'unknown'
        regnumber = 'unknown'
    Event = event_details["Event"]
    Event_comment = event_details["Event_comment"]
    Event_date = event_details["Event_date"]

    insert_event_API(parcel_numb, Event, Event_comment, Event_date, internal_event, regnumber)
    return jsonify(True)


@app.route('/api/add/new_event_other', methods=['POST'])
def insert_event_API_event_other():
    event_details = request.get_json()
    parcel_numb = event_details["parcel_numb"]
    Event = event_details["Event"]
    Event_comment = event_details["Event_comment"]
    Event_date = datetime.datetime.strptime(event_details["Event_date"], "%Y-%m-%d %H:%M:%S").replace(tzinfo=pytz.UTC).astimezone(pytz.timezone("Europe/London")) #
    result = insert_event(parcel_numb, Event, Event_comment, Event_date)
    logger_API_insert.info(f'insert_event_API action: {event_details}')
    return jsonify(result)


@app.route('/api/v1.0/events/', methods=['POST'])
def get_parcel_info_API():
    parcel_details = request.get_json()
    parcel_numb = parcel_details['parcel_numb']
    try:
        con = sl.connect('CEL.db')
        with con:
            df_parc_events = pd.read_sql(f"SELECT * FROM events2 where parcel_numb = '{parcel_numb}'", con)
            df_parc_events['Event_date'] = pd.to_datetime(df_parc_events['Event_date'], utc=True).dt.tz_convert('US/Eastern')
            df_parc_events = df_parc_events.sort_values(by='Event_date')
            df_parc_events['Event_date'] = df_parc_events['Event_date'].astype(str).replace(to_replace=' ', value='T', regex=True)
            df_parc_events['Event_comment'] = df_parc_events['Event_comment'].replace(to_replace='10716050', value='т/п Уссурийский', regex=True)
            if df_parc_events['Event_comment'].str.contains('по уплате таможенных платежей').any():
                df_parc_events['Event'] = df_parc_events['Event'].replace(to_replace='Отказ в выпуске товаров',
                                                                                          value='Продолжается процесс таможенного оформления',
                                                                                          regex=True)
            elif df_parc_events['Event_comment'].str.contains('е уплачены').any():
                df_parc_events['Event'] = df_parc_events['Event'].replace(to_replace='Отказ в выпуске товаров',
                                                                          value='Продолжается процесс таможенного оформления',
                                                                          regex=True)
                df_parc_events.loc[df_parc_events.Event == 'Продолжается процесс таможенного оформления', 'Event_comment'] = ""
            df_parc_events = df_parc_events.rename(columns={"parcel_numb": "HWBRefNumber",
                                                            "Event": "EventText",
                                                            "Event_comment": "EventComment",
                                                            "Event_date": "EventTime"})
        #logger_API_get_info.info(f'get_parcel_info_API: someone see {parcel_details}')
        message = f'get_parcel_info_API: someone see {parcel_details}'
    except Exception as e:
        logger.info(f'get_parcel_info_API parcel {parcel_numb}  - read action faild with error: {e}')
        return {'message': str(e)}, 400
    return Response(df_parc_events.to_json(orient="records", indent=2), mimetype='application/json')


@app.route('/api/v1.0/eventsother/', methods=['POST'])
def get_parcel_info_API_other():
    parcel_details = request.get_json()
    parcel_numb = parcel_details['parcel_numb']
    try:
        con = sl.connect('CEL.db')
        with con:
            df_parc_events = pd.read_sql(f"SELECT * FROM events2 where parcel_numb = '{parcel_numb}'", con)
            df_parc_events['Event_date'] = pd.to_datetime(df_parc_events['Event_date'])
            df_parc_events = df_parc_events.sort_values(by='Event_date')
            df_parc_events['Event_date'] = df_parc_events['Event_date'].astype(str)
    except Exception as e:
        logger.warning(f'get_parcel_info_API parcel {parcel_numb}  - read action faild with error: {e}')
        return {'message': str(e)}, 400
    return Response(df_parc_events.to_json(orient="records", indent=2), mimetype='application/json')


@app.route('/add/manifest', methods=['GET', 'POST'])
def parc_add_manifest():
    now = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
    now_date = datetime.datetime.now()
    if request.method == 'POST':
        uploaded_file = request.files['file']
        filename = uploaded_file.filename
        if filename != '':
            file_ext = os.path.splitext(filename)[1]
            if file_ext not in app.config['UPLOAD_EXTENSIONS']:
                abort(400)
            uploaded_file.save(uploaded_file.filename)
            con = sl.connect('CEL.db')
            # открываем базу
            with con:
                # получаем количество таблиц с нужным нам именем
                data = con.execute("select count(*) from sqlite_master where type='table' and name='manifest_cel'")
                for row in data:
                    # если таких таблиц нет
                    if row[0] == 0:
                        # создаём таблицу
                        with con:
                            con.execute("""
                                            CREATE TABLE manifest_cel (
                                                ID INTEGER PRIMARY KEY AUTOINCREMENT,
                                                manifest_numb VARCHAR(20),
                                                manifest_date VARCHAR(15),
                                                parcel_numb VARCHAR(20) NOT NULL,
                                                second_name VARCHAR(20),
                                                first_name VARCHAR(20),
                                                middle_name VARCHAR(20),
                                                reciver_adress VARCHAR(35),
                                                reciver_city VARCHAR(15),
                                                reciver_state VARCHAR(20),
                                                reciver_index INT(6),
                                                phone_numb VARCHAR(15),
                                                goods_quantity INT(3),
                                                goods_name VARCHAR(60),
                                                goods_price FLOAT,
                                                goods_link VARCHAR(499),
                                                pasport_seria VARCHAR(10),
                                                pasport_numb VARCHAR(15),
                                                pasport_date VARCHAR(15),
                                                reciver_birthday_date VARCHAR(15), 
                                                INN VARCHAR(15),
                                                goods_weight FLOAT 
                                            );
                                        """)
                            con.commit()
                            # закрытие соединения
                            con.close()
            manifest_numb = 'test'
            manifest_date = 'test'
            df = pd.read_excel(filename, sheet_name=0, header=None, engine='openpyxl',
                               skiprows=1, usecols='A:I,K:T')
            manifest_numb = request.form['manifest_numb']
            manifest_date = datetime.datetime.strptime(request.form['manifest_date'], '%d.%m.%Y %H:%M').replace(tzinfo=pytz.UTC).astimezone(pytz.timezone("Europe/London"))
            df['manifest_numb'] = manifest_numb
            df['manifest_date'] = manifest_date
            df.columns = ['parcel_numb', 'second_name', 'first_name', 'middle_name',
                          'reciver_adress', 'reciver_city', 'reciver_state', 'reciver_index', 'phone_numb',
                          'goods_quantity', 'goods_name', 'goods_price',
                          'goods_link', 'pasport_seria', 'pasport_numb', 'pasport_date', 'reciver_birthday_date',
                          'INN', 'goods_weight', 'manifest_numb', 'manifest_date']
            df['goods_price'] = df['goods_price'].replace(to_replace=',,', value='.', regex=True)
            df['goods_price'] = df['goods_price'].replace(to_replace='\.\.', value='.', regex=True)
            df['goods_price'] = df['goods_price'].replace(to_replace='，', value='.', regex=True)
            df['goods_price'] = df['goods_price'].replace(to_replace=',', value='.', regex=True)
            df['goods_price'] = df['goods_price'].replace(to_replace='，', value='.', regex=True)
            df['goods_price'] = df['goods_price'].replace(to_replace='^\.', value='', regex=True)
            df['goods_price'] = df['goods_price'].replace(to_replace='^,', value='', regex=True)
            df['goods_price'] = df['goods_price'].replace(to_replace=',$', value='', regex=True)
            try:
                df['goods_price'] = df['goods_price'].replace(to_replace='\.$', value='', regex=True).astype('float')
            except ValueError as msg:
                logger.warning(f'{str(msg)}\nПроверьте выставлены ли столбцы по шаблону, \nзатем проверьте сам столбец с ценой!')
            df['goods_weight'] = df['goods_weight'].replace(to_replace=',,', value='.', regex=True)
            df['goods_weight'] = df['goods_weight'].replace(to_replace='\.\.', value='.', regex=True)
            df['goods_weight'] = df['goods_weight'].replace(to_replace='，', value='.', regex=True)
            df['goods_weight'] = df['goods_weight'].replace(to_replace=',', value='.', regex=True)
            df['goods_weight'] = df['goods_weight'].replace(to_replace='，', value='.', regex=True)
            df['goods_weight'] = df['goods_weight'].replace(to_replace='^\.', value='', regex=True)
            df['goods_weight'] = df['goods_weight'].replace(to_replace='^,', value='', regex=True)
            df['goods_weight'] = df['goods_weight'].replace(to_replace=',$', value='', regex=True)
            try:
                df['goods_weight'] = df['goods_weight'].replace(to_replace='\.$', value='', regex=True).astype('float')
            except ValueError as msg:
                logger.warning(f'{str(msg)}\nПроверьте выставлены ли столбцы по шаблону, \nзатем проверьте сам столбец с весом!')
            # добавляем фрэйм в базу
            with con:
                df.to_sql('manifest_cel', con=con, if_exists='append', index=False)
            # выводим содержимое таблицы с покупками на экран
            with con:
                data = con.execute("SELECT * FROM manifest_cel")
                for row in data:
                    print(row)
            con.commit()
            # закрытие соединения
            con.close()
    return render_template('add_manifest.html', now=now)


@app.route('/todo/events/', methods=['POST', 'GET'])
def get_parcel_info():
    parcel_numb = request.form['parcel_numb']
    try:
        con = sl.connect('CEL.db')
        with con:
            df_parc_events = pd.read_sql(f"SELECT * FROM events2 where parcel_numb = '{parcel_numb}'", con)
            df_parc_events = df_parc_events.rename(columns=map_eng_to_rus)
            df_parc_events['Event_date'] = pd.to_datetime(df_parc_events['Event_date'])
            df_parc_events = df_parc_events.sort_values(by='Event_date')
            #df_parc_events['Event_date'] = df_parc_events['Event_date'].dt.strftime('%d.%m.%Y %H:%M')
            print(df_parc_events['Event_date'])
            df_parc_includs = pd.read_sql(f"SELECT * FROM manifest_cel where parcel_numb = '{parcel_numb}'", con)
            df_parc_includs = df_parc_includs.rename(columns=map_includs_eng_to_rus)
    except Exception as e:
        logger.warning(f'parcel {parcel_numb}  - read action faild with error: {e}')
        return {'message': str(e)}, 400

    return render_template('parc_info.html', tables=[df_parc_events.to_html(classes='mystyle', index=False),
                                                     df_parc_includs.to_html(classes='mystyle', index=False)],
                           titles=['na', 'Отслеживание (статусы экспресс груза)', 'Информация о посылке'],
                           parcel_numb=parcel_numb)


@app.route('/todo/events/list', methods=['POST', 'GET'])
def get_parcel_info_list():
    parcel_numbs = request.form['parcel_numbs'].replace(' ', ',')
    parcels_list = parcel_numbs.split(",")
    df_all_parcels = pd.DataFrame()
    for parcel_numb in parcels_list:
        try:
            con = sl.connect('CEL.db')
            with con:
                df_parc_events = pd.read_sql(f"SELECT * FROM events2 where parcel_numb = '{parcel_numb}' ORDER BY ID DESC LIMIT 1", con)
                print(df_parc_events)
                df_parc_events = df_parc_events.rename(columns=map_eng_to_rus)
                df_parc_events['Event_date'] = pd.to_datetime(df_parc_events['Event_date'])
                df_all_parcels = df_all_parcels.append(df_parc_events)
        except Exception as e:
            logger.exception("message")
            logger.warning(f'parcel {parcel_numb}  - read action faild with error: {e}')
            return {'message': str(e)}, 400
    return render_template('parc_info.html', tables=[df_all_parcels.to_html(classes='mystyle', index=False)],
                           titles=['na', 'Отслеживание (статусы экспресс груза)', 'Информация о посылке'],
                           parcel_numb=parcel_numb)

@app.route('/todo/pay/list', methods=['POST', 'GET'])
def get_pay_info_list():
    parcel_numbs = request.form['parcel_numbs'].replace(' ', ',')
    parcels_list = parcel_numbs.split(",")
    df_all_parcels = pd.DataFrame()
    for parcel_numb in parcels_list:
        try:
            con_pay = sl.connect('Pay.db')
            with con_pay:
                df_parc_events = pd.read_sql(f"SELECT * FROM pay_customs where TrackingNumber = '{parcel_numb}'", con_pay)
                print(df_parc_events)
                df_all_parcels = df_all_parcels.append(df_parc_events)
        except Exception as e:
            logger.exception("message")
            logger.warning(f'parcel {parcel_numb}  - read action faild with error: {e}')
            return {'message': str(e)}, 400
    format_func = lambda x: x[0:30] + '...'
    file_name = 'OzonPayInfo'
    writer = pd.ExcelWriter(f'{file_name}.xlsx', engine='xlsxwriter')
    df_all_parcels.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.save()
    return render_template('parc_info.html', tables=[df_all_parcels.to_html(formatters={'json_info': format_func}, classes='mystyle', index=False)],
                           titles=['na', '', 'Информация о посылке'],
                           parcel_numb=parcel_numb, file_name=file_name)

@app.route('/todo/events/list_api', methods=['POST', 'GET'])
def get_parcel_info_list_api():
    parcels_list_dict = request.get_json()
    print(parcels_list_dict)
    df = pd.DataFrame(parcels_list_dict)
    con = sl.connect('CEL.db')
    with con:
        df.to_sql('temp_table_parcels_api_cel', con, if_exists="replace")
        query = ("""SELECT events2.parcel_numb, events2.Event, events2.Event_comment, events2.Event_date
                    FROM temp_table_parcels_api_cel
                    LEFT JOIN events2
                    ON events2.parcel_numb = temp_table_parcels_api_cel.parcel_numb""")
        df_result = pd.read_sql(query, con)
        print(df_result)

    return Response(df_result.to_json(orient="records", indent=2), mimetype='application/json')


@app.route('/todo/api/v1.0/events/<string:parcel_numb>', methods=['POST', 'GET'])
def from_main_parcel_info(parcel_numb):
    try:
        con = sl.connect('CEL.db')
        with con:
            df_parc_events = pd.read_sql(f"SELECT * FROM events2 where parcel_numb = '{parcel_numb}'", con)
            df_parc_events = df_parc_events.rename(columns=map_eng_to_rus) # доработать мэпинг
            df_parc_events['Event_date'] = pd.to_datetime(df_parc_events['Event_date'])
            df_parc_events = df_parc_events.sort_values(by='Event_date')
            print(df_parc_events['Event_date'])
            df_parc_includs = pd.read_sql(f"SELECT * FROM manifest_cel where parcel_numb = '{parcel_numb}'", con)
            df_parc_includs = df_parc_includs.rename(columns=map_includs_eng_to_rus)
    except Exception as e:
        logger.warning(f'parcel {parcel_numb}  - read action faild with error: {e}')
        return {'message': str(e)}, 400
    return render_template('parc_info.html', tables=[df_parc_events.to_html(classes='data', index=False),
                                                     df_parc_includs.to_html(classes='data', index=False)],
                           titles=['na', 'Отслеживание (статусы экспресс груза)', 'Информация о посылке'],
                           parcel_numb=parcel_numb)

@app.route('/todo/api/v1.0/manifest_info/', methods=['POST', 'GET'])
def get_manifest_info():
    manifest_numb = request.form['manifest_numb']
    try:
        con = sl.connect('CEL.db')
        with con:
            df_parc__manifest_includs = pd.read_sql(f"SELECT * FROM manifest_cel where manifest_numb = '{manifest_numb}'", con)
            df_parc__manifest_includs = df_parc__manifest_includs.rename(columns=map_includs_eng_to_rus)
    except Exception as e:
        logger.warning(f'parcel {manifest_numb}  - read action faild with error: {e}')
        return {'message': str(e)}, 400

    return render_template('manifest_info.html', tables=[df_parc__manifest_includs.to_html(classes='data', index=False)],
                           titles=df_parc__manifest_includs.columns.values,
                           manifest_numb=manifest_numb)



@app.route('/todo/api/v1.0/tutorials', methods=['GET'])
@jwt_required()
@marshal_with(VideoSchema(many=True))
def get_list():
    try:
        user_id = get_jwt_identity()
        videos = Video.query.get_user_list(user_id=user_id)
    except Exception as e:
        logger.warning(f'user:{user_id}: tutorials - read action faild with error: {e}')
        return {'message': str(e)}, 400
    return videos

@app.route('/todo/api/v1.0/tutorials', methods=['POST'])
@jwt_required()
@use_kwargs(VideoSchema)
@marshal_with(VideoSchema)
def update_list(**kwargs):
    try:
        user_id = get_jwt_identity()
        new_one = Video(user_id=user_id, **kwargs)
        new_one.save()
    except Exception as e:
        logger.warning(f'user:{user_id}: tutorials - create action faild with error: {e}')
        return {'message': str(e)}, 400
    return new_one

@app.route('/todo/api/v1.0/tutorials/<int:tutorial_id>', methods=['PUT'])
@jwt_required()
@use_kwargs(VideoSchema)
@marshal_with(VideoSchema)
def update_tutorial(tutorial_id, **kwargs):
    try:
        user_id = get_jwt_identity()
        item = Video.get(tutorial_id, user_id)
        item.update(**kwargs)
    except Exception as e:
        logger.warning(f'user:{user_id}: tutorial {tutorial_id} - update action faild with error: {e}')
        return {'message': str(e)}, 400
    return item


@app.route('/todo/api/v1.0/tutorials/<int:tutorial_id>', methods=['DELETE'])
@jwt_required()
@marshal_with(VideoSchema)
def delete_tutorial(tutorial_id):
    try:
        user_id = get_jwt_identity()
        item = Video.get(tutorial_id, user_id)
        item.delete()
    except Exception as e:
        logger.warning(f'user:{user_id}: tutorial {tutorial_id} - delete action faild with error: {e}')
        return {'message': str(e)}, 400
    return '', 204


@app.route('/todo/api/v1.0/register', methods=['POST'])
@use_kwargs(UserSchema)
@marshal_with(AuthSchema)
def register(**kwargs):
    try:
        user = User(**kwargs)
        session.add(user)
        session.commit()
        token = user.get_token()
    except Exception as e:
        logger.warning(f'register error: {e}')
        return {'message': str(e)}, 400
    return {'access_token': token}

@app.route('/home')
def home():
    return render_template("home.html")

@app.route('/login', methods=['GET', 'POST'])
def login_start():
    return render_template("login.html")

@app.route('/todo/api/v1.0/login_insert', methods=['POST'])
def login_insert():
    email = request.form['email']
    password = request.form['password']
    log = client.post('/todo/api/v1.0/login', json={'email': email, 'password': password})
    access_token = log.get_json()['access_token']
    resp = make_response(redirect(url_for('parc_searh')))
    set_access_cookies(resp, access_token)
    return resp


@app.route('/todo/api/v1.0/login', methods=['POST'])
@use_kwargs(UserSchema(only=('email', 'password')))
@marshal_with(AuthSchema)
def login(**kwargs):
    user = User.authenticate(**kwargs)
    token = user.get_token()
    return {'access_token': token}

@app.route('/logout')
@jwt_required()
def logout():
    user_id = get_jwt_identity()
    print(user_id)
    resp = make_response(redirect(url_for('login_start')))
    #resp.set_cookie('access_token', max_age=0)
    unset_jwt_cookies(resp)
    return resp
@app.teardown_appcontext
def shutdown_session(exception=None):
    session.remove()


@app.errorhandler(422)
def error_handlers(err):
    headers = err.data.get('headers', None)
    messages = err.data.get('messages', ['Invalid request'])
    if headers:
        return jsonify({'message': messages}), 400, headers
    else:
        return jsonify({'message': messages}), 400


def api_track718(gps_numb):
    cel_api_key = "e0fca820-c3dc-11ee-b960-bdfb353c94dc"
    url = "https://apigetway.track718.net/v2/tracking/query"
    headers = {"Content-Type": "application/json",
    "Track718-API-Key": f"{cel_api_key}"}

    params = [{"trackNum": gps_numb, "code": "gps-truck"}]
    respons = requests.post(url=url, headers=headers, json=params)


def gps_job():
    with con_gps:
        query = "Select DISTINCT gps_numb from gps_parcels"
        df = pd.read_sql(sql=query, con=con_gps)
        for gps_numb in df['gps_numb']:
            print(gps_numb)
            api_track718(gps_numb)


def backup():
    con = sl.connect("CEL.db")
    bck = sl.connect('CEL_backup.db')
    with bck:
        con.backup(bck, pages=1)
    bck.close()
    con.close()


def check_and_backup():
    con = sl.connect("CEL.db")
    cur = con.cursor()
    try:
        cur.execute("PRAGMA integrity_check")
        backup()
    except sl.DatabaseError:
        con.close()


def decod_xml_to_base64(document):
    # convert file content to base64 encoded string
    with open(document, "rb") as file:
        raw_file = file.read()
        encoded = base64.encodebytes(raw_file).decode("utf-8")
    return encoded

@app.route('/api/TaxDocuments', methods=['POST'])
def TaxDocuments():
    try:
        details = request.get_json()
        logger_tax_documents.info(details)
        document = details['documentData']
        #encoded = decod_xml_to_base64(document)
        token = json.loads(get_token_agreg(client_id_agreg, client_secret_agreg))
        token = token["Data"]
        print(token)
        headers = {'Authorization': f'Bearer {token}'}
        url = 'https://api-logistic-platform.ozon.ru/v1/TaxDocuments'

        # print(encoded)
        body = {'documentType': 'Cmn',
                'documentData': document}
        #body_json = json.dumps(body)
        # print(body)
        with open('json_xml.json', 'w', encoding='UTF-8') as fp:
            json.dump(body, fp, ensure_ascii=False)
        response = requests.post(url=url, headers=headers, json=body)
        print(body)
        print(response.status_code)
        print(response.text)
    except Exception as e:
        logger_tax_documents.info(e)
    return 'ok'



#logistick_scan()
#transpriemka_scan()
#logistick_scan_manifest()
#gps_job()
#api_track718("14000132175")
#api_track718_add_track("14000132175")

# Create the background scheduler
#scheduler = BackgroundScheduler(daemon=True)
# Create the job
#scheduler.add_job(func=transpriemka_scan, trigger='cron', hour='20', minute='34') #trigger='cron', hour='22', minute='30'
#scheduler.add_job(func=logistick_scan, trigger='cron', hour='20', minute='36')
#scheduler.add_job(func=logistick_scan_manifest, trigger='cron', hour='16', minute='48')
#Start the scheduler
#scheduler.start()
# /!\ IMPORTANT /!\ : Shut down the scheduler when exiting the app
#atexit.register(lambda: scheduler.shutdown())

docs.register(get_parcel_info)
docs.register(get_list)
docs.register(update_list)
docs.register(update_tutorial)
docs.register(delete_tutorial)
docs.register(register)
docs.register(login)


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)  # http://127.0.0.1:5000
