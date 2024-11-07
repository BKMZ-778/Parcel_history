import json
import os.path
import glob
import logging
import os
import os.path
import os.path
import sqlite3 as sl
import time
import traceback
import zipfile
from tkinter import filedialog
import pandas as pd
import pytz
import rarfile
import requests
import telebot
import xlrd
import tkinter as tk
from imbox import Imbox
import datetime
import shutil
from zipfile import ZipFile
from pathlib import Path
import xml.etree.ElementTree as ET
import tkinter.messagebox as mb


def chunks(lst, n):
    """Yield successive n-sized chunks from lst."""
    for i in range(0, len(lst), n):
        yield lst[i:i + n]


def insert_event_API_chunks():
    #file_name = 'Решения_по_посылкам_8550103-OZON-270.xlsx'
    file_name = filedialog.askopenfilename()
    df = pd.read_excel(file_name)
    print(df)
    df.columns = ['Рег. номер', 'Номер общей накладной', 'parcel_numb',
                  'Пломба', 'Вес брутто', 'Event', 'Event_date', 'Event_comment', 'code']
    df['regnumber'] = df['Рег. номер']
    df = df[['regnumber', 'parcel_numb', 'Event', 'Event_date', 'Event_comment']]
    df['Event_comment'] = df['Event_comment'].fillna('')
    df['Event_date'] = pd.to_datetime(df['Event_date'], format='%d.%m.%Y %H:%M')
    df['Event_date'] = df['Event_date'].dt.strftime("%Y-%m-%d %H:%M:%S")
    print(df['Event_date'])
    list_chanks = list(chunks(df, 1000))
    i = 0
    for chank in list_chanks:
        i += 1
        print(f'chunk {i}')
        body = chank.to_json(orient="columns", force_ascii=False)
        #print(body)
        #print(body)
        # 127.0.0.1:5000 # 164.132.182.145
        response = requests.post('http://164.132.182.145:5000/api/add/new_event_chanks', json=body, headers={'accept': 'application/json'})
        print(response.text)


def insert_event_API_chunks2():
    file_name = filedialog.askopenfilename()
    df = pd.read_excel(file_name)
    print(df)
    df.columns = ['regnumber', 'Номер общей накладной', 'parcel_numb',
                  'Пломба', 'Вес брутто', 'custom_status', 'decision_date', 'refuse_reason', 'decision_code']
    df = df[['regnumber', 'parcel_numb', 'custom_status', 'decision_date', 'refuse_reason', 'decision_code']]
    df['refuse_reason'] = df['refuse_reason'].fillna('')
    df['decision_date'] = pd.to_datetime(df['decision_date'], format='%d.%m.%Y %H:%M')
    df['decision_date'] = df['decision_date'].dt.strftime("%Y-%m-%d %H:%M:%S")
    df['decision_code'] = df['decision_code'].fillna('')
    print(df['decision_date'])
    parcel_list = []
    for index, row in df.iterrows():
        (regnumber, parcel_numb, custom_status, decision_date, refuse_reason, decision_code) = row
        parcel_info = {"regnumber": regnumber, "parcel_numb": parcel_numb,
                       "Event": custom_status, "Event_comment": refuse_reason,
                       "Event_date": decision_date, "decision_code": decision_code}
        parcel_list.append(parcel_info)
    list_chunks = list(chunks(parcel_list, 1000))
    print(list_chunks)
    i = 0
    for chunk in list_chunks:
        start = time.time()
        i += 1
        print(f'chunk {i}')
        #print(chunk)
        # 127.0.0.1:5000 # 164.132.182.145:5000
        response = requests.post('http://164.132.182.145:5000/api/add/new_event_chunks2', json=chunk, headers={'accept': 'application/json'})
        print(response.text)
        end = time.time()
        print("The time of execution of above program is :",
              (end - start), "s")
    msg = "ЗАГРУЖЕНО!"
    mb.showinfo("Информация", msg)

#insert_event_API_chunks2()


login = 'sl_api'
password = 'v3wMuaEeV64'

def authorization():
    url = 'https://mdt-express.deklarant.ru/api/Account/Login_V2'
    body = {
        'Login': login,
        'Password': password
    }
    header = {'Content-Type': 'application/json; charset=UTF-8'}
    print(body)
    respons = requests.post(url, json=body, headers=header)
    print(respons.status_code)
    print(respons.json())
    res_json = respons.json()
    Session = res_json['Content']['Session']
    print(Session)
    return Session


def Monitoring_events_toxl():
    filename = filedialog.askopenfilename()
    start = time.time()
    map_status = {'10': 'Выпуск товаров без уплаты таможенных платежей',
                         '30': 'Выпуск возвращаемых товаров разрешен',
                        '31': 'требуется уплата таможенных платежей',
                         '32': 'Выпуск товаров разрешен, таможенные платежи уплачsены',
                         '33': 'Выпуск разрешён, ожидание по временному ввозу',
                      '40': 'разрешение на отзыв',
                      '70': 'продление срока выпуска',
                      '90': 'отказ в выпуске товаров',
                         '0': 'статус не определен'}
    df = pd.read_excel(filename, sheet_name=0, engine='openpyxl', usecols='AO')
    print(df)
    n = 0
    parcels = []
    reg_numbers = []
    statuses = []
    reasonMessages = []
    Session = authorization()
    print(df)
    parcel_list = df['Номер накладной СДЭК'].tolist()
    #parcel_list = '["CEL2005611148CD", "CEL2005609168CD", "CEL2005606386CD"]'
    print(parcel_list)
    url = 'https://mdt-express.deklarant.ru/api/parcel/GetDecisionInfo'
    headers = {'login': 'sl_api',
                'sessionId': Session,
                'isMobileUser': 'false',
                'Content-Type': 'application/json'}
    list_chunks = list(chunks(parcel_list, 3000))
    df_result_all = pd.DataFrame()
    for chunk in list_chunks:
        # try:
            print(len(chunk))
            print()
            response = requests.post(url=url, headers=headers, json=chunk)
            result_json = response.text
            print(response.status_code)
            print(result_json)
            result = json.loads(result_json)
            df_result = pd.json_normalize(result, 'Decisions', ['Waybill'])
            df_result['Decision'] = df_result['DecisionCode'].map(map_status)
            # result = [parcels, reg_numbers, statuses, reasonMessages]
            # print(result)
            # df = pd.DataFrame(result).transpose()
            # print(df)
            end = time.time()
            print("The time of execution of above program is :",
                  (end - start), "s")
            df_result['Рег. номер'] = df_result['RegistrationNumber']
            df_result['Номер общей накладной'] = 'Unknown'
            df_result['Трек-номер'] = df_result['Waybill']
            df_result['Вес брутто'] = 0.1
            df_result['Пломба'] = None
            df_result['Статус ТО'] = df_result['Decision']
            df_result['Дата решения'] = df_result['DecisionDate']
            df_result['Дата решения'] = pd.to_datetime(df_result['Дата решения'])
            df_result['Дата решения'] = df_result['Дата решения'].dt.strftime('%d.%m.%Y %H:%M')
            df_result['Причина отказа ТО'] = None
            df_result['Код причины отказа'] = None
            df_result = df_result[['Рег. номер', 'Номер общей накладной', 'Трек-номер', 'Вес брутто', 'Пломба'
                                   , 'Статус ТО', 'Дата решения', 'Причина отказа ТО', 'Код причины отказа']]
            df_result_all = df_result_all.append(df_result)
        # except Exception as e:
        #     print(e)
    df_result_all = df_result_all.dropna(axis=0, how='any', subset='Статус ТО', inplace=False)
    writer = pd.ExcelWriter(f'{filename} - РЕШЕНИЯ.xlsx', engine='xlsxwriter')
    df_result_all.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.close()
    print(df_result_all)
    msg = "ВЫГРУЖЕНО!"
    mb.showinfo("Информация", msg)

window = tk.Tk()
window.title('Выгрузка и загрузка решений')
window.geometry("500x250+400+400")
name = tk.Label(window, text="Из мониторинга")

button = tk.Button(text="Выгрузить решения с Мониторинга", width=40, height=2, bg="lightgrey", fg="black", command=Monitoring_events_toxl)
button.configure(font=('hank', 10))
button2 = tk.Button(text="Загрузить решения на сервер", width=40, height=2, bg="lightgrey", fg="black", command=insert_event_API_chunks2)
button2.configure(font=('hank', 10))

name.pack()
button.pack()
button2.pack()
window.mainloop()


