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
from imbox import Imbox
import datetime
import shutil
from zipfile import ZipFile
from pathlib import Path
import xml.etree.ElementTree as ET


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


insert_event_API_chunks()

