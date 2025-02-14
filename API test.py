import base64
import hashlib
import smtplib, ssl
from tkinter import filedialog
import time
import openpyxl
import requests
import json
from app import client
from passlib.hash import bcrypt
import pandas as pd
import datetime
from bs4 import BeautifulSoup
from pathlib import Path
import xml.etree.ElementTree as ET
import json


BASE = "http://185.233.200.168:5000/"

body_test = {"jsonrpc":"2.0","method":"HWB.Upload","params":{"Synch":True,"TotalWeight":439,"TotalVolume":0.0,"TotalCOD":0.0,"HWBQty":4,"ContractNumber":"8801519410","HWB":[{"OrderDate":"2024-08-21 13:59:45+00:00","HWBRefNumber":"AMLRU000000275YQ","DeclaredValue":228.0,"DeclaredCurrency":"RUB","LastMileInfo":{"Contractor":"IML last mile","ContractorICN":"1880157610"},"ServiceCode":"iec","ParcelQty":1,"SenderInfo":{"Name":"Shenzhen Dehong Supply Chain Co., Ltd.","Phone":"8615071940701","mobilePhone":"8615071940701","Email":"1156629561@qq.com","SenderAddress":{"CountryCode":"CN","City":"Dongguan","PostCode":"523721","Street":"China, Dongguan, tangxiazhenshahuyilu3hao, 3","Building":"","Company":"SHENZHEN DEHONG SUPPLY CHAIN CO., LTD."}},"ReceiverInfo":{"Name":"Кузина Дарья Александровна","Phone":"","MobilePhone":"79688261128","Email":"dakuzina@ozon.ru","PersonalData":{"FullName":"Кузина Дарья Александровна","IDIssueDate":"2018-05-05","IDNumber":"4518 563626","IdentityCardGiven":"ГУ МВД РОССИИ ПО Г.МОСКВЕ","TaxNumber":"773321667518","BirthDate":"1998-04-04","Files":[]},"ReceiverAddress":{"CountryCode":"RU","City":"Санкт-Петербург","Street":"Россия, 197229, г. Санкт-Петербург, просп. Юнтоловский, д. 51 корп. 1 стр. 1","PostCode":"197228"}},"Parcels":[{"Weight":108,"ParcelNo":1,"Items":[{"Quantity":1,"UnitValue":228.0,"Description":"Smartphone Case","DetailedDescription":"","URL":"https://www.ozon.ru/product/1664679332/","HTSCode":"","UnitWeight":150}]}],"ContractNumber":"8801519410","HWBWeight":108,"GoodsCurrency":"RUB"},{"OrderDate":"2024-08-21 13:58:31+00:00","HWBRefNumber":"AMLRU000000316YQ","DeclaredValue":228.0,"DeclaredCurrency":"RUB","LastMileInfo":{"Contractor":"IML last mile","ContractorICN":"1880157610"},"ServiceCode":"iec","ParcelQty":1,"SenderInfo":{"Name":"Shenzhen Dehong Supply Chain Co., Ltd.","Phone":"8615071940701","mobilePhone":"8615071940701","Email":"1156629561@qq.com","SenderAddress":{"CountryCode":"CN","City":"Dongguan","PostCode":"523721","Street":"China, Dongguan, tangxiazhenshahuyilu3hao, 3","Building":"","Company":"SHENZHEN DEHONG SUPPLY CHAIN CO., LTD."}},"ReceiverInfo":{"Name":"Кузина Дарья Александровна","Phone":"","MobilePhone":"79688261128","Email":"dakuzina@ozon.ru","PersonalData":{"FullName":"Кузина Дарья Александровна","IDIssueDate":"2018-05-05","IDNumber":"4518 563626","IdentityCardGiven":"ГУ МВД РОССИИ ПО Г.МОСКВЕ","TaxNumber":"773321667518","BirthDate":"1998-04-04","Files":[]},"ReceiverAddress":{"CountryCode":"RU","City":"Санкт-Петербург","Street":"Россия, 197229, г. Санкт-Петербург, просп. Юнтоловский, д. 51 корп. 1 стр. 1","PostCode":"197228"}},"Parcels":[{"Weight":123,"ParcelNo":1,"Items":[{"Quantity":1,"UnitValue":228.0,"Description":"Smartphone Case","DetailedDescription":"","URL":"https://www.ozon.ru/product/1664679332/","HTSCode":"","UnitWeight":150}]}],"ContractNumber":"8801519410","HWBWeight":123,"GoodsCurrency":"RUB"},{"OrderDate":"2024-08-21 14:00:12+00:00","HWBRefNumber":"AMLRU000000055YQ","DeclaredValue":228.0,"DeclaredCurrency":"RUB","LastMileInfo":{"Contractor":"IML last mile","ContractorICN":"1880157610"},"ServiceCode":"iec","ParcelQty":1,"SenderInfo":{"Name":"Shenzhen Dehong Supply Chain Co., Ltd.","Phone":"8615071940701","mobilePhone":"8615071940701","Email":"1156629561@qq.com","SenderAddress":{"CountryCode":"CN","City":"Dongguan","PostCode":"523721","Street":"China, Dongguan, tangxiazhenshahuyilu3hao, 3","Building":"","Company":"SHENZHEN DEHONG SUPPLY CHAIN CO., LTD."}},"ReceiverInfo":{"Name":"Кузина Дарья Александровна","Phone":"","MobilePhone":"79688261128","Email":"dakuzina@ozon.ru","PersonalData":{"FullName":"Кузина Дарья Александровна","IDIssueDate":"2018-05-05","IDNumber":"4518 563626","IdentityCardGiven":"ГУ МВД РОССИИ ПО Г.МОСКВЕ","TaxNumber":"773321667518","BirthDate":"1998-04-04","Files":[]},"ReceiverAddress":{"CountryCode":"RU","City":"Санкт-Петербург","Street":"Россия, 197229, г. Санкт-Петербург, просп. Юнтоловский, д. 51 корп. 1 стр. 1","PostCode":"197228"}},"Parcels":[{"Weight":104,"ParcelNo":1,"Items":[{"Quantity":1,"UnitValue":228.0,"Description":"Smartphone Case","DetailedDescription":"","URL":"https://www.ozon.ru/product/1664679332/","HTSCode":"","UnitWeight":150}]}],"ContractNumber":"8801519410","HWBWeight":104,"GoodsCurrency":"RUB"},{"OrderDate":"2024-08-21 13:40:27+00:00","HWBRefNumber":"AMLRU000000026YQ","DeclaredValue":228.0,"DeclaredCurrency":"RUB","LastMileInfo":{"Contractor":"IML last mile","ContractorICN":"1880157610"},"ServiceCode":"iec","ParcelQty":1,"SenderInfo":{"Name":"Shenzhen Dehong Supply Chain Co., Ltd.","Phone":"8615071940701","mobilePhone":"8615071940701","Email":"1156629561@qq.com","SenderAddress":{"CountryCode":"CN","City":"Dongguan","PostCode":"523721","Street":"China, Dongguan, tangxiazhenshahuyilu3hao, 3","Building":"","Company":"SHENZHEN DEHONG SUPPLY CHAIN CO., LTD."}},"ReceiverInfo":{"Name":"Кузина Дарья Александровна","Phone":"","MobilePhone":"79688261128","Email":"dakuzina@ozon.ru","PersonalData":{"FullName":"Кузина Дарья Александровна","IDIssueDate":"2018-05-05","IDNumber":"4518 563626","IdentityCardGiven":"ГУ МВД РОССИИ ПО Г.МОСКВЕ","TaxNumber":"773321667518","BirthDate":"1998-04-04","Files":[]},"ReceiverAddress":{"CountryCode":"RU","City":"Москва","Street":"г. Москва, наб. Пресненская, д. 10 блок C","PostCode":"123112"}},"Parcels":[{"Weight":104,"ParcelNo":1,"Items":[{"Quantity":1,"UnitValue":228.0,"Description":"Smartphone Case","DetailedDescription":"","URL":"https://www.ozon.ru/product/1664679332/","HTSCode":"","UnitWeight":150}]}],"ContractNumber":"8801519410","HWBWeight":104,"GoodsCurrency":"RUB"}]},"id":"cb0804c0-2c04-46bf-9641-e9fc82a7c68a"}

def chunks(lst, n):
    """Yield successive n-sized chunks from lst."""
    for i in range(0, len(lst), n):
        yield lst[i:i + n]


def insert_event_API_test():
    file_name = filedialog.askopenfilename()
    df = pd.read_excel(file_name)
    i = 0
    for parcel_numb in df["Трек-номер"]:
        i += 1
        print(parcel_numb)
        Event = df.loc[df["Трек-номер"] == parcel_numb]["Статус ТО"].values[0]
        Event_comment = df.loc[df["Трек-номер"] == parcel_numb]["Причина отказа ТО"].values[0]
        Event_date = df.loc[df["Трек-номер"] == parcel_numb]["Дата решения"].values[0]
        Event_date_edit = datetime.datetime.strptime(Event_date, "%d.%m.%Y %H:%M")
        Event_date_edit = datetime.datetime.strftime(Event_date_edit, "%Y-%m-%d %H:%M:%S")
        internal_event = "Принято решение"
        regnumber = df.loc[df["Трек-номер"] == parcel_numb]["Рег. номер"].values[0]
        body = {"parcel_numb": parcel_numb, "internal_event": internal_event, "regnumber": regnumber,
        "Event": f"{Event}", "Event_comment": f"{Event_comment}", "Event_date": f"{Event_date_edit}"}
        print(body)
        response = requests.post("http://164.132.182.145:5000/api/add/new_event", json=body, headers={"accept": "application/json"})   #"http://164.132.182.145:5000/api/add/new_event"
        print(response.text)
        print(i)

# "regnumber": regnumber  "internal_event": internal_event,   //164.132.182.145:5000

def get_parcel_info_API_test():
    body = {"parcel_numb": "CEL9000291373CD"}

    headers = {"accept": "application/json"}
    response = requests.post("http://127.0.0.1:5000/api/v1.0/events/", json=body) # http://127.0.0.1:5000  # "http://185.233.200.168:5000/api/v1.0/events/"
    print(type(response.json()))


def send_json_to_SVH():
    parcel_list = ["FS20234J157712718061",
                   "FS20234O155710356824"
                   ]
    for parcel in parcel_list:
        event_details = {"parcel_numb": parcel, "Event": "Выпуск ТЕСТ", "Event_comment": "т/п Уссурийский (10716050)", "Event_date": "2023-04-15 09:34:00"}
        response = requests.post("http://192.168.0.100:5001/api/add_decision", json=event_details,
                                 headers={"accept": "application/json"})
        try:
            return response.json()
        except ValueError:
            pass

def login_insert():
    log = client.post("/todo/api/v1.0/login", json={"email": "test2@gmail.com", "password": "password"})
    res = log.get_json()
    print(res)

def Django_insert_event_API_test():
    file_name = filedialog.askopenfilename()
    df = pd.read_excel(file_name)
    i = 0
    for parcel_numb in df["Трек-номер"]:
        i += 1
        print(parcel_numb)
        Event = df.loc[df["Трек-номер"] == parcel_numb]["Статус ТО"].values[0]
        Event_comment = str(df.loc[df["Трек-номер"] == parcel_numb]["Причина отказа ТО"].values[0])
        if Event_comment == "nan":
            Event_comment = ""
        Event_date = df.loc[df["Трек-номер"] == parcel_numb]["Дата решения"].values[0]
        Event_date_edit = datetime.datetime.strptime(Event_date, "%d.%m.%Y %H:%M")
        Event_date_edit = datetime.datetime.strftime(Event_date_edit, "%Y-%m-%d %H:%M:%S")
        body = {"parcel_numb": parcel_numb, "status_name": Event, "comment": Event_comment,
                "time": Event_date_edit}
        response = requests.post("http://164.132.182.145:8000/api_insert_decisions/", json=body, headers={"accept": "application/json"})
        print(response.text)
        print(i)


def ozon_api_send_manifest_info():
    filename = filedialog.askopenfilename()
    wb = openpyxl.load_workbook(filename)
    ws = wb.active
    seal_number = ws["A1"].value
    print(seal_number)
    df = pd.read_excel(filename, sheet_name=0, skiprows=3, engine="openpyxl", usecols="A, B, C, D", header=None)
    df = df[df[1].notna()]
    df = df[df[3].str.contains("OZON")]
    posting_numbers = df[1].to_list()
    print(len(posting_numbers))
    url = "http://hccd.rtb56.com/webservice/Ozon/OzonRocketSealNumber.ashx"
    data = {"seal_number": seal_number, "posting_numbers": posting_numbers}
    # json_obj = json.dumps(data, separators=(", ", ": "))
    json_obj = str(data)
    print(len(json_obj))
    print(json_obj)
    m = hashlib.md5(json_obj.encode("utf-8")).hexdigest().upper()
    print(m)
    headers = {"sign": f"{m}"}
    print(headers)
    response = requests.post(headers=headers, url=url, data=json_obj)
    print(response.status_code, response.text)


def create_invoice(parcel_numb, receiver, pay_sum):
    Bearer_token = "Bearer 9e257ccdb5044bb48c2d559a050aa563"
    url = "https://api.intellectmoney.ru/merchant/createInvoice"
    delta = datetime.timedelta(days=+10)
    expired_date = datetime.datetime.now() + delta
    expireDate = expired_date.strftime("%Y-%m-%d %H:%M:%S")
    print(expireDate)
    eshopId = "467362"
    orderId = parcel_numb
    serviceName = ""
    recipientAmount = pay_sum
    recipientCurrency = "RUB"
    userName = receiver
    email = "info@cellog.ru"
    successUrl = ""
    failUrl = ""
    backUrl = ""
    resultUrl = ""
    holdMode = ""
    merchantReceipt = json.dumps({
        "inn": "2725106960",
        "group": "dee7b4d6-3f33-4923-b0b4-b2c1031d3071",
        "content":
            {
                "type": 1,
                "customerContact": email,

                "positions": [
                    {
                        "quantity": 1.000,
                        "price": pay_sum,
                        "tax": 6,
                        "text": "Оплата пошлины за экспресс груз",
                        "paymentSubjectType": 4,
                        "paymentMethodType": 4
                    }
                ],
                "checkClose":
                    {"payments": [
                        {
                            "type": 2,
                            "amount": pay_sum}],
                        "taxationSystem": 2
                    },
            }
    })
    preference = ""
    signSecretKey = "8cac752e75044bd081ad1817bdbbf457"
    shop_secret_key = "LSrB53*Ds8#2j%%C"
    list_params = [eshopId, orderId,
                   serviceName, recipientAmount,
                   recipientCurrency, userName,
                   email, successUrl,
                   failUrl, backUrl,
                   resultUrl, expireDate,
                   holdMode, preference]
    print(list_params)
    data = "::".join(list_params)
    data_param_hash = data + "::" + shop_secret_key
    m = hashlib.md5(data_param_hash.encode("utf-8")).hexdigest()
    print(m)
    params = {"eshopId": eshopId,
                "orderId": orderId,
                "serviceName": "",
                "recipientAmount": recipientAmount,
                "recipientCurrency": recipientCurrency,
              "userName": userName,
              "email": email,
              "successUrl": "",
              "failUrl": "",
              "backUrl": "",
              "resultUrl": "",
              "expireDate": expireDate,
              "holdMode": "",
              'merchantReceipt': merchantReceipt,
              "preference": "",
                "hash": m,
                }
    data_sign_hash = data + "::" + signSecretKey
    print(data_sign_hash)
    sign = hashlib.sha256(data_sign_hash.encode("UTF-8")).hexdigest()
    print(sign)
    headers = {"Authorization": Bearer_token,
               "Sign": sign}
    response = requests.post(url=url, data=params, headers=headers)
    print(response.status_code)
    html_resp = response.text
    print(html_resp)
    soup = BeautifulSoup(html_resp, "xml")
    invoice_id = str(soup.findAll("InvoiceId")).replace("[<InvoiceId>", "")
    invoice_id = invoice_id.replace("</InvoiceId>]", "")
    print(invoice_id)
    return invoice_id


def invoice_from_excel():
    file_name = filedialog.askopenfilename()
    df = pd.read_excel(file_name, usecols="A, K, Q, R")

    def toFixed(numObj, digits=0):
        return f"{numObj:.{digits}f}"
    dict_result = {}
    list_parcels = []
    list_telephone = []
    list_url_send = []
    for row in df.iterrows():
        parcel_numb = row[1]["Трек-номер"]
        receiver = row[1]["Получатель"]
        pay_sum = float(row[1]["Сумма сборов и пошлины"].replace(",", "."))
        toFixed(pay_sum, 2)
        pay_sum = str(pay_sum)
        print(pay_sum)
        telephone = row[1]["Телефон"]
        invoice_id = create_invoice(parcel_numb, receiver, pay_sum)
        url_send = f"https://merchant.intellectmoney.ru/v2/ru/process/{invoice_id}/acquiring"
        list_parcels.append(parcel_numb)
        list_telephone.append(telephone)
        list_url_send.append(url_send)
    dict_result["Номер отправления"] = list_parcels
    dict_result["Номер телефона"] = list_telephone
    dict_result["ССЫЛКА"] = list_url_send
    print(dict_result)
    df_dict_result = pd.DataFrame(dict_result)
    file_name_only = Path(file_name).stem
    writer = pd.ExcelWriter(f"ФТС РАССЫЛКА {file_name_only}.xlsx", engine="xlsxwriter")
    df_dict_result.to_excel(writer, sheet_name="Sheet1", index=False)
    writer.save()

def pay_result():
    json = {"eshopId": 467362,
            "paymentId": 3872927931,
            "orderId": "CEL6000502158CD",
            "eshopAccount": 6775290743,
            "serviceName": "",
            "recipientAmount": 1312.96,
            "commissionAmount": 31.51,
            "recipientOriginalAmount": 1312.96,
            "recipientCurrency": "RUB",
            "paymentStatus": '5',
            "userName": "Корчагин Владимир Юрьевич",
            "userEmail": "info@cellog.ru",
            "paymentData": "2024-05-06 14:45:10",
            "payMethod": "Acquiring",
            "gatewayName": "GpBankDirect",
            "hash": "b1f7701b320cb1426d91260db6b9f8cb",
            "rrn": "412711124782 "}

    response = requests.post("http://164.132.182.145:5000/api/payresult", json=json,
                             headers={"accept": "application/json"})
    print(response.text)
    return True

#grg
def payresult_to_china():
    file_name = filedialog.askopenfilename()
    df = pd.read_excel(file_name, usecols=[0])
    print(df)
    i = 0
    for parcel_numb in df['Заказ']:

        url = "http://hccd.rtb56.com/webservice/Ozon/OzonUpdatePayTaxData.ashx"
        data = [
            {
                "posting_number": "",
                "tracking_number": parcel_numb,
                "is_paid": f"Y"
            }
        ]
        i += 1
        print(i)
        response = requests.post(url=url, json=data)
        print(response.text)

#invoice_from_excel()
#create_invoice("test35", 'pupkin', '1')
#insert_event_API_test()
#ozon_api_send_manifest_info()
#pay_result()
#payresult_to_china()


def invoice_status_to_china(parcel_numb, invoice_id, pay_sum):
    delta = datetime.timedelta(days=+10)
    expired_date = datetime.datetime.now() + delta
    expired_date = expired_date.strftime("%Y-%m-%d")
    print(expired_date)
    url = "http://hccd.rtb56.com/webservice/Ozon/OzonSavePayTaxData.ashx"
    data = [
        {
            "posting_number": "",
            "tracking_number": parcel_numb,
            "pay_tax_end_time": expired_date,
            "pay_tax_link": f"https://merchant.intellectmoney.ru/v2/ru/process/{invoice_id}/acquiring",
            "tax_amount": pay_sum,
            "is_paid": f"N"
        }
    ]
    response = requests.post(url=url, json=data)
    print(response.text)


def get_user_token():
    url = 'https://api.intellectmoney.ru/personal/user/getUserToken'
    Login = 'Programmist'
    Password = '4W86L&d)xe9)'
    params = {"Login": Login, "Password": Password}
    response = requests.post(url=url, params=params)
    resp_html = response.text
    print(resp_html)
    root = ET.fromstring(resp_html)
    UserToken = root.findall("./Result/UserToken")[0].text
    print(UserToken)
    return UserToken

#get_user_token()


def cancel_invoice(invoice_id, UserSid):
    url = 'https://lk.intellectmoney.ru/api/payment/CancelInvoice'
    json_str = {"ObjectId": f"{invoice_id}", "UserSid": UserSid}
    try:
        response = requests.post(url=url, json=json_str)
        print(response.text)
    except:
        time.sleep(7)
        response = requests.post(url=url, json=json_str)
        print(response.text)
    resp = response.text
    try:
        success = json.loads(resp)['Result']['Data']['Success']
        if success is True:
            pass
        else:
            UserSid = get_user_token()
            json_str = {"ObjectId": f"{invoice_id}", "UserSid": UserSid}
            response = requests.post(url=url, json=json_str)
            print(response.text)
    except Exception as e:
        print(e)
        time.sleep(7)
        UserSid = get_user_token()
        json_str = {"ObjectId": f"{invoice_id}", "UserSid": UserSid}
        response = requests.post(url=url, json=json_str)
        print(response.text)


def cancel_excel():
    file_name = filedialog.askopenfilename()
    df = pd.read_excel(file_name, usecols="B")
    print(df)
    i = 0
    UserSid = get_user_token()
    for invoice in df['СКО']:
        i += 1
        print(i)
        print(invoice)

        cancel_invoice(invoice, UserSid)


def send_creating_pay_info():
    file_name = filedialog.askopenfilename()
    df = pd.read_excel(file_name)
    print(df)
    for parcel in df['Номер отправления']:
        parcel_numb = parcel
        phone = df.loc[df['Номер отправления'] == parcel]['Номер телефона'].values[0]
        pay_sum = df.loc[df['Номер отправления'] == parcel]['Сумма с комиссией'].values[0].astype(str).replace(',', '.')
        print(parcel_numb, phone, pay_sum)
        json = {
                "parcel_numb": parcel_numb,
                "pay_sum": str(pay_sum),
                "phone": str(phone),
                }

        response = requests.post("http://164.132.182.145:5000/api/add/creating_pay_info", json=json,
                                 headers={"accept": "application/json"})
        print(response.text)
    #return True


def send_customs_pay_info():
    json = {
        "TaxReports": [
                    {
                      "PostingNumber": "string",
                      "TrackingNumber": "string",
                      "TaxPayment": 0,
                      "CustomsDuty": 0,
                      "Total": 0,
                      "Currency": "string",
                      "InvoiceNumber": "string",
                      "DateOfPayment": "2024-07-08T11:47:34.822Z",
                      "RegisterNumber": "string",
                      "Provider ": "ETS-AIR"
                    }
                  ]
                }


    response = requests.post("http://164.132.182.145:5000/api/add/pay_customs_info", json=json,
                             headers={"accept": "application/json"})
    print(response.text)


client_id_agreg = '78d96cd9-909c-4e17-ac73-1b6669cbbd43'
client_secret_agreg = 'fyfJmgyftyYB'

client_id = 'a67f70f2-0cc8-4cbf-9385-1dfd0f52c055'
client_secret = 'jNKpOzYdJHAk'


def get_token_rocket(client_id, client_secret):
    url = 'https://xapi.ozon.ru/principal-auth-api/connect/token'
    headers = {'content-type': 'application/x-www-form-urlencoded'}
    data = f'grant_type=client_credentials&client_id={client_id}&client_secret={client_secret}'
    response = requests.post(url=url, data=data, headers=headers)
    token = response.text
    print(token)
    return token


#get_token_rocket(client_id, client_secret)


def get_token_agreg(client_id_agreg, client_secret_agreg):
    url = 'https://api-logistic-platform.ozonru.me/GetAuthToken'
    headers = {'content-type': 'application/json'}
    data = {
    "ClientId": client_id_agreg,
    "ClientSecret": client_secret_agreg
    }
    response = requests.post(url=url, json=data, headers=headers)
    token = response.text
    print(token)
    return token

#get_token_agreg(client_id_agreg, client_secret_agreg)

def ozon_send_customs_pay_info():
    # # Таможенный портал
    # client_id_agreg = 'ddd0762f-ec85-4ab5-9974-4fef92af541b'
    # client_secret_agreg = 'KyhgEttYXnYa'
    # token = json.loads(get_token_agreg(client_id_agreg, client_secret_agreg))
    # token = token["Data"]
    # print(token)
    # headers = {'Authorization': f'Bearer {token}'}
    url = 'http://62.109.0.39:5000/api/add/pay_customs_info'
    #url = 'http://192.168.1.237:5000/api/add/pay_customs_info'
    file_name = filedialog.askopenfilename()
    df = pd.read_excel(file_name)
    print(df)
    df['Дата решения'] = pd.to_datetime(df['Дата решения'], format='%d.%m.%Y %H:%M')
    for index, item in df.iterrows():
        PostingNumber = item['Номер отправления']
        TrackingNumber = item['Трек-номер']
        CustomsDuty = float(item['Таможенные сборы'].replace(',', '.'))
        TaxPayment = float(item['Таможенная пошлина'].replace(',', '.'))
        Total = float(item['Сумма сборов и пошлины'].replace(',', '.'))
        InvoiceNumber = item['WaybillID']
        DateOfPayment = item['Дата решения'].strftime('%Y-%m-%dT%H:%M:%S') + '+10:00'
        RegisterNumber = item['Рег. номер']
        provider = item['Примечание']
        body = {
            "TaxReports": [
                        {
                          "PostingNumber": PostingNumber,
                          "TrackingNumber": TrackingNumber,
                          "TaxPayment": TaxPayment,
                          "CustomsDuty": CustomsDuty,
                          "Total": Total,
                          "Currency": "RUB",
                          "InvoiceNumber": InvoiceNumber,
                          "DateOfPayment": DateOfPayment,
                          "RegisterNumber": RegisterNumber,
                            "Provider": provider
                        }
                      ]
                    }

        print(body)
        response = requests.post(url=url, json=body)
        print(response.text)
        print(response.status_code)


ozon_send_customs_pay_info()


def send_pay_customs_info():
    file_name = filedialog.askopenfilename()
    df = pd.read_excel(file_name, usecols='B')
    print(df)
    for index, item in df.iterrows():
        print(item['json'])
        json_pay = item['json'].replace("\'", "\"")
        json_pay = json.loads(json_pay)
        url = 'http://164.132.182.145:5000/api/add/pay_customs_info'
        # url = 'http://192.168.0.104:5000/api/add/pay_customs_info'
        body = {
        "TaxReports": [
        json_pay
            ]
            }

        print(body)

        response = requests.post(url=url, json=json_pay)
        print(response.text)
        print(response.status_code)


#send_pay_customs_info()

def send_pay_customs_wb():
    url = 'https://integrations.wb.ru/rupost-marketplace/external/api/v1/customs_duties/cel/save'
    headers = {
        "Authorization": "Bearer mcly5djawjb3ur0070q3t6c6465x8hoy81broovyu6w2pn99xke4v5gyb9p82f6t",
        "Content-Type": "application/json"
    }

    body = {
            "customs_duties": [
            {
            "register_number": "10716050/091124/П818468",
            "order_number": "WBCNRUCLDAJ20000BN",
            "customs_duty": 681.58,
            "tax_payment": 500,
            "total": 1181.570,
            "currency": "RUB",
            "date_of_payment": "2024-11-10T11:55:00+10:00"
            },
                {
                    "register_number": "10716050/091124/П821814",
                    "order_number": "WBCNRUCLDAJ290009J",
                    "customs_duty": 12400.090,
                    "tax_payment": 500,
                    "total": 12900.020,
                    "currency": "RUB",
                    "date_of_payment": "2024-11-10T11:55:00+10:00"
                },

            ]
            }


    print(body)

    response = requests.post(url=url, json=body, headers=headers)
    print(response.text)
    print(response.status_code)


#send_pay_customs_wb()


def get_goods_info():
    filename = filedialog.askopenfilename()
    df = pd.read_excel(filename, header=None)
    url = 'https://cellog.deklarant.ru/api/external/parcel-products'
    headers = {
                "api-token": "40e2f498-450c-4b9f-a509-7f4c8877a6ff",
               "Content-Type": "application/json"
               }
    body = []
    print(df)
    for parcel_numb in df[0]:
        body.append(parcel_numb)

    print(body)
    response = requests.post(url=url, headers=headers, json=body)
    print(response.text)
    print(response.status_code)
    result = json.loads(response.text)
    print(result)
    df_result = pd.json_normalize(result, 'products', ['trackingNumber'])
    print(df_result)
    writer = pd.ExcelWriter('GOODS.xlsx', engine='xlsxwriter')
    df_result.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.close()





def decod_xml_to_base64():
    # convert file content to base64 encoded string
    with open("216b8252-2eb1-4b0a-82a4-37178c17be5f.xml", "rb") as file:
        raw_file = file.read()
        print(raw_file)
        encoded = base64.encodebytes(raw_file).decode("utf-8")


    # output base64 content
    #print(encoded)


    #print(encoded)
    return encoded

def send_base64_ozon():
    encoded = decod_xml_to_base64()
    # token = json.loads(get_token_agreg(client_id_agreg, client_secret_agreg))
    # token = token["Data"]
    # #print(token)
    # headers = {'Authorization': f'Bearer {token}'}
    url = 'http://164.132.182.145:5000/api/TaxDocuments' # 'https://api-logistic-platform.ozon.ru/v1/TaxDocuments'

    #print(encoded)
    body_json = {'documentType': 'Cmn',
            'documentData': encoded}
    #body_json = json.dumps(body)
    #print(body)
    with open('json_xml.json', 'w', encoding='UTF-8') as fp:
        json.dump(body_json, fp, ensure_ascii=False)
    response = requests.post(url=url, json=body_json)
    print(body_json)
    print(response.status_code)
    print(response.text)

#send_base64_ozon()

def get_parcel_info_list():
    url = 'http://164.132.182.145:5000/todo/events/list_api'
    body_json = ['CEL1000474095CD']
    response = requests.post(url=url, json=body_json)
    print(response.status_code)
    print(response.text)


username = "cel-python-automatization@yandex.ru"
password = "vmpqeopkfptvejiz"


def send_email(body_text, subject):
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


def send_scarif_json():
    print('start')
    with open('parcel_pay_customs_notification_202411072212.json', 'r') as file:
        data = json.load(file)
        for payload in data['info']:
            print(payload['payload'])




def login_api():
    user_name = 'iml'
    password = 'H3m76Opq78'
    url = 'http://192.168.0.102:5000/todo/api/v1.0/login'

    response = requests.post(url=url, json={'email': user_name, 'password': password})
    print(response.text)


def login_register():
    user_name = 'iml'
    password = 'H3m76Opq78'
    url = 'http://192.168.0.102:5000/todo/api/v1.0/register'

    response = requests.post(url=url, json={'name': 'user_name', 'email': user_name, 'password': password})
    print(response.text)

def test_request_register():
    body = body_test
    url = 'http://62.109.0.39:5000/api/add/register'

    response = requests.post(url=url, json=body)
    print(response.text)


#get_goods_info()

ozon_send_customs_pay_info()

#test_request_register()

#login_register()

#login_api()


#send_scarif_json()


#send_email()


#get_parcel_info_list()

#send_base64_ozon()


#send_pay_customs_info()

#get_parcels_info('0251969948-0001-1')


#registrate_postings()




#payresult_to_china()

#send_creating_pay_info()

#pay_result()