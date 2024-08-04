import hashlib
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


client_id_agreg = 'a67f70f2-0cc8-4cbf-9385-1dfd0f52c055'
client_secret_agreg = 'jNKpOzYdJHAk'

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
    token = json.loads(get_token_agreg(client_id_agreg, client_secret_agreg))
    token = token["Data"]
    print(token)
    headers = {'Authorization': f'Bearer {token}'}
    url = 'https://api-logistic-platform.ozonru.me/v1/TaxReport'
    body = {
        "TaxReports": [
                    {
                      "PostingNumber": "0251969948-0001-1",
                      "TrackingNumber": "HC000013517CD",
                      "TaxPayment": 2013.32,
                      "CustomsDuty": 500,
                      "Total": 2513.32,
                      "Currency": "RUB",
                      "InvoiceNumber": "211cb049-d6c0-414a-b504-af618629819d",
                      "DateOfPayment": "2024-07-25T11:47:34.822Z",
                      "RegisterNumber": "10716050/220724/П498487"
                    }
                  ]
                }


    response = requests.post(url=url, json=body,
                             headers=headers)
    print(response.text)
    print(response.status_code)


def send_pay_customs_info():

    url = 'http://164.132.182.145:5000/api/add/pay_customs_info'

    body = {
        "TaxReports": [
            {
                "PostingNumber": "0251969948-0001-1",
                "TrackingNumber": "HC000013517CD",
                "TaxPayment": 2013.32,
                "CustomsDuty": 500,
                "Total": 2513.32,
                "Currency": "RUB",
                "InvoiceNumber": "211cb049-d6c0-414a-b504-af618629819d",
                "DateOfPayment": "2024-07-25T11:47:34.822Z",
                "RegisterNumber": "10716050/220724/П498487",
                "provider": "OZON-AIR-260"
            }
        ]
    }

    response = requests.post(url=url, json=body)
    print(response.text)
    print(response.status_code)


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
    result = json.loads(response.text)
    print(result)
    df_result = pd.json_normalize(result, 'products', ['trackingNumber'])
    print(df_result)
    writer = pd.ExcelWriter('GOODS.xlsx', engine='xlsxwriter')
    df_result.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.close()

get_goods_info()


#send_pay_customs_info()

#get_parcels_info('0251969948-0001-1')


#registrate_postings()


#ozon_send_customs_pay_info()


#payresult_to_china()

#send_creating_pay_info()

#pay_result()