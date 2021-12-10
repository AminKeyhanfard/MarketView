import requests
from decimal import Decimal
import xlwings as xw
import sched
import time

s = sched.scheduler(time.time, time.sleep)
wb = xw.Book('file.xlsx')
worksheet = wb.sheets('Tab1')

url = 'https://wallex.ir/api/v2/markets'
payload = ""
headers = {}
cart_tmn = ['ETH-TMN', 'XLM-TMN', 'EOS-TMN', 'BCH-TMN',
            'TRX-TMN', 'DOGE-TMN', 'SHIB-TMN', 'LTC-TMN']
cart_usdt = ['ETH-USDT', 'XLM-USDT', 'EOS-USDT', 'BCH-USDT',
             'TRX-USDT', 'DOGE-USDT', 'SHIB-USDT', 'LTC-USDT']

def refresh_api():
    response = requests.request("GET", url, headers=headers, data=payload).json()
    r = 1
    for item in cart_tmn:
        worksheet.range('F'+str(r+1)).value = float(Decimal(
            response['result']['symbols'][item]['stats']['bidPrice']).normalize())
        r += 1
    r = 1
    for item in cart_usdt:
        worksheet.range('H'+str(r+1)).value = float(Decimal(
            response['result']['symbols'][item]['stats']['bidPrice']).normalize())
        r += 1
    s.enter(5, 1, refresh_api)

print("Refreshing...")
s.enter(5, 1, refresh_api)
s.run()
