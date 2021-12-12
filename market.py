import requests
from decimal import Decimal
import xlwings as xw
import sched
import time
import datetime
#from secrets import technical_analysis_api


s = sched.scheduler(time.time, time.sleep)
wb = xw.Book('Wallet.xlsx')
worksheet = wb.sheets('Tab1')


url_wallex = 'https://wallex.ir/api/v2/markets'
payload = ""
headers = {}
response_wallex = requests.request(
    "GET", url_wallex, headers=headers, data=payload).json()
USDT_TMN = float(Decimal(
    response_wallex['result']['symbols']['USDT-TMN']['stats']['askPrice']).normalize())
print('USDT : ', USDT_TMN, 'Toman')


def refresh_api():
    wallet_list = ['ethereum', 'stellar', 'eos', 'bitcoin-cash',
                   'tron', 'dogecoin', 'shiba-inu', 'litecoin']
    ids_string = []
    ids_string.append(','.join(wallet_list))
    api_url = f'https://api.coingecko.com/api/v3/coins/markets?vs_currency=usd&ids={ids_string[0]}&sparkline=false&order=id_asc'
    data = requests.get(api_url).json()
    print(datetime.datetime.now(), '  Package recieved! Updating...')
    r = 1
    for item in data:
        #        api_url_recommend = f"https://technical-analysis-api.com/api/v1/analysis/{item['symbol'].upper()}?apiKey={technical_analysis_api}"
        #        print(api_url_recommend)
        #        data_recommend = requests.get(api_url_recommend).json()
        #        print(data_recommend)
        # ,data_recommend['recommendation'])
        print(item['name'], ' : ', item['current_price'])
        worksheet.range('F'+str(r+1)).value = USDT_TMN*item['current_price']
        worksheet.range('H'+str(r+1)).value = float(Decimal(
            item['current_price']).normalize())
        r += 1
    print('Waiting for the next cycle...')
    s.enter(10, 1, refresh_api)


print("Starting...")
s.enter(0, 1, refresh_api)
s.run()
