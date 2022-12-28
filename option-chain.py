import xlwings as xw
import requests
import json
from nsetools import Nse
from time import sleep

nse = Nse()

url_oc = "https://www.nseindia.com/option-chain"
url_eq = "https://www.nseindia.com/api/option-chain-equities?symbol="

headers = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.149 Safari/537.36',
            'accept-language': 'en,gu;q=0.9,hi;q=0.8',
            'accept-encoding': 'gzip, deflate, br'}

sess = requests.Session()
cookies = dict()

def get_cookies():
    global cookies
    r = sess.get(url_oc, headers=headers)
    cookies = r.cookies.get_dict()

def get_option_chain(symbol):
    get_cookies()
    global cookies
    r = sess.get(url_eq + symbol, headers=headers, cookies=cookies)
    return json.loads(r.text)

with xw.App() as app:
    row = 1
    book = app.books.open('option-chain.xlsx')
    while True:
        while (book.sheets[0].range('A' + str(row)).value != None):
            symbol = book.sheets[0].range('A' + str(row)).value
            strike_price = book.sheets[0].range('B' + str(row)).value
            expiry_date = book.sheets[0].range('C' + str(row)).value
            try:
                expiry_date = expiry_date.strftime('%d-%b-%Y')
                stock_price = nse.get_quote(symbol)['lastPrice']
                print(symbol, strike_price, expiry_date)
                opt_chain = get_option_chain(symbol)
                for opt in opt_chain['records']['data']:
                    if (opt['expiryDate'] == expiry_date and float(opt['strikePrice']) == float(strike_price)):
                        print(opt)
                        book.sheets[1].range('A' + str(row+1)).value = symbol
                        book.sheets[1].range('B' + str(row+1)).value = strike_price
                        book.sheets[1].range('C' + str(row+1)).value = expiry_date
                        book.sheets[1].range('D' + str(row+1)).value = stock_price
                        book.sheets[1].range('E' + str(row+1)).value = opt['CE']['impliedVolatility']
                        book.sheets[1].range('F' + str(row+1)).value = opt['CE']['lastPrice']
            except:
                print("Error fetching data for " + symbol + " " + str(strike_price) + " " + str(expiry_date))
            row += 1
        book.sheets[1].range('A1').value = 'Symbol'
        book.sheets[1].range('B1').value = 'Strike Price'
        book.sheets[1].range('C1').value = 'Expiry Date'
        book.sheets[1].range('D1').value = 'Stock Price'
        book.sheets[1].range('E1').value = 'IV'
        book.sheets[1].range('F1').value = 'Last Price'
        book.save()
        sleep(5)