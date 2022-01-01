#A program to scrap stock prices from various companies

import requests
from fake_useragent import UserAgent
from bs4 import BeautifulSoup
import pandas as pd
from pandas import ExcelWriter
from datetime import date
from openpyxl import load_workbook
import os

today = date.today()

#Create a soup object from a URL
def getSoup(url):
    ua = UserAgent()
    header = {'User-Agent':str(ua.random)}
    r = requests.get(url, headers=header)
    data = r.text
    soup=BeautifulSoup(data, features='lxml')
    return soup
    

def make_hyperlink(url,value):
    return '=HYPERLINK("%s", "%s")' % (url, value)


symbols = []
soup = getSoup('https://www.slickcharts.com/sp500')

#Find the desired table
table500 = soup.find('table',attrs={'class':'table table-hover table-borderless table-sm'})
tr500 = table500.findChildren('tbody',recursive=False)[0].findChildren('tr',recursive=False)
for tr in tr500:
    #Get the stock symbol
    symbol = tr.findChildren('td',recursive=False)[2].text
    symbols.append(symbol)
names = {}
for s in symbols:
    #Get stock prices
    daysclose = {}
    url = f'https://finance.yahoo.com/quote/{s}/history?p={s}'
    soup = getSoup(url)
    name = soup.find('h1', attrs={'class':'D(ib) Fz(18px)'}).text
    print(f'Scraping close values from: {name}')
    table = soup.find('table',attrs={'class':'W(100%) M(0)'})
    tbody = table.findChildren("tbody" , recursive=False)[0]
    dates = tbody.findChildren("tr" , recursive=False)
    for d in dates:
        try:
            if len(daysclose) >= 10:
                break
            day = d.findChildren('td')[0].text
            close = d.findChildren('td')[4].text
            daysclose[day] = close
        except:
            continue
    names[make_hyperlink(url,name)] = daysclose
    df = pd.DataFrame.from_dict(names).transpose()
df.dropna(inplace=True)
try:
    #Create an excel sheet
    ExcelWorkbook = load_workbook(os.path.dirname(os.path.abspath(__file__))+'\Stocks.xlsx')
    writer = ExcelWriter('Stocks.xlsx',engine = 'openpyxl')
    writer.book = ExcelWorkbook
except:
    writer = ExcelWriter('Stocks.xlsx')
df.to_excel(writer,f'{today}')
writer.save()
writer.close