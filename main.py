import requests, openpyxl
import pandas as pd
from bs4 import BeautifulSoup

cur_usd = 'USD/RUB'
cur_eur = 'EUR/RUB'
mom_st = '2021-05-01'
mom_end = '2021-06-01'

url1 = 'https://www.moex.com/export/derivatives/currency-rate.aspx?language=ru' \
       '&currency={cur}&moment_start={mom_st}&moment_end={mom_end}'

url_usd = url1.format(cur=cur_usd, mom_st=mom_st, mom_end=mom_end)
url_eur = url1.format(cur=cur_eur, mom_st=mom_st, mom_end=mom_end)


def parsing(url):
    res = requests.get(url)
    html_str = str(res.text)
    soup = BeautifulSoup(html_str, 'lxml')

    rates = soup.find_all('rate')

    change = []
    for rate in rates:
        y = rate.attrs
        print(y['moment'], y['value'])
        change.append(y['value'])

    for sum in range(len(change)-1):
        num1, num2 = float(change[sum]), float(change[sum+1])
        print(num2-num1)


parsing(url_usd)
parsing(url_eur)
