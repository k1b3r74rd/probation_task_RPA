import requests, openpyxl
import pandas as pd
from bs4 import BeautifulSoup

cur_usd = 'USD/RUB'
cur_eur = 'EUR/RUB'
mom_st = '2021-05-01'
mom_end = '2021-06-01'

url_pattern = 'https://www.moex.com/export/derivatives/currency-rate.aspx?language=ru' \
              '&currency={cur}&moment_start={mom_st}&moment_end={mom_end}'
url_usd = url_pattern.format(cur=cur_usd, mom_st=mom_st, mom_end=mom_end)
url_eur = url_pattern.format(cur=cur_eur, mom_st=mom_st, mom_end=mom_end)


class Currency:
    def link(cur):
        if cur == 'usd':
            return url_usd
        elif cur == 'eur':
            return url_eur

    def column_moment(cur):
        if cur == 'usd':
            return 1
        elif cur == 'eur':
            return 4

    def column_value(cur):
        if cur == 'usd':
            return 2
        elif cur == 'eur':
            return 5

    def column_dif(cur):
        if cur == 'usd':
            return 3
        elif cur == 'eur':
            return 6


wb = openpyxl.Workbook()
sheet = wb['Sheet']
sheet.cell(row=1, column=1).value = 'Дата Доллар'
sheet.cell(row=1, column=2).value = 'Курс Доллар'
sheet.cell(row=1, column=3).value = 'Изменение Доллар'
sheet.cell(row=1, column=4).value = 'Дата Евро'
sheet.cell(row=1, column=5).value = 'Курс Евро'
sheet.cell(row=1, column=6).value = 'Изменение Евро'


def excel_filler(cur):
    res = requests.get(Currency.link(cur))
    html_str = str(res.text)
    soup = BeautifulSoup(html_str, 'lxml')

    rates = soup.find_all('rate')

    values = []
    for row, rate in enumerate(rates):
        data = rate.attrs
        cell_moment = sheet.cell(row=row + 2, column=Currency.column_moment(cur))
        cell_value = sheet.cell(row=row + 2, column=Currency.column_value(cur))
        cell_moment.value = data['moment']
        cell_value.value = float(data['value'])
        values.append(data['value'])

    for position in range(len(values)-1):
        value1, value2 = float(values[position+1]), float(values[position])
        cell_dif = sheet.cell(row=position+2, column=Currency.column_dif(cur))
        cell_dif.value = value2 - value1


excel_filler('usd')
excel_filler('eur')

wb.save('test.xlsx')