# !/usr/bin/python3
# -*- coding: utf-8 -*-

import requests, openpyxl
from bs4 import BeautifulSoup

# Параметры для ссылки
cur_usd = 'USD/RUB'
cur_eur = 'EUR/RUB'
mom_st = '2021-05-01'
mom_end = '2021-06-01'

url_pattern = 'https://www.moex.com/export/derivatives/currency-rate.aspx?language=ru' \
              '&currency={cur}&moment_start={mom_st}&moment_end={mom_end}'

url_usd = url_pattern.format(cur=cur_usd, mom_st=mom_st, mom_end=mom_end)
url_eur = url_pattern.format(cur=cur_eur, mom_st=mom_st, mom_end=mom_end)


# Конфигурация для правильной расстановки данных по столбикам.
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


# Получение информации с сайта с последующей записью в excel.
def excel_filler(cur):
    res = requests.get(Currency.link(cur))
    html_str = str(res.text)
    soup = BeautifulSoup(html_str, 'lxml')

    rates = soup.find_all('rate')

    format_cur_ruble = '#,##0.00_-₽'
    values = []
    for row, rate in enumerate(rates):
        data = rate.attrs
        cell_moment = sheet.cell(row=row + 2, column=Currency.column_moment(cur))
        cell_value = sheet.cell(row=row + 2, column=Currency.column_value(cur))
        cell_moment.value = data['moment']
        cell_value.number_format = format_cur_ruble
        cell_value.value = float(data['value'])
        values.append(data['value'])

    for position in range(len(values)-1):
        value1, value2 = float(values[position+1]), float(values[position])
        cell_change = sheet.cell(row=position+2, column=Currency.column_dif(cur))
        cell_change.value = value2 - value1
        cell_change.number_format = format_cur_ruble


# Деление курса евро на доллар и запись результата в столбик 'G'.
def eur_on_dollar():
    for pos in range(2, sheet.max_row+1):
        eur_value = float(sheet.cell(row=pos, column=5).value)
        usd_value = float(sheet.cell(row=pos, column=2).value)
        sheet.cell(row=pos, column=7).value = eur_value / usd_value


if __name__ == '__main__':
    # pattern.xlsx - декоративный шаблон
    wb = openpyxl.load_workbook('pattern.xlsx')
    sheet = wb.active

    excel_filler('usd')
    excel_filler('eur')
    eur_on_dollar()

    wb.save('Динамика курса за прошлый месяц.xlsx')