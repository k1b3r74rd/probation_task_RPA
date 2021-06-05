# !/usr/bin/python3
# -*- coding: utf-8 -*-

import requests, openpyxl, smtplib, configparser
from bs4 import BeautifulSoup
from email.message import EmailMessage


# Данные и значения переменных из конфиг-файла settings.ini
class Config:
    config = configparser.ConfigParser()
    config.read('settings.ini')

    cur_usd = config.get('Link_params', 'cur_usd')
    cur_eur = config.get('Link_params', 'cur_eur')
    mom_st = config.get('Link_params', 'mom_st')
    mom_end = config.get('Link_params', 'mom_end')

    addr_from = config.get('Email_params', 'addr_from')
    password = config.get('Email_params', 'password')
    addr_to = config.get('Email_params', 'addr_to')

    # Нет в файле
    url_pattern = 'https://www.moex.com/export/derivatives/currency-rate.aspx?' \
                  'language=ru&currency={cur}&moment_start={mom_st}&moment_end={mom_end}'


# Данные для правильной расстановки данных по столбикам.
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

    for position in range(len(values) - 1):
        value1, value2 = float(values[position + 1]), float(values[position])
        cell_change = sheet.cell(row=position + 2, column=Currency.column_dif(cur))
        cell_change.value = value2 - value1
        cell_change.number_format = format_cur_ruble


# Деление курса евро на доллар и запись результата в столбик 'G'.
def eur_on_dollar():
    for pos in range(2, sheet.max_row + 1):
        eur_value = float(sheet.cell(row=pos, column=5).value)
        usd_value = float(sheet.cell(row=pos, column=2).value)
        sheet.cell(row=pos, column=7).value = eur_value / usd_value


# Отправка письма на почту.
def send_email(addr_to, msg_subj, msg_text, excel_file):
    msg = EmailMessage()
    msg['From'] = Config.addr_from
    msg['To'] = addr_to
    msg['Subject'] = msg_subj

    msg.set_content(msg_text)

    with open(excel_file, 'rb') as f:
        file_data = f.read()
    msg.add_attachment(file_data, maintype="application", subtype="xlsx", filename=excel_file)

    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(Config.addr_from, Config.password)
        server.send_message(msg)
        server.quit()


# Склонение слова "строка" после числа.
def declination(number: int, titles: list):
    cases = [ 2, 0, 1, 1, 1, 2 ]
    if 4 < number % 100 < 20:
        idx = 2
    elif number % 10 < 5:
        idx = cases[number % 10]
    else:
        idx = cases[5]

    return str(number) + ' ' +  titles[idx] + '.'


if __name__ == '__main__':
    url_usd = Config.url_pattern.format(cur=Config.cur_usd, mom_st=Config.mom_st, mom_end=Config.mom_end)
    url_eur = Config.url_pattern.format(cur=Config.cur_eur, mom_st=Config.mom_st, mom_end=Config.mom_end)

    wb = openpyxl.load_workbook('pattern.xlsx')  # pattern.xlsx - декоративный шаблон
    sheet = wb.active
    excel_filler('usd')
    excel_filler('eur')
    eur_on_dollar()
    rows_number = int(sheet.max_row)
    wb.save('Динамика курса за прошлый месяц.xlsx')

    excel_file = 'Динамика курса за прошлый месяц.xlsx'
    msg_subj = "Динамика курса с " + Config.mom_st + " по " + Config.mom_end
    msg_text = declination(rows_number, ['строка','строки','строк'])
    send_email(Config.addr_to, msg_subj, msg_text, excel_file)
