# -*- coding: utf-8 -*-

import win32com.client as win32
import psutil
import os
import subprocess
import logging
import xlrd
import logging
import schedule

from time import sleep
from random import randint
from datetime import datetime

send_email_today = 0

logging.basicConfig(filename="errors.log", level=logging.INFO)

row_name = 0
row_mail = 1
row_date = 2
row_header = 3
row_body = 4

messages = [
"Прими самые добрые и искренние поздравления по случаю твоего дня рождения!\nС особой теплотой хотим сказать, что гордимся и дорожим сложившимися отношениями теплого сотрудничества и взаимопонимания.\nОт всей души желаем тебе крепкого здоровья, счастья, успехов в работе.\nПусть удача сопутствуют тебе и твои близким",

'Пусть и на работе, и в семье тебе сопутствует успех и благополучие. Желаем успешного осуществления всех твоих благих начинаний, а еще — оставаться таким же профессионалом своего дела и просто замечательным человеком!\nПусть удача сопутствуют тебе и твои близким'
]


def excel_birthdays(file_name):
    wb = xlrd.open_workbook(file_name)
    ws = wb.sheet_by_index(0)
    flag = 0
    list_of_adresses = []
    for rownum in range(ws.nrows)[1:]:
        excel_row = ws.row_values(rownum)
        for num, cell in enumerate(excel_row):
            if cell == "":
                pass
            else:
                if not flag:
                    flag = 1
                    break
                try:
                    birthday = datetime(*xlrd.xldate_as_tuple(excel_row[num + 2], wb.datemode)).strftime("%m-%d")
                except:
                    logging.error("wrong date format in {}, {} row, {} cell. Must be dd.mm.yyyy (например 21.06.1990)".format(file_name, rownum, num + 3))
                    break
                today_date = datetime.today().strftime("%m-%d")
                if birthday == today_date:
                    list_of_adresses.append(excel_row[num:])
                break
    return (list_of_adresses)


def mail_go_go(birthday_data, time_send, friends):

    outlook = win32.Dispatch("Outlook.Application")

    mails = friends[0]
    names = friends[1]
    print(mails, names)
    #footer = ""
    #for mail, name in zip(mails, names):
    #    footer += name + " (" + mail + ")\n"

    try:
        mail = outlook.CreateItem(0)
        mail.To = birthday_data[0]
        mail.CC = ";".join(mails)
        mail.Subject = 'С Днем Рождения!'
        #header = "Мы, " + copy_names + " поздравляем тебя с Днем Рождения!\n" if len(copy_names) > 1 else ""
        mail.body = birthday_data[2] + "\n" + birthday_data[3]
        mail.send
        logging.info(u"   {}: Send message to {}, copy for {}".format(time_send, birthday_data[0], ",".join(mails)))
    except Exception as e:
        logging.error(u"   {}: Error {}, while tried send message to {}".format(time_send, e, birthday_data[0]))


def get_friends_mails_and_names(birthday_boy, time_send, birthday_name, birthday_mail):

    friends_names = []
    friends_mails = []
    birthday_boy_size = len(birthday_boy)
  
    for friend_id in range(row_body + 1, birthday_boy_size, 2):

        if friend_id + 1 > birthday_boy_size:
            logging.error("   {}: Error {}, именинник {} {}".format(time_send, "неверные данные у {}".format(birthday_boy[friend_id]), birthday_mail, birthday_name))
            continue
        if  birthday_boy[friend_id] == '' or  birthday_boy[friend_id + 1] == '':
            continue

        friends_names.append(birthday_boy[friend_id])
        friends_mails.append(birthday_boy[friend_id + 1])

    send_friends_names = friends_names[0]
    #if len(friends_names) > 1:
    #    send_friends_names = ", ".join(friends_names[0:-1]) + " и " + friends_names[-1]

    return friends_mails, friends_names


# Drafting and sending email notification to senders. You can add other senders' email in the list
def generate_email():

    birthdays_today = excel_birthdays("поздр от ДСР.xlsx")

    time_send = datetime.now().strftime("%m-%d %H:%M:%S")
    for birthday_boy in birthdays_today:
        birthday_name = birthday_boy[row_name]
        birthday_mail = birthday_boy[row_mail]

        header = birthday_name.split(' ')
        try:
            header = header[1] + " " + header[2] + ",\nС ДНЁМ РОЖДЕНИЯ!!"
        except:
            header = "С ДНЁМ РОЖДЕНИЯ!!"

        text_header = header if birthday_boy[row_header] == "" else (birthday_boy[row_header] + ",")
        text_body = messages[randint(0, 1)] if birthday_boy[row_body] == "" else birthday_boy[row_body]
        
        birthday_data = [birthday_mail, birthday_name, text_header, text_body]
        
        friends_mails, friends_names = get_friends_mails_and_names(birthday_boy, time_send, birthday_name, birthday_mail)
    
        friends = [friends_mails, friends_names]

        #print(birthday_data)
        mail_go_go(birthday_data, time_send, friends)


# Open Outlook.exe. Path may vary according to system config
# Please check the path to .exe file and update below

def open_outlook():
    os.chdir('C:/Program Files (x86)/Microsoft Office/Office16')
    print(os.getcwd())
    print(os.system('OUTLOOK.exe'))
    #subprocess.call(['OUTLOOK.exe'])


def check_opened():
    # Checking if outlook is already opened. If not, open Outlook.exe and send email
    for item in psutil.pids():
        p = psutil.Process(item)
        if p.name() == "OUTLOOK.EXE":
            return 1
    return 0

def start_process():
    if (check_opened()):
        generate_email()
    else:
        open_outlook()
        generate_email()
    # прикрутить проверку что сообщение отправлено
    # Как?


alarm_time = "11:00"

def reset_send_email_today():
    print("hoodie-ho")

schedule.every().day.at("19:09").do(start_process)
schedule.every().day.at("19:10").do(reset_send_email_today)

check_hour = alarm_time.split(':')[0]
check_min = alarm_time.split(':')[1]

if int(datetime.now().strftime("%H")) >= int(check_hour) and int(datetime.now().strftime("%M")) > int(check_min):
    try:
        start_process()
    except:
        sleep(10)
        start_process

while 1:
    try:
        schedule.run_pending()
    except:
        sleep(10)
    sleep(10)

