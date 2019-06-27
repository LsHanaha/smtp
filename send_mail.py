# -*- coding: utf-8 -*-

import win32com.client as win32
import psutil
import os, sys
import logging
import xlrd
import schedule

from time import sleep
from random import randint
from datetime import datetime
from logging.handlers import RotatingFileHandler

send_email_today = 0

row_name = 0
row_mail = 1
row_date = 2
row_header = 3
row_body = 4

messages = [
"Прими самые добрые и искренние поздравления по случаю твоего дня рождения!\nС особой теплотой хотим сказать, что гордимся и дорожим сложившимися отношениями теплого сотрудничества и взаимопонимания.\nОт всей души желаем тебе крепкого здоровья, счастья, успехов в работе.\nПусть удача сопутствуют тебе и твои близким",

'Пусть и на работе, и в семье тебе сопутствует успех и благополучие. Желаем успешного осуществления всех твоих благих начинаний, а еще — оставаться таким же профессионалом своего дела и просто замечательным человеком!\nПусть удача сопутствуют тебе и твои близким'
]


if not os.path.exists('logs'):
    os.mkdir('logs')
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
file_handler = RotatingFileHandler('logs/mailer.log', maxBytes=1000000,
                                       backupCount=10)

formatter = logging.Formatter(
    '%(asctime)s %(levelname)s: %(message)s [in %(pathname)s:%(lineno)d] %(module)s.%(funcName)s')
file_handler.setFormatter(formatter)
logger.addHandler(file_handler)

def start_logging():
    logger.info("START LOGGING")
# Open Outlook.exe. Path may vary according to system config
# Please check the path to .exe file and update below

def open_outlook():
    temp = os.getcwd()
    os.chdir('C:/Program Files (x86)/Microsoft Office/Office16')
    os.system('START OUTLOOK.exe')
    os.chdir(temp)

def check_opened():
    for item in psutil.pids():
        p = psutil.Process(item)
        if p.name() == "OUTLOOK.EXE":
            return 1
    return 0

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
                    logger.error("wrong date format in {}, {} row, {} cell. Must be dd.mm.yyyy (например 21.06.1990)".format(file_name, rownum, num + 3))
                    break
                today_date = datetime.today().strftime("%m-%d")
                if birthday == today_date:
                    list_of_adresses.append(excel_row[num:])
                break
    return (list_of_adresses)


def mail_go_go(birthday_data, time_send, friends):

    try:
        outlook = win32.Dispatch("Outlook.Application")
    except Exception as e:
        logger.error("need to open outlook in mail_go_go")
        open_outlook()
        sleep(20)
        outlook = win32.Dispatch("Outlook.Application")
    logger.info("Outlook worked")

    mails = friends[0]
    names = friends[1]
    logger.info(str(mails) + str(names))
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
        logger.info(u"   {}: Send message to {}, copy for {}".format(time_send, birthday_data[0], ",".join(mails)))
    except Exception as e:
        logger.error(u"   {}: Error {}, while tried send message to {}".format(time_send, e, birthday_data[0]))


def get_friends_mails_and_names(birthday_boy, time_send, birthday_name, birthday_mail):

    friends_names = []
    friends_mails = []
    birthday_boy_size = len(birthday_boy)

    for friend_id in range(row_body + 1, birthday_boy_size, 2):

        if friend_id + 1 > birthday_boy_size:
            logger.error("   {}: Error {}, именинник {} {}".format(time_send, "неверные данные у {}".format(birthday_boy[friend_id]), birthday_mail, birthday_name))
            continue
        if  birthday_boy[friend_id] == '' or  birthday_boy[friend_id + 1] == '':
            continue

        friends_names.append(birthday_boy[friend_id])
        friends_mails.append(birthday_boy[friend_id + 1])

    send_friends_names = friends_names[0]
    #if len(friends_names) > 1:
    #    send_friends_names = ", ".join(friends_names[0:-1]) + " и " + friends_names[-1]
    return friends_mails, friends_names

def reset_send_email_today():
    logger.info("Reset email send")
    with open("logs/qtc_pp.bat", 'w') as f:
        f.write("0")

def set_send_email_today():
    logger.info("Set email send")
    with open("logs/qtc_pp.bat", 'w') as f:
        f.write("1")

def check_sended():
    try:
        with open("logs/qtc_pp.bat", 'r') as f:
            text = f.read()
            if text == "1":
                res = 0
            else:
                res = 1
    except:
        res = 0
    return res

# Drafting and sending email notification to senders. You can add other senders' email in the list
def generate_email():

    birthdays_today = excel_birthdays("поздр от ДСР.xlsx")

    time_send = datetime.now().strftime("%m-%d %H:%M:%S")

    set_send_email_today()

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

        mail_go_go(birthday_data, time_send, friends)


def start_process():
    logger.info("start for today")
    while not check_opened():
        logger.warning("Need to open outlook")
        open_outlook()
        sleep(20)
    generate_email()

start_logging()

alarm_time = "11:00"

schedule.every().day.at(alarm_time).do(start_process)
schedule.every().day.at("00:00").do(reset_send_email_today)

check_hour = alarm_time.split(':')[0]
check_min = alarm_time.split(':')[1]

if int(datetime.now().strftime("%H")) >= int(check_hour) and int(datetime.now().strftime("%M")) > int(check_min) and not check_sended():
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


