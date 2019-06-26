#coding utf-8

import win32com.client as win32
import psutil
import os
import subprocess
import logging
import xlrd
import schedule

from time import sleep
from random import randint
from datetime import datetime

import logging

logging.basicConfig(filename="errors.log", level=logging.INFO)

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
                    birthday = datetime(*xlrd.xldate_as_tuple(excel_row[num + 2], wb.datemode)).strftime("%Y-%m-%d")
                except:
                    print("error date format in {}, {} row, {} cell. Must be dd.mm.yyyy (например 21.06.1990)".format(file_name, rownum, num + 3))
                    break
                today_date = datetime.today().strftime("%Y-%m-%d")
                if birthday == today_date:
                    list_of_adresses.append(excel_row[num:])
                break
    return (list_of_adresses)



def mail_go_go(email_address, reciever_name, copy_addresses, copy_names, time_send, *args):

    outlook = win32.Dispatch("Outlook.Application")

    mails = []
    names = []
    if len(args) > 1:
        mails = args[0]
        names = args[1]
    footer = ""
    for mail, name in zip(mails, names):
        footer += name + " (" + mail + ")\n"
    if footer != "":
        footer = "\n\n\nС уважением, \n" + footer
    try:
        mail = outlook.CreateItem(0)
        mail.To = email_address
        mail.CC = copy_addresses
        mail.Subject = 'С Днем Рождения!'
        messages = [
        "Прими самые добрые и искренние поздравления по случаю твоего дня рождения!\nС особой теплотой хотим сказать, что гордимся и дорожим сложившимися отношениями теплого сотрудничества и взаимопонимания.\nОт всей души желаем тебе крепкого здоровья, счастья, успехов в работе.\nПусть удача сопутствуют тебе и твои близким",
        'This email alert is auto generated. Please do not respond'
        ]
        header = "Мы, " + copy_names + " поздравляем тебя с Днем Рождения!\n" if len(copy_names) > 1 else ""
        mail.body = header + messages[randint(0, 1)] + footer
        mail.send
        logging.info(u"   {}: Send message to {}, copy for {}".format(time_send, email_address, copy_addresses))
    except Exception as e:
        logging.error(u"   {}: Error {}, while tried send message to {}".format(time_send, e, email_address))


def get_mails_and_names(birthday_boy, time_send, birthday_name, birthday_mail):

    friends_names = []
    friends_mails = []
    birthday_boy_size = len(birthday_boy)
  
    for friend_id in range(3, birthday_boy_size, 2):

        if friend_id + 1 > birthday_boy_size:
            logging.error("   {}: Error {}, именинник {} {}".format(time_send, "неверные данные у {}".format(birthday_boy[friend_id]), birthday_mail, birthday_name))
            continue
        if  birthday_boy[friend_id] == '' or  birthday_boy[friend_id + 1] == '':
            continue

        friends_names.append(birthday_boy[friend_id])
        friends_mails.append(birthday_boy[friend_id + 1])

    send_friends_names = friends_names[0]
    if len(friends_names) > 1:
        send_friends_names = ", ".join(friends_names[0:-1]) + " и " + friends_names[-1]

    return ';'.join(friends_mails), send_friends_names, friends_mails, friends_names


# Drafting and sending email notification to senders. You can add other senders' email in the list
def send_notification():

    birthdays_today = excel_birthdays("ДР_ДСР.xlsx")

    time_send = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for birthday_boy in birthdays_today:
        birthday_name = birthday_boy[0]
        birthday_mail = birthday_boy[1]

        send_friends_mails, send_friends_names, friends_mails, friends_names = get_mails_and_names(birthday_boy, time_send, birthday_name, birthday_mail)

        mail_go_go(birthday_mail, birthday_name, send_friends_mails, send_friends_names, time_send, friends_mails, friends_names)


# Open Outlook.exe. Path may vary according to system config
# Please check the path to .exe file and update below

def open_outlook():
    subprocess.call([r'C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.exe'])


def check_opened():
    # Checking if outlook is already opened. If not, open Outlook.exe and send email
    for item in psutil.pids():
        p = psutil.Process(item)
        if p.name() == "OUTLOOK.EXE":
            return 1
    return 0

def start_process():
    if (check_opened() == 1):
        send_notification()
    else:
        open_outlook()
        #sleep(5)
        send_notification()



alarm_time = "11:00"

schedule.every().day.at(alarm_time).do(start_process)

check_hour = alarm_time.split(':')[0]
check_min = alarm_time.split(':')[1]

if int(datetime.now().strftime("%H")) > int(check_hour) and int(datetime.now().strftime("%M")) > int(check_min):
    start_process()
while 1:
    try:
        schedule.run_pending()
    except:
        sleep(10)
    sleep(10)

