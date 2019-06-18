import win32com.client as win32
import psutil
import os
import subprocess
from time import sleep
from random import randint
from read_data_from_excel import excel_birthdays

# Drafting and sending email notification to senders. You can add other senders' email in the list
def send_notification():
    print(3)
    outlook = win32.Dispatch("Outlook.Application")
    temp = excel_birthdays("ДР_ДСР.xlsx")
    for elem in temp:
        mail = outlook.CreateItem(0)
        mail.To = elem[1]
        mail.CC = "kirik193@yandex.ru; kirik193@rambler.ru"
        mail.Subject = 'Sent through Python'
        messages = [f"first message with greetings", f'This email alert is auto generated. Please do not respond mr {elem[0]}']
        mail.body = messages[randint(0, 1)]
        mail.send
     
# Open Outlook.exe. Path may vary according to system config
# Please check the path to .exe file and update below
     
def open_outlook():
    print(1)
    subprocess.call([r'C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.exe'])
    print(1.5)

    

def check_opened():
    # Checking if outlook is already opened. If not, open Outlook.exe and send email
    for item in psutil.pids():
        p = psutil.Process(item)
        if p.name() == "OUTLOOK.EXE":
            return 1
    return 0

if (check_opened() == 1):
    send_notification()
else:
    open_outlook()
    print(2)
    #sleep(5)

    send_notification()

