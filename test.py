import win32com.client as win32
import psutil
import os
import subprocess
from time import sleep
import base_smpt as bs
 
# Drafting and sending email notification to senders. You can add other senders' email in the list
def send_notification():
    outlook = win32.Dispatch("Outlook.Application")
    temp = bs.select_all_persons()
    for elem in temp:
        mail = outlook.CreateItem(0)
        mail.To = elem[1]
        mail.Subject = 'Sent through Python'
        mail.body = f'This email alert is auto generated. Please do not respond mr {elem[0]}'
        mail.send
     
# Open Outlook.exe. Path may vary according to system config
# Please check the path to .exe file and update below
     
def open_outlook():
    try:
        print(3)
        subprocess.call([r'C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.exe'])
        print(4)
        os.system(r"C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.exe")
        print(5)
    except:
        print("no-no")

# Checking if outlook is already opened. If not, open Outlook.exe and send email
for item in psutil.pids():
    p = psutil.Process(item)
    if p.name() == "OUTLOOK.EXE":
        flag = 1
        break
    else:
        flag = 0
 
if (flag == 1):
    send_notification()
else:
    open_outlook()
    print(1)
    #sleep(5)
    print(2)
    send_notification()

