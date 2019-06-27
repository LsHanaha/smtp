@ECHO OFF
cd C:/Program Files (x86)/Microsoft Office/Office16
START OUTLOOK.exe
timeout /t 20
cd C:/Users/Kirill/Desktop/smtp
python send_mail.py
PAUSE