import win32com.client
from win32com.client import Dispatch, constants
 
const=win32com.client.constants
olMailItem = 0x0
obj = win32com.client.Dispatch("Outlook.Application")
newMail = obj.CreateItem(olMailItem)
newMail.Subject = "I AM test!!"
newMail.Body = "I AM IN THE BODY\nSO AM I!!!"
newMail.To = "kirik193@yandex.ru"

newMail.display()
temp = newMail
print(type(newMail))
temp.send