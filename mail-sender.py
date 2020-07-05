import smtplib, os, shutil, openpyxl

from openpyxl import Workbook
from datetime import datetime
from os import path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication


#ИЗПРАЩАНЕ НА МЕЙЛ
def send_test_mail(body):
    sender_email = '-----------------'
    receiver_email = '---------------'
    msg = MIMEMultipart()
    msg['Subject'] = '--------------------------------'
    msg['From'] = sender_email
    msg['To'] = receiver_email

    msgText = MIMEText('<b>%s</b>' % (body), 'html')
    msg.attach(msgText)

    pdf = MIMEApplication(open('-----------------.xlsx', 'rb').read()) #Добавяне на файл към мейла
    pdf.add_header('Content-Disposition','attachment',filename = 'tablica.xls')
    msg.attach(pdf)

    try:
      with smtplib.SMTP('------------------', --) as smtpObj: #server + port (from web provider)
        smtpObj.ehlo()
        smtpObj.starttls()
        smtpObj.login("-------------", "-----------") #acc + pass
        smtpObj.sendmail(sender_email, receiver_email, msg.as_string())
    except Exception as e:
      print(e)


send_test_mail("Поздрави и успешен ден, юрк. Георгиев!") #calling the function(sending the email) + the text message!

#КОПИРАНЕ НА ФАЙЛА В ПАПКА АРХИВ
src = '-----------------.xlsx'
dst = '----------------.xlsx'
shutil.copy(src, dst)


#ПРЕИМЕНУВАНЕ НА АРХИВ ФАЙЛА С ДАТА И ЧАС НА ИЗПРАЩАНЕТО
now = datetime.now()
today = now.strftime("--------------------/Архив/%d-%b-%Y %H-%M-%S.xlsx")
def main():
    os.rename("---------------------------/Архив/tablica.xlsx",str(today))

main()


#ИЗЧИСТВАНЕ НА ОРИГИНАЛНИЯ ФАЙЛ, ЗА ДА Е ГОТОВ ЗА РАБОТА
wbkName = '-----------------------------------.xlsx'
wbk = openpyxl.load_workbook(wbkName, keep_vba=True)

for wks in wbk.worksheets:
    for row in wks["B2:E300"]:
        for cell in row:
            cell.value = None

wbk.save(wbkName)
wbk.close()





print("Успешно завършен процес!")
