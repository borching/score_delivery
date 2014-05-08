import smtplib
from email.mime.text import MIMEText
import gspread

# In your local directory, create a key.py, containing the following things
# user = 'yourGmailAccount' 
# password = 'yourGmailPassword'

from key import *		
from param import *

#寄件人的信箱
gmail_user = user
gmail_pwd = password
address = '@ntu.edu.tw'
gc = gspread.login(gmail_user,gmail_pwd)

#開啟的文件
print("Connecting to the worksheet" + worksheet + "....")
wks = gc.open(worksheet).sheet1
print("Worksheet successfully connected")

#GMAIL的SMTP伺服器
print("Connecting to the SMTP server....")
smtpserver = smtplib.SMTP("smtp.gmail.com",587)
smtpserver.ehlo()
smtpserver.starttls()
smtpserver.ehlo()

#登入系統
smtpserver.login(gmail_user, gmail_pwd)
print("SMTP server successfully connected")

#寄件人資訊
fromaddr = "borching@gmail.com"

for i in userRange:
#組員信箱    
    x=int(wks.cell(3+i,3).value)
    if x==3:
       toaddrs = [wks.cell(3+i,5).value.lower()+address,wks.cell(3+i,8).value.lower()+address,wks.cell(3+i,11).value.lower()+address]
    if x!=3:
       toaddrs = [wks.cell(3+i,5).value.lower()+address,wks.cell(3+i,8).value.lower()+address]
    toaddrs.append("borching@gmail.com")
    
#內容
    str = "<html><head><style>table, th, td{border-collapse:collapse;border:1px solid black;}th,td{padding:15px;}</style></head><body>To:"
    str = str + ", ".join(toaddrs)
    str = str + mailFirstSentence
    str = str + """<table style="width:1200px">"""
    s = mailTitle
    s = s + '(組員:'+wks.cell(3+i,4).value+','+wks.cell(3+i,7).value+','+wks.cell(3+i,10).value+',本次得分'+wks.cell(3+i,2).value+')'
    for j in cellRange:
        print(j, end=" ");
        str=str+"<tr><td>"+(wks.cell(1,j+1).value or " ")+"</td><td>"+(wks.cell(3+i,j+1).value or " ")+"</td><td>"+(wks.cell(2,j+1).value or " ")+"</td></tr>"
    str = str + "</table></body></html>"
    msg=MIMEText(str,'html','utf-8') 
    msg['Subject'] = s

#測試信箱
    if test == 1:
        toaddrs = ['borching@gmail.com']
       
#設定寄件資訊
    smtpserver.sendmail(fromaddr, toaddrs, msg.as_string())
        
#登出
smtpserver.quit()


print('end')
