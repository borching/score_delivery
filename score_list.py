import smtplib
from email.mime.text import MIMEText
import gspread

# In your local directory, create a key.py, containing the following things
# user = 'yourGmailAccount' 
# password = 'yourGmailPassword'

from key import *		

#寄件人的信箱
gmail_user = user
gmail_pwd = password
address = '@ntu.edu.tw'
gc = gspread.login(gmail_user,gmail_pwd)

#開啟的文件
wks = gc.open("102-2線代成績").sheet1


#GMAIL的SMTP伺服器
smtpserver = smtplib.SMTP("smtp.gmail.com",587)
smtpserver.ehlo()
smtpserver.starttls()
smtpserver.ehlo()

#登入系統
smtpserver.login(gmail_user, gmail_pwd)

#寄件人資訊
fromaddr = "borching@gmail.com"

for i in range(71):
#學生信箱    
    sid = wks.cell(2 + i, 4).value;
    name = wks.cell(2 + i, 5).value;
    toaddrs = [sid.lower() + address]
    toaddrs.append("borching@gmail.com")
    print(sid + name, end = ":")
#內容
    str = "<html><head><style>table, th, td{border-collapse:collapse;border:1px solid black;}th,td{padding:15px;}</style></head><body>To:"
    str = str + ", ".join(toaddrs) + "<BR>"
    str = str + "以下為系統記錄之歷次課堂練習分數統計，若有問題請洽老師。小考與期中考之成績請見助教之公告（不在此列中）。<BR>"
    str = str + "目前統計Week3 ~ Week10 (扣掉Week5(春假), Week6(期中考前複習)) 共６週之成績。<BR>"
    str = str + "Week 8以前每週成績=(2/3)*Wed+(1/3)*Thr; Week9以後每週成績 = 1 * Wed.<BR>"
    str = str + "預計至期末將有１２週之成績。平時成績計算以１２次中取８次最高的成績之平均作計算。<BR>"
    str = str + """<table style="width:300px">"""
    s = '102-2線性代數課堂作業歷次成績核對(至4/23止) (學號:'+ sid +', 姓名:'+ name +')'
    for j in range(22):
        print(j, end=" ");
        str=str+"<tr><td>"+(wks.cell(1, j + 1).value or " ")+"</td><td>"+(wks.cell(2 + i, j + 1).value or " ")+"</td></tr>"
    str = str + "</table></body></html>"
    msg=MIMEText(str,'html','utf-8') 
    msg['Subject'] = s
    print(" ")

#測試信箱
#    toaddrs = ['borching@gmail.com']
       
#設定寄件資訊
    smtpserver.sendmail(fromaddr, toaddrs, msg.as_string())
        
#登出
smtpserver.quit()


print('end')
