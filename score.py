import smtplib
from email.mime.text import MIMEText
import gspread


#寄件人的信箱
gmail_user = 'Google帳號'
gmail_pwd = 'Google密碼'
address = '@ntu.edu.tw'
gc = gspread.login(gmail_user,gmail_pwd)

#開啟的文件
wks = gc.open("線性代數 4/23 課堂作業 及 批改結果").sheet1


#GMAIL的SMTP伺服器
smtpserver = smtplib.SMTP("smtp.gmail.com",587)
smtpserver.ehlo()
smtpserver.starttls()
smtpserver.ehlo()

#登入系統
smtpserver.login(gmail_user, gmail_pwd)

#寄件人資訊
fromaddr = "borching@gmail.com"

#內容
for i in range(1):
    str ="""<head><style>table, th, td{border-collapse:collapse;border:1px solid black;}th,td{padding:15px;}</style></head><table style="width:1200px">"""
    s = '線性代數 4/23 課堂作業批改結果 (組員:'+wks.cell(3+i,4).value+','+wks.cell(3+i,7).value+','+wks.cell(3+i,10).value+',本次得分'+wks.cell(3+i,2).value+')'
    for j in range(76):
        print(j);
        str=str+"<tr><td>"+(wks.cell(1,j+1).value or " ")+"</td><td>"+(wks.cell(3+i,j+1).value or " ")+"</td><td>"+(wks.cell(2,j+1).value or " ")+"</td></tr>"
    str=str+"</table>"
    msg=MIMEText(str,'html','utf-8') 
    msg['Subject'] = s
#組員信箱    
#    x=int(wks.cell(3+i,3).value)
#    if x==3:
#       toaddrs = [wks.cell(3+i,5).value.lower()+address,wks.cell(3+i,8).value.lower()+address,wks.cell(3+i,11).value.lower()+address]
#    if x!=3:
#       toaddrs = [wks.cell(3+i,5).value.lower()+address,wks.cell(3+i,8).value.lower()+address]

#測試信箱
    toaddrs = ['borching@gmail.com']
       
#設定寄件資訊
    smtpserver.sendmail(fromaddr, toaddrs, msg.as_string())
        
#登出
smtpserver.quit()


print('end')
