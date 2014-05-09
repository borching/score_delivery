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

ll = wks.get_all_values()

for i in userRange:
    if sendMode == 1: # 當週分組成績
#組員信箱    
        x = int(ll[2+i][2])
        if x==3:
            toaddrs = [ll[2+i][4].lower()+address, ll[2+i][7].lower()+address, ll[2+i][10].lower()+address]
        if x!=3:
            toaddrs = [ll[2+i][4].lower()+address, ll[2+i][7].lower()+address]
        toaddrs.append(fromaddr)
        print("Processing the " + str(i) + "th group...")
    elif sendMode == 2: # 全學期個人成績
#學生信箱    
        sid = wks.cell(2 + i, 4).value;
        name = wks.cell(2 + i, 5).value;
        toaddrs = [sid.lower() + address]
        toaddrs.append("borching@gmail.com")
        print("Processing the " + str(i) + "th person: " + sid + name, end = ":")
    
#內容
    st = "<html><head><style>table, th, td{border-collapse:collapse;border:1px solid black;}th,td{padding:15px;}</style></head><body>To:"
    st = st + ", ".join(toaddrs)
    st = st + mailFirstSentence
    st = st + """<table style="width:1200px">"""
    s = mailTitle
    if sendMode == 1: # 當週分組成績
        s = s + '(組員:'+ ll[2+i][3] + ',' + ll[2+i][6] + ',' + ll[2+i][9] + ',本次得分' + ll[2+i][1] +')'
    elif sendMode == 2: # 全學期個人成績
        s = s + '(學號:'+ sid +', 姓名:'+ name +')'
    for j in cellRange:
        print(j, end=" ");
        #st = st +"<tr><td>"+(wks.cell(1,j+1).value or " ")+"</td><td>"+(wks.cell(3+i,j+1).value or " ")+"</td><td>"+(wks.cell(2,j+1).value or " ")+"</td></tr>"
        if sendMode == 1: # 當週分組成績
            st = st +"<tr><td>"+(ll[0][j] or " ")+"</td><td>"+(ll[2+i][j] or " ")+"</td><td>"+(ll[1][j] or " ")+"</td></tr>"
        elif sendMode == 2: # 全學期個人成績
            st = st +"<tr><td>"+(ll[0][j] or " ")+"</td><td>"+(ll[1+i][j] or " ")+"</td></tr>"

    st = st + "</table></body></html>"
    msg = MIMEText(st,'html','utf-8') 
    msg['Subject'] = s

#測試信箱
    if test == 1:
        toaddrs = ['borching@gmail.com']
       
#設定寄件資訊
    print("Sending the mail to the " + str(i) + "th group...", end=" ")
    smtpserver.sendmail(fromaddr, toaddrs, msg.as_string())
    print("OK")
        
#登出
smtpserver.quit()


print('end')
