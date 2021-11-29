import requests
import win32com.client
import os
from datetime import datetime,timedelta
import re

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case,
                                    # the inbox. You can change that number to reference
                                    # any other folder
messages = inbox.Items
received_dt = datetime.today() 
received_dt = received_dt.strftime('%m/%d/%Y')+ " 00:00 AM"
print(received_dt)
messages = messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")
messages = messages.Restrict("[SenderEmailAddress] = 'do-not-reply@uwaterloo.ca'")
message = messages.Restrict("[Subject] = 'Welcome to Campus - Thank you for completing your COVID-19 screening'")
if message.GetFirst():
    print("already done for today")
else:
    messages = messages.Restrict("[Subject] = 'Welcome to Campus - Please complete COVID-19 screening'")

    message = messages.GetFirst()
    body_content = message.body
    url = re.search("(?P<url>https?://[^\s]+)", body_content).group("url")
    print(url)

    #url = "https://checkin.uwaterloo.ca/campuscheckin/screen.php?key=thqRYCyfdJWR4WYuhYPqM70zXITENdD1"
    ans = {"q8":"No","q1":"No","q2":"No","q3":"No","q4":"No","q5":"No","q6":"No","q7":"No","q8":"No","ukey":"g1jDL7vFufPszHbvH9q","utime":1638081584590,"what":"Submit"}
    x = requests.post(url, data = ans)
    length = len(x.text)
    print(length)
    utime = x.text[-185:-175]
    ukey = x.text[-246:-230]
    print(ukey,utime)
    # ukey = str[5324:5340]
    # utime = str[5385:5395]

    #print(ukey,utime)


    ans = {"q8":"No","q1":"No","q2":"No","q3":"No","q4":"No","q5":"No","q6":"No","q7":"No","q8":"No","ukey":str(ukey),"utime":int(utime),"what":"Submit"}
    x= requests.post(url,data = ans)
    print(x.text)
