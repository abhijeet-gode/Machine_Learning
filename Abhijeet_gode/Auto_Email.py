import win32
import win32com.client as win32
import os
import pandas as pd
import getpass
import datetime
from datetime import datetime
import re
import glob
month = datetime.now().strftime("%b")
year = datetime.now().strftime("%Y")
user = getpass.getuser()
path = "C:\\Users\\"+ str(getpass.getuser()) +"\\Box\\EMEA Mumbai Data\\Operations\\Sourcing_Donottouch\\2022\\Jul-2022\\Chasing Sheet_{}{}".format(month, year+".xlsx")
Chasing_sheet = pd.read_excel(path, sheet_name = "RMBS,ABS")#, skiprows=1
Chasing_sheet.reset_index(drop=True, inplace=True)
Network_days = int(input('Enter the Network days: '))
Chasing_sheet['Received Date'] = (Chasing_sheet['Received Date'].astype(str).replace({'NaT': ''}))
emailer = Chasing_sheet[(Chasing_sheet['Network days'] == Network_days) & (Chasing_sheet['Received Date'] == '')]
ab = emailer[['Receipient List', 'DEAL_NAME']].dropna()#.str.split(';')
ind = ab.reset_index(drop=True)
deal = ind['DEAL_NAME']
outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
folder = outlook.Folders("EMEAServicer")
sentItems = folder.Folders("Inbox")
messages_sent = sentItems.Items
found = False
for i in range(len(deal)):
    message_subject_to_find = deal[i]
    subject_found = ''
    for message in messages_sent:
        if message.Class == 43:
            if message_subject_to_find in message.Subject:
                subject_found = message.Subject
                found = True
                reply = message.ReplyAll()
                reply.HTMLBody = f"""
<b> Hi Team,</b><br><br> 

This is a friendly reminder that the following report is due.<br><br>
{deal[i]} period.<br><br>

Please email the report to <u>rmbseuropeansurveillance@spglobal.com</u> or publish them on the agreed website as soon as possible. If the report is not yet available please let us know when you expect it to be available.<br><br>

If you are not the contact for this report, or if you believe we have the wrong due date for this report, kindly let us know and we will work to correct this.<br><br>
Please ignore this email if the report has already been made available.<br><br>

"""
                reply.Display()
                break
if found:
  print('Done for this item !! -> {}'.format(subject_found))
else:
  print('Subject not found')