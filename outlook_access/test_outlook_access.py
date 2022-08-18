import win32com.client
#other libraries to be used in this script
import os
from datetime import datetime, timedelta

# initiate session with Outlook
outlook = win32com.client.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")

# if multiple accounts in Outlook, need to pass account name when accessing folders, the following 
# assumes only 1 account
for account in mapi.Accounts:
	print(account.DeliveryStore.DisplayName)

# TODO : dynamic input for specific account/folder access
account_name = 'andrew.tandy@waterbabies.co.uk'

# to access inbox folder, pass folder type (6 in this instance) - full list of folder types here:
# https://docs.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders
# 
# subfolder access can be achieved with .Folders att. e.g. 'your_sub_folder':
# inbox = mapi.GetDefaultFolder(6).Folders["your_sub_folder"]
inbox = mapi.Folders(account_name).Folders('Inbox').Items

# for m in inbox:
#     print(m)

# collect all messages from folder
messages = inbox

# message = messages.GetFirst()
# body_content = message.body

# print(body_content)

# filter messages using Restrict function
received_dt = datetime.now() - timedelta(days=1)
received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
messages = messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")
messages = messages.Restrict("[SenderEmailAddress] = 'no-reply@waterbabies.co.uk'")
messages = messages.Restrict("[Subject] = 'WaterBabies Authentication Code'")

# print all messages subject to filter criteria
for message in list(messages):
    print(message.body)