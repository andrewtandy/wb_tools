# import win32com.client

#other libraries to be used in this script
# import os
from datetime import datetime, timedelta

# initiate session with Outlook
# outlook = win32com.client.Dispatch('outlook.application')
# mapi = outlook.GetNamespace("MAPI")


'''
 if multiple accounts in Outlook, need to pass account name when accessing folders, the following 
 assumes only 1 account 
'''
# for account in mapi.Accounts:
	# print(account.DeliveryStore.DisplayName)

# TODO : dynamic input for specific account/folder access
# account_name = 'andrew.tandy@waterbabies.co.uk'

'''
 to access inbox folder, pass folder type (6 in this instance) - full list of folder types here:
 https://docs.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders

 subfolder access can be achieved with .Folders att. e.g. 'your_sub_folder':
 inbox = mapi.GetDefaultFolder(6).Folders["your_sub_folder"]
'''
# inbox = mapi.Folders(account_name).Folders('Inbox').Items

# collect all messages from folder
# messages = inbox

# set message filter types
# message_sender = 'no-reply@waterbabies.co.uk'
# message_subject = 'WaterBabies Authentication Code'

# received_dt = datetime.now() - timedelta(days=1)
# received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')

# filter messages using Restrict function
# messages = messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")
# messages = messages.Restrict(f"[SenderEmailAddress] = '{ message_sender }'")
# messages = messages.Restrict(f"[Subject] = '{ message_subject }'")

# use GetFirst to locate latest mail received. Consider loop if need to access more than
# 1 message
# message = messages.GetFirst()
# body = message.body

# use the known word 'Account' to find index in body and pull auth code
# keyword = body.index("Account.")
# code = body[keyword+12:keyword+16]
# print(code)

import folder_actions

account_name = 'andrew.tandy@waterbabies.co.uk'
folder = 'Inbox'

emails = folder_actions.select_email_folder(account_name, folder)

message_sender = 'no-reply@waterbabies.co.uk'
message_subject = 'WaterBabies Authentication Code'

days = 1
received_dt = datetime.now() - timedelta(days=days)
received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')

filters = [
	f"[ReceivedTime] >= '{ received_dt }'",
	f"[SenderEmailAddress] = '{ message_sender }'",
	f"[Subject] = '{ message_subject }'"
]


filtered = folder_actions.filter_emails(emails, filters)

