import win32com.client
from datetime import datetime, timedelta

# initiate session with Outlook
outlook = win32com.client.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")

# select account and folder
def select_email_folder(account_name, folder):
    folder_emails = mapi.Folders(account_name).Folders(folder).Items
    return folder_emails

# set filters on folder using kwargs filters and Restrict()
def filter_emails(emails, **filters):
    for value in filters.items():
        emails = emails.Restrict(value)
    return emails


