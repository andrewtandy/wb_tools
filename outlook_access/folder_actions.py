import win32com.client

# initiate session with Outlook
outlook = win32com.client.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")

# select account and folder
def select_email_folder(account_name, folder):
    folder_emails = mapi.Folders(account_name).Folders(folder).Items
    return folder_emails

# set filters on folder using kwargs filters and Restrict()
def filter_emails(emails, filters): 
    # TODO: Make this prettier!!
    i = 0
    while i < len(filters):
        print(filters[i])
        emails = emails.Restrict(filters[i])
        i += 1

        if i == len(filters):
            message = emails.GetFirst()
            return message.body