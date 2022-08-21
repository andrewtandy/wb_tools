import win32com.client

# following function taken from S/O, tried to implement own version but couldn't get to work
# https://stackoverflow.com/questions/68775139/using-python-to-extract-the-names-of-every-folder-in-outlook-mailbox
# TODO: implement below in own 

def list_all_folders(folders):
    my_list = []
    for folder in folders:
        print(folder.name)
        my_list.append(folder.name)
        # my_list += list_all_folders(folder.Folders)
    return my_list

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
# user = outlook.Application.Session.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress
user = 'andrew.tandy@waterbabies.co.uk'
z = list_all_folders(outlook.Folders[user].Folders)