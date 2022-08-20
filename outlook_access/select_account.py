import win32com.client

# initiate session with Outlook
outlook = win32com.client.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")

def select_account():
    print("List of active Outlook accounts: ")
    # list accounts in Outlook
    i = 0
    for account in mapi.Accounts:
        i += 1
        print(f"{ i }. { account.DeliveryStore.DisplayName }")

    query = input("Please select account: ")
    
    if query != i or query > i:


select_account()