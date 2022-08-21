import win32com.client

# initiate session with Outlook
outlook = win32com.client.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")

def select_account():
    print("List of active Outlook accounts: ")
    
    # place all accounts into dictionary
    account_dict = {}
    i = 0
    for account in mapi.Accounts:
        i += 1
        account_dict[i] = account.DeliveryStore.DisplayName

    # list all accounts for selection by user
    for index, account_name in account_dict.items():
        print(f"{ index }. { account_name }")

    # input choice, check valid
    # TODO: provide escape or previous menu action
    choice = int(input(f"Select account: "))
    if choice not in account_dict:
        print(f"{ choice } not valid.")
        select_account()
    
    # return str of account choice
    return account_dict[choice]