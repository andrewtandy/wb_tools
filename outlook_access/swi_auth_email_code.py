from datetime import datetime, timedelta

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