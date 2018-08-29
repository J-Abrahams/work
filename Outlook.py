import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)

messages = inbox.Items
message = messages.GetFirst()
rec_time = message.CreationTime
body_content = message.body
subj_line = message.subject

while message:
    print(message.subject, message.CreationTime)
    message = messages.GetNext()