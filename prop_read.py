import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = outlook.GetDefaultFolder(6).Parent.Folders("epdm") # "6" refers to the index of a folder - in this case,
                                    # the inbox. You can change that number to reference
                                    # any other folder
messages = inbox.Items
for i in range(3):
    m=messages[i]
    #print the message send date
    print(m.SentOn) 
    
# message = messages.GetLast()
# body_content = message.body
# body_content = body_content.split("\n")
# for i in body_content:
#     print(i)