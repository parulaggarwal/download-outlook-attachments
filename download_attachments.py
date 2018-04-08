import win32com.client
import os

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

'''
This method validates if the message satisfies some conditions, say the sender name and subject.
Refer to more properties of email message at https://msdn.microsoft.com/en-us/vba/outlook-vba/articles/mailitem-object-outlook 
'''
def validate_message(message):
    if 'Hogwarts' in message.subject and 'Harry Potter' in message.SenderName:
        return True
    return False


def download_attachment_from_message(message):
    attachments = message.attachments
    for attachment in attachments:
        download_attachment(attachment)


def download_attachment(attachment):
    storage_location = os.getcwd() + "\\" + attachment.FileName
    attachment.saveAsFile(storage_location)
    return storage_location


'''
If the attachment is an email file itself, and you want to download content inside that attachment
'''
def download_attachment_inside_email_attachment(message):
    attachments = message.attachments
    for attachment in attachments:
        if attachment.FileName.encode("utf-8").endswith("msg"):
            storage_location = download_attachment(attachment)
            attached_email = outlook.OpenSharedItem(storage_location)
            download_attachment_from_message(attached_email)


def main():
    inbox = outlook.Folders.Item("<your email address").Folders.Item("inbox")

    '''
    The following can be used if we want to access any subfolder inside Inbox. 
    subfolder_in_inbox = inbox.Folders.Item("subfolder name")
    '''

    messages = inbox.Items
    for message in messages:
        if validate_message(message):
            download_attachment_from_message(message)


if __name__ == '__main__':
    main()
