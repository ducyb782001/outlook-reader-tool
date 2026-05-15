import os
import win32com.client
import re

TEMP_FOLDER = "temp"


def connect_outlook():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    return outlook


def get_inbox():
    outlook = connect_outlook()
    inbox = outlook.GetDefaultFolder(6)
    return inbox


def find_latest_mail(subject_keyword=None, sender_keyword=None):
    inbox = get_inbox()

    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)

    for message in messages:

        try:
            subject = str(message.Subject)
            sender = str(message.SenderName)

            subject_match = (
                subject_keyword.lower() in subject.lower()
                if subject_keyword
                else True
            )

            sender_match = (
                sender_keyword.lower() in sender.lower()
                if sender_keyword
                else True
            )

            if subject_match and sender_match:
                return message

        except Exception as e:
            print(f"Error reading mail: {e}")

    return None


def download_attachments(message):

    if not os.path.exists(TEMP_FOLDER):
        os.makedirs(TEMP_FOLDER)

    downloaded_files = []

    attachments = message.Attachments

    for i in range(1, attachments.Count + 1):

        attachment = attachments.Item(i)

        filename = attachment.FileName
        save_path = os.path.join(TEMP_FOLDER, filename)

        attachment.SaveAsFile(os.path.abspath(save_path))

        downloaded_files.append(save_path)

        print(f"Downloaded: {filename}")

    return downloaded_files

def extract_password(message):
    try:
        body = str(message.Body)

        patterns = [
            r"Password:\s*(.+)",
            r"Pass:\s*(.+)",
            r"Mat khau:\s*(.+)",
            r"MK:\s*(.+)"
        ]

        for pattern in patterns:

            match = re.search(pattern, body, re.IGNORECASE)

            if match:
                password = match.group(1).strip()

                password = password.splitlines()[0]

                return password

        return None

    except Exception as e:
        print(f"Error extracting password: {e}")
        return None