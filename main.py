from outlook_reader import find_latest_mail
from outlook_reader import download_attachments
from outlook_reader import extract_password

def main():

    print("Searching mail...")

    mail = find_latest_mail(
        subject_keyword="BACKUP",
        sender_keyword=None
    )

    if not mail:
        print("No matching mail found")
        return

    print("Mail found")
    print("Subject:", mail.Subject)
    password = extract_password(mail)
    if password:
        print("Password found:", password)
    else:
        print("Password not found")

    files = download_attachments(mail)

    print("\nDownloaded files:")

    for file in files:
        print(file)


if __name__ == "__main__":
    main()