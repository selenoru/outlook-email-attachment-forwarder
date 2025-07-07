import win32com.client
import os
from datetime import datetime, timedelta

# Change this to a valid folder on your system
save_path = r"C:\Users\YourName\Documents\temp_folder"

def get_previous_business_day():
    today = datetime.today()
    one_day = timedelta(days=1)
    previous = today - one_day
    while previous.weekday() >= 5:  # Saturday = 5, Sunday = 6
        previous -= one_day
    return previous.strftime("%Y%m%d")

# Access Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Access mailbox and folders (replace with your structure)
mailbox = outlook.Folders.Item("youremail@example.com")
main_folder = mailbox.Folders.Item("Main Folder")

# First try "Team A" folder
print("Searching Team A folder...")
sender1_folder = main_folder.Folders.Item("Team A")
messages = sender1_folder.Items
messages.Sort("[ReceivedTime]", True)

found_mail = None

# Try to find email in first folder
for msg in messages:
    try:
        if msg.SenderName.strip() == "Sender One" and msg.Subject.strip() == "ABC Daily Report":
            found_mail = msg
            print("Found email in Team A folder.")
            break
    except AttributeError:
        continue

# If not found, try "Team B" folder
if not found_mail:
    print("Email not found in Team A. Trying Team B...")
    sender2_folder = main_folder.Folders.Item("Team B")
    messages = sender2_folder.Items
    messages.Sort("[ReceivedTime]", True)
    for msg in messages:
        try:
            if msg.SenderName.strip() == "Sender Two" and msg.Subject.strip() == "ABC Daily Report":
                found_mail = msg
                print("Found email in Team B folder.")
                break
        except AttributeError:
            continue

# If still not found
if not found_mail:
    print("No matching email found in either folder.")
    exit()

# Build attachment filename prefix
attachment_prefix = f"123_ABC_{get_previous_business_day()}"

# Look for the correct attachment
if found_mail:
    for attachment in found_mail.Attachments:
        if attachment.FileName.startswith(attachment_prefix):
            os.makedirs(save_path, exist_ok=True)
            full_path = os.path.join(save_path, attachment.FileName)
            attachment.SaveAsFile(full_path)
            print(f"Attachment saved to: {full_path}")

            # Create new email
            outlook_app = win32com.client.Dispatch("Outlook.Application")
            new_mail = outlook_app.CreateItem(0)
            new_mail.SentOnBehalfOfName = "youremail@example.com"
            new_mail.To = "recipient@example.com"
            new_mail.CC = "ccperson@example.com"
            new_mail.Subject = "ABC Report // Daily Update"
            new_mail.Body = (
                "Hello,\n\n"
                "Attached is the ABC file.\n\n"
                "Best regards,"
            )
            new_mail.Attachments.Add(full_path)
            new_mail.Display()
            print("Email ready to send!")

            # Optionally delete the file after displaying
            os.remove(full_path)
            print("Temporary file deleted.")
            break
    else:
        print(f"No attachment starting with '{attachment_prefix}' was found.")
else:
    print("No matching email found.")
