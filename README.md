# 📧 Email Attachment Forwarder (Python Automation)

This Python script automates the process of locating a specific email in an Outlook mailbox, verifying and downloading a time-sensitive attachment, and preparing a new email with that file attached. It reduces manual effort and ensures accuracy in daily email workflows.

---

## 🚀 What It Does

✅ Searches specific Outlook inbox folders for an email matching:
- A known **sender name**
- A known **subject line**

✅ Within that email:
- Identifies an attachment with a filename like `123_ABC_YYYYMMDD`, where the date matches the **previous business day**

✅ Then:
- Downloads the attachment to a temporary folder
- Prepares a new Outlook email with preset recipients, subject, body, and the attachment
- Deletes the temporary file after composing the email

---

## 🛠 Technologies Used

- Python 3
- `pywin32` (for Outlook automation)
- `datetime`, `os` (for file and date handling)

---

## 🧩 How to Use

1. Install required libraries:
   ```bash
   pip install -r requirements.txt
