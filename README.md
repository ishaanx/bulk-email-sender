# 📧 Bulk Email Sender (PySide6)

A simple desktop GUI app for sending bulk emails using an Excel file of recipients. Supports both **HTML** and **plain text** emails, automatic delays, and progress tracking.

---

## 🚀 Features

- ✅ Send HTML or Plain Text Emails
- 📁 Import Recipients from Excel (.xlsx) — with columns for Name and Email
- ⏱️ 30-Second Delay Between Sends
- 🔐 SMTP Authentication (Gmail, Outlook, etc.)
- 📊 Progress Bar to show email sending progress
- 🌓 Dark Mode Toggle
- 📄 Logs saved to timestamped `.log` file
- 📥 Export Success/Failure Counts after sending

---

## 📂 Excel Format

Make sure your Excel file has headers in the first row:

| Email               |
|---------------------|
| john@example.com    |
| jane@example.com    |

---

## 🧪 Email Body Format

### Plain Text
Hello,

This is a plain text email.

Regards,
Team


### HTML (with optional variables like `{{name}}`)

```html
<html>
  <body>
    <h2>Hello!</h2>
    <p>This is a personalized email sent using the Bulk Email Sender.</p>
    <p>Regards,<br>Your Team</p>
  </body>
</html>
```
---
### 🛠 How to Run
1. Install Dependencies
Make sure Python 3.7+ is installed. Then install required packages:

```
pip install PySide6 openpyxl pyinstaller
```

2. Build  the App

```
pyinstaller --onefile scr.py
```

- --onefile creates a single executable.

The output EXE will be in the dist/ folder.

---
### 📤 SMTP Setup
Example for Gmail:

```
SMTP Server: smtp.gmail.com
Port: 587
```

Make sure to enable App Passwords in your Google Account settings and use the app password instead of your regular password.

---
### 📁 Logs & Export
Logs are saved automatically to a file named like: log_2025-04-14_15-32-10.log

After all emails are sent, you can export a summary of success/failure counts.

---
### 🌓 Dark Mode
Use the toggle in the top navigation bar to switch between light and dark mode.

---
### 🙌 Contributing
Pull requests and feature suggestions welcome!
