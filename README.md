# 📧 Automated Bulk Email Sender (Tkinter)

## 📌 Overview
This is a Python-based GUI application that allows users to send bulk emails using predefined templates. It supports personalization, CSV contact import, and email history tracking — simulating a real-world email marketing system.

## 🚀 Features
- Send emails to multiple recipients
- Load contacts from CSV file
- Predefined email templates with placeholders ({name}, {company})
- Personalized email content for each recipient
- Gmail SMTP integration (secure via App Password)
- Email sending history with status tracking (Sent/Failed)
- User-friendly Tkinter GUI with multiple tabs

## 🛠 Technologies Used
- Python
- Tkinter (GUI)
- smtplib (Email sending)
- JSON & CSV (Data storage)

## 📂 Project Structure
- `main.py` → Main application file  
- `history.json` → Stores sent email logs  
- `contacts.csv` → Input file for bulk contacts  

## 📋 CSV Format
Your CSV file must include:
- Name
- Email
- Company

## 🔐 Gmail Setup (Important)
- Enable **2-Step Verification**
- Generate an **App Password**
- Use the App Password in the application (NOT your real password)

## ▶️ How to Use
1. Run the program  
2. Enter your Gmail and App Password  
3. Load contacts from CSV  
4. Choose or customize a template  
5. Click **Send Emails**  

## ⚠️ Notes
- Ensure internet connection is active
- Invalid emails will be skipped automatically
- History is saved in `history.json`

