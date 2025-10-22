# Birthday-Email-Bot-Automation
An automated solution for sending beautifully designed birthday greetings via Microsoft Outlook using employee data, personalized cards, and custom image placements.

---

## 📌 Overview

**BirthdayBot** automatically:
- Reads employee birthday data from an Excel file
- Generates personalized birthday cards with employee photos and names placed on a custom background
- Sends the card through Microsoft Outlook to employees whose birthdays are today
- BCCs a fixed recipient (e.g., HR or Operations)

---

## 🚀 Features

- 💼 Outlook email automation with embedded birthday cards
- 📅 Dynamically reads birthdays from `employees.xlsx`
- 🖼️ Customizable background template (`bg1.png`)
- 👤 Dynamic image and name placement using `NewplacementsTry.json`
- 📐 Final card is auto-resized to 65% of the original size
- ⏰ Can be scheduled to run daily at 11:00 AM IST via Task Scheduler and `.bat` file

---

## 🗂️ Directory Structure
```text
📁 My Birthday Trial/
├── SendBirthdayEmail.py           # Main script to generate and send emails
├── employees.xlsx                 # Excel file with Name, Email, DOB, Photo
├── employee_photos/              # Folder with employee images (JPG/PNG)
├── NewplacementsTry.json         # Layout config for 1–8 employees
├── bg1.png                        # Birthday card background template
├── run_birthday_mailer.bat       # Optional: For scheduling via Task Scheduler
```


---

## 📊 Excel Format (`employees.xlsx`)

| Name        | Email               | DOB        | Photo     |
|-------------|---------------------|------------|-----------|
| Jane Doe    | jane@company.com    | 1994-07-17 | jane.jpg  |
| John Smith  | john@company.com    | 1990-07-17 | john.jpg  |

> 📌 Ensure the `DOB` is in `YYYY-MM-DD` format.  
> 🖼️ Photo filenames must exactly match the image names in the `employee_photos/` folder.

---

## 🛠️ Requirements

- **Python**: 3.9 or above
- **Microsoft Outlook**: Installed and configured on the machine

### 📦 Required Python Packages

Install the following dependencies using `pip`:

```bash
pip install pandas pillow pywin32 openpyxl

> ✅ **Ensure** Outlook is set as the default email client and is accessible through the desktop app for proper automation.
```
---

## How to Run

To run the script manually:

```bash
python SendBirthdayEmail.py
```
If today's date matches any birthday(s) in `employees.xlsx`, a birthday card will be created and an email draft will open in Outlook.

### The email will:
- **To**: All birthday recipients (from the Excel sheet)
- 📎 Embed the birthday card directly in the email body

---

## ⏰ Automate Daily at 11 AM (Optional)

Create a `.bat` file (already included as `run_birthday_email_Script.bat`):

```bat
cd /d "C:\Path\To\Your\Project"
python SendBirthdayEmail_Final.py
```
Then schedule it via **Windows Task Scheduler**:

- **Trigger**: Daily at **11:00 AM**
- **Action**: Run the above `.bat` file
- ✅ Ensure Microsoft Outlook is running or set to open automatically

---

## 🧠 Future Improvements

- 📅 Integrate with Outlook calendar or LDAP directory  
- ⏳ Add support for upcoming birthday reminders  
- 💬 Push birthday wishes to Slack or Microsoft Teams  
- 📤 Support for other email services like Gmail or SMTP  

---

## 👨‍💼 Built for Internal Use

This bot streamlines birthday wishes within the team and enhances employee engagement by sending timely and personalized greetings.


