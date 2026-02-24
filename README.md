# Automated Excel Email Sender (Outlook + Python)

This script automatically:

- Finds the latest Excel file in a folder
- Reads the `Status` value where `ID == 2`
- Sends an HTML email via Microsoft Outlook
- Attaches the latest Excel file
- Includes extracted Excel data inside the email body

Designed for Windows environments using Outlook Desktop.

---

## Requirements

- Windows OS
- Microsoft Outlook (Desktop version installed)
- Python 3.9+
- Excel files in `.xlsx` format
- Required Python packages:
- pandas
- pywin32
- openpyxl

# Automating with Windows Task Scheduler
You can automate this script to run daily, weekly, or at a specific time.
