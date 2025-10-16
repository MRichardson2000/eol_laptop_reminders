💻 EOL Laptop Reminder Tool
This script automates the process of identifying laptops nearing their end-of-life (EOL) and sends reminder emails to the helpdesk team. It helps ensure timely replacements and support.

-- THIS IS ON EBSFS005 (PRTG Server) --

📦 Features
- Scans a CSV file containing laptop inventory and EOL dates.
- Identifies laptops due for refresh within the next 90 days.
- Prevents duplicate notifications using a log file.
- Sends a summary email via Outlook to the helpdesk.

🧰 Requirements
- Python 3.8+
- pandas
- openpyxl
- pywin32 (for Outlook integration)
Install dependencies:
uv add pandas openpyxl pywin32 
<!-- run each one of the above individually -->



📁 File Structure
project/
│
├── main.py                # Main script
├── config.py              # Contains CSV_FILE and LOG_FILE paths
├── notified_devices.log   # Tracks already notified laptops
└── README.md              # Documentation



⚙️ Configuration
Edit config.py to set your file paths if you need to change them but they're constants so should only need changing for testing:
CSV_FILE = "K:/IT/Restricted/Ecology Network/devops_automation_mr/eol_laptop_reminders/eol_laptops.xlsx"
LOG_FILE = "K:/IT/Restricted/Ecology Network/devops_automation_mr/eol_laptop_reminders/notified_devices.txt"



🚀 Usage
Set up a batch file with the below, put your user path in and then specify on your c drive where you cloned the repo. This is just where most of my stuff went:
C:\Users\YOURUSERPATHHERE\AppData\Local\Programs\Python\Python313\python.exe C:\Utilities\Python\eol_laptop_refresh_reminders\main.py
Then set up a schedule task to run once a week on your chosen day and time that targets this batch file and runs it. 


This will:
- Read the laptop inventory file on the K drive.
- Check for laptops with EOL dates within 90 days.
- Skip devices already logged.
- Send an email to the helpdesk with the list of upcoming refreshes.

✉️ Email Setup
The script uses win32com.client to send emails via Outlook. The recipient is currently set to:
mail.To = "helpdesk@ecology.co.uk"


You can change this to any valid email address or distribution list.

🧪 Testing
To test with alternate files you just pass in your file paths. It uses the constants if nothings passed in:
eol_laptops(csv_path="test_inventory.xlsx", log_path="test_log.log")



📌 Notes
- The script assumes the EOL date is in the third column of the spreadsheet.
- The first row is skipped (assumed to be headers).
- Only laptops not previously notified will trigger a new reminder.
- run a uv sync in the terminal to pull the dependencies from the toml
- if you need uv run pip install uv in the terminal to add to your global scope
