import win32com.client
import pandas as pd
from datetime import datetime as dt, timedelta as td
import os
from config import XLSX_FILE, LOG_FILE


def eol_laptops(xlsx_path: str = XLSX_FILE, log_path: str = LOG_FILE) -> None:
    """
    This function has two parameters passed in but they use constants. You don't need to pass anything in when you call it unless you want to change
    which path is being used, maybe for testing for example. The log file is used to record which laptops have already been picked up and emailed through
    to prevent duplicates. We use pandas to read the xlsx file (we skip the first row as they're headers) and then we set the threshhold day which is 90 days.
    This gives us enough time to organise a replacement laptop for them. We check if the log path exists and if it does we open it as read and then we
    establish a variable called notified devices that's equal to each line in the file. Each line is a computer name. If it doesn't exist we establish an empty
    set. We then set up and empty list for the reminders and the newly notified devices. Then we iterate through the excel document using pandas. We're targetting the eol
    date column so if index is equal to 0 we just skip and continue. We then read the dates in the column and if they are within 90 days and if they're
    not already in the log file then it appends the computer name, laptop refresh due on, and then the date formatted nicely. Then we append to newly
    notified the computer names to ensure we don't get duplicates to the helpdesk. Otherwise it returns None but I've added a print statement just saying
    no laptops are due for a refresh. When this runs, this function is called which then calls the send reminder function.
    """
    df = pd.read_excel(xlsx_path, engine="openpyxl")
    today = dt.today()
    reminder_threshold = today + td(days=90)
    if os.path.exists(log_path):
        with open(log_path, "r") as file:
            notified_devices = set(line.strip() for line in file)
    else:
        notified_devices = set()
    reminders = []
    newly_notified = []
    for index, row in df.iterrows():
        if index == 0:
            continue
        next_due_date = row[2]
        if isinstance(next_due_date, pd.Timestamp):
            next_due_date = next_due_date.to_pydatetime()
        if pd.notna(next_due_date) and today <= next_due_date <= reminder_threshold:
            computer_name = row[0]
            if computer_name in notified_devices:
                pass
            else:
                reminders.append(
                    f"{computer_name}: Laptop Refresh due on {next_due_date.strftime('%Y-%m-%d')}"
                )
                newly_notified.append(computer_name)
    if reminders:
        send_reminder(reminders)
        with open(log_path, "a") as f:
            for device in newly_notified:
                f.write(device + "\n")
    else:
        print("No Laptop Refreshes due within the next 90 days.")


def send_reminder(reminder_list: list) -> None:
    """
    This is the function that sends the reminder email. It uses the win32com module and as you can see below you specify the email, subject and
    body and then it sends the email, we're currently sending this to helpdesk for centralisation but this can be changed if required.
    """
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.Subject = "EOL Laptop Reminders"
    mail.To = "helpdesk@ecology.co.uk"
    mail.Body = (
        "The following Laptops are due to be replaced within the next 90 days:\n\n"
        + "\n".join(reminder_list)
    )
    mail.Send()
    print("Reminder email sent.")


if __name__ == "__main__":
    eol_laptops()
