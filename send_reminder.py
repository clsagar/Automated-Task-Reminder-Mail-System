import pandas as pd
import smtplib
from email.message import EmailMessage
from dotenv import load_dotenv
import os
from datetime import datetime

load_dotenv("credentials.env")

SENDER_EMAIL = os.getenv("SENDER_EMAIL")
SENDER_PASSWORD = os.getenv("SENDER_PASSWORD")

file_path = "tasks.xlsx"

if not os.path.exists(file_path):
    print("Task file not found.")
    exit()

df = pd.read_excel(file_path)
now = datetime.now().strftime('%Y-%m-%d %H:%M')

for i, row in df.iterrows():
    task_time = row["Task DateTime"]
    status = row.get("Status", "Pending")

    if status == "Completed":
        continue

    if isinstance(task_time, pd.Timestamp):
        task_time = task_time.strftime('%Y-%m-%d %H:%M')

    if task_time == now:
        msg = EmailMessage()
        msg['Subject'] = row['Task Name']
        msg['From'] = SENDER_EMAIL
        msg['To'] = row['Recipient Email']
        msg.set_content(row['Task Description'])

        try:
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(SENDER_EMAIL, SENDER_PASSWORD)
                smtp.send_message(msg)
            print(f"Reminder sent to {row['Recipient Email']}")

            df.at[i, 'Status'] = 'Completed'
        except Exception as e:
            print(f"Error sending to {row['Recipient Email']}: {e}")

df.to_excel(file_path, index=False)