import streamlit as st
import pandas as pd
import datetime
import os

file_path = "tasks.xlsx"

if not os.path.exists(file_path):
    df = pd.DataFrame(columns=["Name", "Task Name", "Task Description", "Recipient Email", "Task DateTime", "Status"])
    df.to_excel(file_path, index=False)

st.title("Automated Task Reminder")

name = st.text_input("Enter Your Name")
task_name = st.text_input("Task Name")
task_description = st.text_area("Task Description")
email = st.text_input("Recipient Email")

date = st.date_input("Enter Task date")
time = st.time_input("Enter Task time", datetime.time(14,00))
task_datetime = datetime.datetime.combine(date, time)

if st.button("Submit Task", type="primary"):
    new_task = pd.DataFrame(
        [[name, task_name, task_description, email, task_datetime.strftime("%Y-%m-%d %H:%M"), "Pending"]],
        columns=["Name", "Task Name", "Task Description", "Recipient Email", "Task DateTime", "Status"]
    )
    existing = pd.read_excel(file_path)
    updated = pd.concat([existing, new_task], ignore_index=True)
    updated.to_excel(file_path, index=False)
    st.success("Task submitted successfully!")