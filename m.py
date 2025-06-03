import streamlit as st
import openpyxl
import os
from datetime import datetime

file_name = "data.xlsx"

if not os.path.exists(file_name):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["Name", "Age", "Job", "Email", "Timestamp"])
    wb.save(file_name)

st.set_page_config(page_title="Company Data Form", layout="centered")
st.title("Employee Data Entry")

name = st.text_input("Name")
age = st.text_input("Age")
job = st.text_input("Job Title")
email = st.text_input("Email")

if st.button("Submit"):
    if name and age and job and email:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        wb = openpyxl.load_workbook(file_name)
        ws = wb.active
        ws.append([name, age, job, email, timestamp])
        wb.save(file_name)
        st.success("Data saved successfully!")
    else:
        st.warning("Please fill in all fields.")
