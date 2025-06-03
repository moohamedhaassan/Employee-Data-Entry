import streamlit as st
import openpyxl
import os
from datetime import datetime

file_name = "data.xlsx"

if not os.path.exists(file_name):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet.append(["Name", "Nationality", "Department", "Position", "Iqama No", "Phone Number", "Email", "Timestamp"])
    workbook.save(file_name)

st.title("Employee Data Entry")

if "last_submission" not in st.session_state:
    st.session_state.last_submission = {
        "name": "", "nationality": "", "department": "",
        "position": "", "iqama": "", "phone": "", "email": ""
    }

with st.form("data_form"):
    name = st.text_input("Name")
    nationality = st.selectbox("Nationality", ["", "Saudi", "Egyptian", "Indian", "Pakistani", "Other"])
    department = st.selectbox("Department", ["", "HR", "IT", "Maintenance", "Finance", "Operations"])
    position = st.text_input("Position")
    iqama = st.text_input("Iqama No")
    phone = st.text_input("Phone Number")
    email = st.text_input("Email")

    submitted = st.form_submit_button("Submit")

    if submitted:
        if not all([name, nationality, department, position, iqama, phone, email]):
            st.error("Please fill in all fields!")
        elif (
            name == st.session_state.last_submission["name"] and
            nationality == st.session_state.last_submission["nationality"] and
            department == st.session_state.last_submission["department"] and
            position == st.session_state.last_submission["position"] and
            iqama == st.session_state.last_submission["iqama"] and
            phone == st.session_state.last_submission["phone"] and
            email == st.session_state.last_submission["email"]
        ):
            st.warning("You have already submitted this data. Please change something before submitting again.")
        else:
            workbook = openpyxl.load_workbook(file_name)
            sheet = workbook.active
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            sheet.append([name, nationality, department, position, iqama, phone, email, timestamp])
            workbook.save(file_name)

            st.session_state.last_submission = {
                "name": name,
                "nationality": nationality,
                "department": department,
                "position": position,
                "iqama": iqama,
                "phone": phone,
                "email": email
            }

            st.success("Data saved successfully!")
            st.experimental_rerun()
