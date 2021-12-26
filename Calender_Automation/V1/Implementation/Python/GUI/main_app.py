"""
App for login
Created by: Govind Bansal
Version 0.6
Date: 30/09/2021
"""
import hashlib
import sqlite3
import datetime
# import os

import streamlit as st
# import streamlit.components.v1 as components
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
# from xlsx2html import xlsx2html

import FacultyCalendarFunction
import FacultyLoadSheet
import MasterCalendarFunction
import UpdateMasterCalendarAsInput

gauth = GoogleAuth()
gauth.LoadCredentialsFile("credentials.txt")

drive = GoogleDrive(gauth)


# Security
def make_hashes(password):
    """Function for password encoding"""
    return hashlib.sha256(str.encode(password)).hexdigest()


def check_hashes(password, hashed_text):
    """Function for checking hashes"""
    if make_hashes(password) == hashed_text:
        return hashed_text
    return False


# Database creation for faculty details
faculty_user = sqlite3.connect('data.db')
faculty_cursor = faculty_user.cursor()

# Database creation for admin details
admin_user = sqlite3.connect('data2.db')
admin_cursor = admin_user.cursor()


def create_usertable_faculty():
    """Creating table for storing ID and password for faculty"""
    faculty_cursor.execute('CREATE TABLE IF NOT EXISTS userstable(username TEXT, password TEXT)')


def add_userdata_faculty(username, password):
    """Adding new set of ID and password for faculty"""
    faculty_cursor.execute('INSERT INTO userstable(username, password) VALUES (?,?)', (username, password))
    faculty_user.commit()


def login_user_faculty(username, password):
    """Function to login for faculty"""
    faculty_cursor.execute('SELECT * FROM userstable WHERE username = ? AND password = ?', (username, password))
    data = faculty_cursor.fetchall()
    return data


def view_all_users_faculty():
    """Function to fetch all users from DB for faculty"""
    faculty_cursor.execute('SELECT * FROM userstable')
    data = faculty_cursor.fetchall()
    return data


def create_usertable_admin():
    """Creating table for storing ID and password for admin"""
    admin_cursor.execute('CREATE TABLE IF NOT EXISTS userstable(username TEXT, password TEXT)')


def add_userdata_admin(username, password):
    """Adding new set of ID and password for admin"""
    admin_cursor.execute('INSERT INTO userstable(username, password) VALUES (?,?)', (username, password))
    admin_user.commit()


def login_user_admin(username, password):
    """Function to login for admin"""
    admin_cursor.execute('SELECT * FROM userstable WHERE username = ? AND password = ?', (username, password))
    data = admin_cursor.fetchall()
    return data


def view_all_users_admin():
    """Function to fetch all users from DB for admin"""
    admin_cursor.execute('SELECT * FROM userstable')
    data = admin_cursor.fetchall()
    return data


def greeting():
    current_time = datetime.datetime.now()
    current_hour = current_time.hour
    if current_hour < 12:
        return "Good Morning!"
    elif 12 <= current_hour < 16:
        return "Good Afternoon!"
    else:
        return "Good Evening!"


# main function
def main():
    """GUI LOGIN PAGE"""
    st.set_page_config(page_title="GEA Calendar")
    st.image("resource/ltts.png")
    st.subheader(f"{greeting()} Welcome to GEA calendar automation app")
    st.sidebar.title("GEA calendar automation")
    st.sidebar.image("resource/gea.jpg")
    info = st.sidebar.selectbox("Info:", ["Select", "Help", "About"])
    if info == "Help":
        st.subheader("Help section")
        st.write("This automation app helps in creating master calendar, faculty calendar and faculty loadsheet")
        st.write("Python is used to create this app. Libraries like pandas, openpyxl, pydrive, streamlit are widely "
                 "used for the app.")
        st.write("To deploy the app, heroku servers are used.")
        st.write("To view source code, you can visit the following github link:")
        st.write("[Github Link](https://github.com/GENESIS2021Q1/Product_Calender_Automation.git)")
    elif info == "About":
        st.subheader("About the team")
        st.write("This app is created by a team of 5. The team consists of:")
        st.write("[Dr. Prithvi Sekhar Pagala](https://www.linkedin.com/in/pspphd/)")
        st.write("[Vivek Ashar](https://www.linkedin.com/in/vivek-ashar-82940b209/)")
        st.write("[Priyadharshni N](https://www.linkedin.com/in/priyadharshni-n-205b351a4/)")
        st.write("[Archana Arun](https://www.linkedin.com/in/archana-arun-5959021b1/)")
        st.write("[Govind Bansal](https://www.linkedin.com/in/govind-bansal13/)")
        st.write("")
        st.write("Special thanks to Kartik Mudaliar and Srinivas K for their valuable feedback.")
    st.sidebar.markdown("____")
    # Getting Credentials from user
    st.sidebar.subheader("Login using your credentials")
    username = st.sidebar.text_input("User Name")
    password = st.sidebar.text_input("Password", type='password')
    ac_type = st.sidebar.selectbox("Account Type", ["Faculty", "Admin"])
    if ac_type == "Admin":
        # If account type is admin, provide special privileges
        if st.sidebar.checkbox("Login"):
            # Checking credentials
            create_usertable_admin()
            hashed_password = make_hashes(password)
            result = login_user_admin(username, check_hashes(password, hashed_password))
            if result:
                # If credentials are correct, display menu
                st.sidebar.success("Logged in as {}".format(username))
                st.subheader("Admin Panel")
                menu = st.selectbox("Menu", ["Select", "Update calendar", "Update using existing master calendar",
                                             "View and Download calendar"])
                # if menu == "View calendar":
                #     # Action to view calendar
                #     month_view = st.selectbox("Select Month", ["Select", "January", "February", "March", "April",
                #                                                "May", "June", "July", "August",
                #                                                "September", "October", "November", "December"])
                #
                #     # Master Calendar fetching and displaying
                #     mc = drive.ListFile({'q': "title contains 'Master' and trashed=false"}).GetList()
                #     mc_file_id = mc[0]['id']
                #     mc_file = drive.CreateFile({'id': mc_file_id})
                #     mc_file_title = 'UpdatedMasterCalendar.xlsx'
                #     mc_file.GetContentFile(mc_file_title,
                #                            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                #     if month_view != "Select":
                #         st.subheader("Master Calendar")
                #         out_stream_mc = xlsx2html(filepath=mc_file_title,
                #                                   sheet=month_view,
                #                                   output=(os.getcwd() + "/" + mc_file_title + ".html"),
                #                                   default_cell_border="10")
                #         out_stream_mc.seek(0)
                #         html_file_mc = open(mc_file_title + ".html", 'r', encoding='utf-8')
                #         source_code_mc = html_file_mc.read()
                #         components.html(source_code_mc, height=500, width=700, scrolling=True)
                #
                #     # Faculty Calendar fetching and displaying
                #     fc = drive.ListFile({'q': "title contains 'FacultyCalendar' and trashed=false"}).GetList()
                #     fc_file_id = fc[0]['id']
                #     fc_file = drive.CreateFile({'id': fc_file_id})
                #     fc_file_title = 'UpdatedFacultyCalendar.xlsx'
                #     fc_file.GetContentFile(fc_file_title,
                #                            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                #     if month_view != "Select":
                #         st.subheader("Faculty Calendar")
                #         out_stream_fc = xlsx2html(filepath=fc_file_title,
                #                                   sheet=month_view,
                #                                   output=(os.getcwd() + "/" + fc_file_title + ".html"),
                #                                   default_cell_border=10)
                #         out_stream_fc.seek(0)
                #         html_file_fc = open(fc_file_title + ".html", 'r', encoding='utf-8')
                #         source_code_fc = html_file_fc.read()
                #         components.html(source_code_fc, height=500, width=700, scrolling=True)
                #
                #     # Faculty Loadsheet fetching and displaying
                #     fl = drive.ListFile({'q': "title contains 'FacultyLoad' and trashed=false"}).GetList()
                #     fl_file_id = fl[0]['id']
                #     fl_file = drive.CreateFile({'id': fl_file_id})
                #     fl_file_title = 'UpdatedFacultyLoadSheet.xlsx'
                #     fl_file.GetContentFile(fl_file_title,
                #                            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                #     if month_view != "Select":
                #         st.subheader("Faculty Loadsheet")
                #         out_stream_fl = xlsx2html(filepath=fl_file_title,
                #                                   sheet=month_view,
                #                                   output=(os.getcwd() + "/" + fl_file_title + ".html"),
                #                                   default_cell_border=10)
                #         out_stream_fl.seek(0)
                #         html_file_fl = open(fl_file_title + ".html", 'r', encoding='utf-8')
                #         source_code_fl = html_file_fl.read()
                #         components.html(source_code_fl, height=500, width=700, scrolling=True)

                if menu == "Update calendar":
                    # Action to update calendar
                    file = st.file_uploader("Upload input file", type=['.xlsx'])
                    st.write("To download input template, visit below link:")
                    st.write("[Template link](https://drive.google.com/drive/folders/1QyJvL8i1bLX_qZIVCiRo4O_jjeuO66vp?usp=sharing)")
                    year = st.selectbox("Select Year", ["Select", "2021", "2022"])
                    month = st.selectbox("Select Month", ["Select", "January", "February", "March", "April",
                                                          "May", "June", "July", "August",
                                                          "September", "October", "November", "December"])
                    initiative = st.selectbox("Select Initiative", ["GENESIS", "GENESIS PRO",
                                                                    "BUILD / STEP UP", "OPEN TRAININGS",
                                                                    "STEPin", "OTHERS"])
                    if st.button("Update calendar"):
                        MasterCalendarFunction.MasterCalendarFunction("Test_vector", file, initiative, month)
                        FacultyCalendarFunction.FacultyCalendarFunction()
                        FacultyLoadSheet.FacultyLoadSheetFunction()

                    if file is not None:
                        file_details = {"filename": file.name, "filetype": file.type}
                        print(file_details)

                elif menu == "Update using existing master calendar":
                    # Action to update master calendar using existing master calendar
                    file_existing_cal = st.file_uploader("Upload existing master calendar file", type=['.xlsx'])
                    month_existing_cal = st.selectbox("Select Month",
                                                      ["Select", "January", "February", "March", "April",
                                                       "May", "June", "July", "August",
                                                       "September", "October", "November", "December"])
                    if st.button("Update calendar"):
                        UpdateMasterCalendarAsInput.UpdateMasterCalendarAsInput(file_existing_cal, month_existing_cal)
                        FacultyCalendarFunction.FacultyCalendarFunction()
                        FacultyLoadSheet.FacultyLoadSheetFunction()
                    if file_existing_cal is not None:
                        file_details = {"filename": file_existing_cal.name, "filetype": file_existing_cal.type}
                        print(file_details)

                elif menu == "View and Download calendar":
                    # Downloading calendar
                    st.write("To download the calendar, please visit the below link.")
                    st.write("[Google drive link](https://drive.google.com/drive/folders/13RC_H50afz9NVmyI2xaqGzgOJhGFn3US?usp=sharing)")
                    st.write("Note: This link contains all files updated till date, choose the latest file according"
                             " to timestamp.")
                # elif menu == "Add new user":
                #     # Action to sign up new user
                #     st.subheader("New user registration")
                #     new_user = st.text_input("Enter username")
                #     new_password = st.text_input("Enter password", type='password')
                #     new_ac_type = st.selectbox("Account type", ["Admin", "Faculty"])
                #     if new_ac_type == "Faculty":
                #         if st.button("Add new faculty"):
                #             create_usertable_faculty()
                #             add_userdata_faculty(new_user, make_hashes(new_password))
                #             st.success("You have successfully created a valid account")
                #     elif new_ac_type == "Admin":
                #         if st.button("Add new admin"):
                #             create_usertable_admin()
                #             add_userdata_admin(new_user, make_hashes(new_password))
                #             st.success("You have successfully created a valid account")
            else:
                # If the credentials are wrong
                st.sidebar.warning("Invalid username/password")

    elif ac_type == "Faculty":
        # If account type is faculty, allow only viewing of calendar
        if st.sidebar.checkbox("Login"):
            # Checking credentials
            create_usertable_faculty()
            hashed_password = make_hashes(password)
            result = login_user_faculty(username, check_hashes(password, hashed_password))
            if result:
                # If credentials are correct, display calendar
                st.sidebar.success("Logged in as {}".format(username))
                st.subheader("Faculty view")
                # menu = st.selectbox("Menu", ["View calendar", "Download calendar"])
                # if menu == "View calendar":
                #     month_view = st.selectbox("Select Month", ["Select", "January", "February", "March", "April",
                #                                                "May", "June", "July", "August",
                #                                                "September", "October", "November", "December"])
                #     # Master Calendar fetching and displaying
                #     mc = drive.ListFile({'q': "title contains 'Master' and trashed=false"}).GetList()
                #     mc_file_id = mc[0]['id']
                #     mc_file = drive.CreateFile({'id': mc_file_id})
                #     mc_file_title = 'UpdatedMasterCalendar.xlsx'
                #     mc_file.GetContentFile(mc_file_title,
                #                            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                #     if month_view != "Select":
                #         st.subheader("Master Calendar")
                #         out_stream_mc = xlsx2html(filepath=mc_file_title,
                #                                   sheet=month_view,
                #                                   output=(os.getcwd() + "/" + mc_file_title + ".html"),
                #                                   default_cell_border="10")
                #         out_stream_mc.seek(0)
                #         html_file_mc = open(mc_file_title + ".html", 'r', encoding='utf-8')
                #         source_code_mc = html_file_mc.read()
                #         components.html(source_code_mc, height=500, width=700, scrolling=True)
                #
                #     # Faculty Calendar fetching and displaying
                #     fc = drive.ListFile({'q': "title contains 'FacultyCalendar' and trashed=false"}).GetList()
                #     fc_file_id = fc[0]['id']
                #     fc_file = drive.CreateFile({'id': fc_file_id})
                #     fc_file_title = 'UpdatedFacultyCalendar.xlsx'
                #     fc_file.GetContentFile(fc_file_title,
                #                            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                #     if month_view != "Select":
                #         st.subheader("Faculty Calendar")
                #         out_stream_fc = xlsx2html(filepath=fc_file_title,
                #                                   sheet=month_view,
                #                                   output=(os.getcwd() + "/" + fc_file_title + ".html"),
                #                                   default_cell_border=10)
                #         out_stream_fc.seek(0)
                #         html_file_fc = open(fc_file_title + ".html", 'r', encoding='utf-8')
                #         source_code_fc = html_file_fc.read()
                #         components.html(source_code_fc, height=500, width=700, scrolling=True)
                #
                #     # Faculty Loadsheet fetching and displaying
                #     fl = drive.ListFile({'q': "title contains 'FacultyLoad' and trashed=false"}).GetList()
                #     fl_file_id = fl[0]['id']
                #     fl_file = drive.CreateFile({'id': fl_file_id})
                #     fl_file_title = 'UpdatedFacultyLoadSheet.xlsx'
                #     fl_file.GetContentFile(fl_file_title,
                #                            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                #     if month_view != "Select":
                #         st.subheader("Faculty Loadsheet")
                #         out_stream_fl = xlsx2html(filepath=fl_file_title,
                #                                   sheet=month_view,
                #                                   output=(os.getcwd() + "/" + fl_file_title + ".html"),
                #                                   default_cell_border=10)
                #         out_stream_fl.seek(0)
                #         html_file_fl = open(fl_file_title + ".html", 'r', encoding='utf-8')
                #         source_code_fl = html_file_fl.read()
                #         components.html(source_code_fl, height=500, width=700, scrolling=True)

                # elif menu == "Download calendar":
                # Downloading calendar
                st.write("To download the calendar, please visit the below link.")
                st.write("https://drive.google.com/drive/folders/13RC_H50afz9NVmyI2xaqGzgOJhGFn3US?usp=sharing")
                st.write("Note: This link contains all files updated till date, choose the latest file according to "
                         "timestamp")
            else:
                # If the credentials are wrong
                st.sidebar.warning("Invalid username/password")


if __name__ == "__main__":
    main()
