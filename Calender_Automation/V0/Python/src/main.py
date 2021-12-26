"""
Main file
Importing function files, creating objects and calling functions
"""
import tkinter as tk
from openpyxl import Workbook
import calendar_func
import GUI
from faculty_calendar import create_faculty_calendar
import ifls

if __name__ == '__main__':
    cal = calendar_func.MasterCalendar()  # created object of Master Calendar
    aut = GUI.Automation()  # created object of gui class
    root = tk.Tk()  # created window
    aut.fun1(root)
    file = GUI.filepath(root)
    SUCCESS = cal.create_master_cal(file)

    inp = Workbook()
    FAC_CAL = create_faculty_calendar(file, inp)
    if FAC_CAL == 1:
        print("Faculty sheet created")
    if SUCCESS == 1:
        print("Master Calendar Created")
    TEMP = ifls.create_ifls(file, inp)
    if TEMP == 1:
        print("Integrated Load Sheet Created")

    root.mainloop()
