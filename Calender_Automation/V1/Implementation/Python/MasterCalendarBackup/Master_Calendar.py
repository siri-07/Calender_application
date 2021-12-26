
# importing the dependencies

from string import ascii_uppercase
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment, Border, Side
import numpy as np

#Taking data from GUI
Initiative = "GENESIS"
OutMonth = "July"


# Extracting data from the input calender which is in the day wise format
InputDataframe = pd.read_excel(r"C:\Users\vv972\OneDrive\Documents\MATLAB\Excel case study\Excel_Automation_Test\Automation_Sample Calender_v0.6.xlsx", sheet_name='Sample_GENESIS')
InputDataframe.columns = ['Month', 'Date', 'Day', 'Course Code', 'Module','Lead1', 'Lead2', 'Lead3', 'Session Slot', 'Session Time','Comments']
InputDataframe = InputDataframe.drop([0, 1])
Date = InputDataframe['Date']
Date = Date.dropna()
Date.index = Date.index - 1
Month = InputDataframe['Month']
Month = set(Month.dropna())
Day = InputDataframe['Day']
Day = Day.dropna()
Day.index = Day.index - 1
CourseCode = InputDataframe['Course Code']
CourseCode.index = CourseCode.index - 1
Module = InputDataframe['Module']
Module.index = Module.index - 1
Lead1 = InputDataframe['Lead1']
Lead1.index = Lead1.index - 1
Lead2 = InputDataframe['Lead2']
Lead2.index = Lead2.index - 1
Lead3 = InputDataframe['Lead3']
Lead3.index = Lead3.index - 1
SessionSlot = InputDataframe['Session Slot']
SessionSlot.index = SessionSlot.index - 1
Comments = InputDataframe['Comments']
Comments.index = Comments.index - 1


# Extracting data from the Keys sheet of the Master calendar
KeysDataframe = pd.read_excel(r"C:\Users\vv972\OneDrive\Documents\MATLAB\Calender auomation product\Product_Calender_Automation\V1\Implementation\MasterCalendar/Master.xlsx", sheet_name='Key')
KeysDataframe.columns = ["FixedInitiativeTitles", "FixedInitiativeCodes", "FixedInitiativeColourCodes", "VarName4", "VarName5", "VarName6", "FixedCourseCodes", "FixedCourseTitles"]
KeysDataframe = KeysDataframe.drop(["VarName4", "VarName5", "VarName6"], axis=1)
FixedInitiativeTitles = KeysDataframe['FixedInitiativeTitles']
FixedInitiativeTitles = FixedInitiativeTitles.dropna()
FixedInitiativeTitles.index = FixedInitiativeTitles.index + 1
FixedInitiativeCodes = KeysDataframe['FixedInitiativeCodes']
FixedInitiativeCodes = FixedInitiativeCodes.dropna()
FixedInitiativeCodes.index = FixedInitiativeCodes.index + 1
FixedInitiativeColourCodes = KeysDataframe['FixedInitiativeColourCodes']                                             # reads empty data
FixedCourseCodes = KeysDataframe['FixedCourseCodes']
FixedCourseCodes = FixedCourseCodes.dropna()
FixedCourseCodes.index = FixedCourseCodes.index + 1
FixedCourseTitles = KeysDataframe['FixedCourseTitles']
FixedCourseTitles = FixedCourseTitles.dropna()
FixedCourseTitles.index = FixedCourseTitles.index + 1
#FixedCourses = [FixedCourseCodes,FixedCourseTitles]
#FixedCourses = pd.concat([FixedCourseCodes, FixedCourseTitles], axis = 1)                                            # Fixed Courses dataframe


#print(FixedCourses)
#print(CourseCode)
# print(Date, Day, CourseCode, Module, SessionSlot, Lead1, Lead2, Lead3)
# print(FixedInitiativeTitles, FixedInitiativeCodes, FixedCourseCodes, FixedCourseTitles)
# TO DO fix errors
# ExistingDataframe = pd.read_excel('/Users/achu/Downloads/Calendar/Master.xlsx', sheet_name=OutMonth)
# ExistingDataframe = ExistingDataframe.drop([0,1])
# ExistingDataframe.index = ExistingDataframe.index -1
# UniqueCourseCode = ExistingDataframe.iloc[:,0]
# RespectiveCourseTitleOutMonth = ExistingDataframe.iloc[:,1]
# RespectiveFacultyOutMonth = ExistingDataframe.iloc[:2:6]
# TimeTableOutMonth = ExistingDataframe.iloc[:,7:68]
# print(UniqueCourseCode, RespectiveCourseTitleOutMonth, RespectiveFacultyOutMonth ,TimeTableOutMonth)
#print(CourseCode)
#print(CourseCode[1])
#print(FixedCourseCodes)
#print(FixedCourses.iloc[1,0])




"""Error correction :
if course code incorrect , course title correct         =   corrects course code
if course code correct   , course title  incorrect      =   corrects course title
if course code incorrect , course title also incorrect  =   replaces the course code with ""
"""

"""Fixing the error course codes"""
for i in range(1, len(CourseCode)+1):
    TempFlag=0
    for j in range(1, len(FixedCourseCodes)+1):
        if (CourseCode[i] == FixedCourseCodes[j]):
            TempFlag = 1
    if TempFlag == 0:
        TempFlagError = 1
        for k in range(1, len(FixedCourseTitles)+1):
            if (Module[i] == FixedCourseTitles[k]):
                CourseCode[i] = FixedCourseCodes[k]
                TempFlagError = 0
        if TempFlagError == 1:
            CourseCode[i] = ""

"Fixing the error course titles"
for i in range(1, len(Module)+1):
    TempFlag=0
    for j in range(1, len(FixedCourseTitles)+1):
        if (Module[i] == FixedCourseTitles[j]):
            TempFlag = 1
    if TempFlag == 0:
        TempFlagError = 1
        for k in range(1, len(FixedCourseCodes)+1):
            if (CourseCode[i] == FixedCourseCodes[k]):
                Module[i] = FixedCourseTitles[k]
                TempFlagError = 0
        if TempFlagError == 1:
            CourseCode[i] = ""



"""Selecting the particular initaitive code"""
InitiativeCode = 11
for i in range(1, len(FixedInitiativeTitles) + 1):
    if Initiative == FixedInitiativeTitles[i]:
        InitiativeCode = FixedInitiativeCodes[i]



"""UniqueCourseCode containing unique data for CourseCode"""
UniqueCourseCode = []
for i in range(1, len(CourseCode)+1):
    if CourseCode[i] != '' and CourseCode[i] not in UniqueCourseCode:
        UniqueCourseCode.append(CourseCode[i])
UniqueCourseCode=pd.Series(UniqueCourseCode)
UniqueCourseCode.index = UniqueCourseCode.index + 1



"""Declaring variable to hold respective CourseTitle for UniqueCourseCode"""
RespectiveCourseTitle = pd.Series([""]*len(UniqueCourseCode))
RespectiveCourseTitle.index = RespectiveCourseTitle.index + 1
print(RespectiveCourseTitle)


"""Declaring matrix to hold repeatitive list of faculties for respective UniqueCourseCode"""
Faculty =pd.DataFrame([[""]*len(CourseCode)*3]*len(UniqueCourseCode))
print(Faculty)
Faculty.index = Faculty.index + 1
Faculty.iloc[1,0] = "Vivek"
print(Faculty)