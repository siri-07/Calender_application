from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

gauth = GoogleAuth()
gauth.LoadCredentialsFile("credentials.txt")

drive = GoogleDrive(gauth)

file1 = drive.CreateFile({"mimeType": "application/vnd.ms-excel"})
file1.SetContentFile("MasterCalendar.xlsx")
file1.Upload({"convert": True})


file2 = drive.CreateFile({"mimeType": "application/vnd.ms-excel"})
file2.SetContentFile("FacultyCalendar.xlsx")
file2.Upload({"convert": True})

file3 = drive.CreateFile({"mimeType": "application/vnd.ms-excel"})
file3.SetContentFile("FacultyLoadSheet.xlsx")
file3.Upload({"convert": True})

file4 = drive.CreateFile({"mimeType": "application/vnd.ms-excel"})
file4.SetContentFile("Keys.xlsx")
file4.Upload({"convert": True})
