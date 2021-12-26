"""
Template for google drive operations like upload, download, search for files
"""
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from datetime import datetime

gauth = GoogleAuth()
gauth.LoadCredentialsFile("credentials.txt")

drive = GoogleDrive(gauth)


def list_files():
    """
    Listing all files -> insert file name after contains
    This will list all the files that have similar name with their ID
    """
    file_list = drive.ListFile({'q': "title contains 'MasterCalendar' and trashed=false"}).GetList()
    for file in file_list:
        print(file['title'], file['id'])


def download_file():
    """
    Searching for a file and downloading it -> insert file name after contains
    give file name to variable file_name
    """
    file_list = drive.ListFile({'q': "title contains 'MasterCalendar' and trashed=false"}).GetList()
    print(file_list[0]['title'])
    file_id = file_list[0]['id']
    print(file_id)
    file = drive.CreateFile({'id': file_id})
    file_title = 'UpdatedMasterCalendar'
    file.GetContentFile(file_title,
                        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


def upload_file():
    """Upload file -> change name in setcontentfile"""
    file1 = drive.CreateFile({"mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"})
    file1.SetContentFile("MasterCalendar.xlsx")
    file1.Upload({"convert": True})


if __name__ == "__main__":
    list_files()

'''
Script for saving file name with timestamp:
file_name = 'MasterCalendar'
    file_time = datetime.now().strftime(" %Y-%m-%d_%I-%M-%S_%p")
    file_format = '.xlsx'
    file_title = file_name + file_time + file_format
    
    USE THIS file_title to save file and upload
'''