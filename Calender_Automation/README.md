# Product_Calender_Automation

## Requirements for updating Master calendar with input file using standard template

# Link for template standard input template: 

https://docs.google.com/spreadsheets/d/1EWYp_1iyK2wLMfKGJOiTJAk5WexZusCP/edit?usp=sharing&ouid=113003694561146884677&rtpof=true&sd=true

* Using the template above, training schedule can be added monthwise and initiatives wise
* The name of the input excel sheet MUST be named as "Test_vector"(as shown in template)
* Along with the Test_vector sheet, "Key" sheet MUST be present under the columns assigned as in the template
* The "Key" sheet must contain all times the 6 fixed initiatives with their respective codes and total list of course code and course title in order to refer for corrections while writing to output files
* Appending additional slots for existing courses is possible by adding just the additional slots in the input file for the same course

## Requirements for updating Master calendar using Master calendar as input

# Link for template

2 Slots format - M/A : https://docs.google.com/spreadsheets/d/1jtKnXV12VE1fH20CGDo4B3uNWRTAhQCWz-hHUDWUe3I/edit?usp=sharing

4 Slots format - M1/M2/A1/A2 : https://docs.google.com/spreadsheets/d/1jVheSPZkOtfNKRNoc_858nwk2UaHCe0gExTNZfZ8vxA/edit?usp=sharing

* Any of the two templates can be used for updating Master calendar monthwise onto the drive
* The blocked slots must have the corresponding initiative code in the cell according to the key as shown in the sample data in the template
* The name of the sheet must be the name of the month to be updated
* The "Key" sheet must be present with the fixed list of initiatives and initiative code 

## App deployment

* The app is deployed on heroku servers.
* To add/modify new features, you will be required to install HEROKU CLI [link](https://devcenter.heroku.com/articles/getting-started-with-python#set-up)
* After installation, open terminal in working directory and enter the following commands:
  - "heroku git:clone -a geacalendar"
  - login using heroku credentials
* After pulling and making changes, enter the following commands to push app and deploy on server
  - git add .
  - git commit -m "commit message"
  - git push heroku master

### Additional features for V1 to do
* Update keysheet by appending new initiatives/courses list
* Check for duplicate course entries in input file
* Using built in libraries to identify number of days in month, current year and highlight weekend and holidays
* Function to remove a course schedule 
* Read multiple months data in one sheet as input file (currently takes data one by one month)


