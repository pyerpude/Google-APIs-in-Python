# Google-APIs-in-Python
Learn about how to use Google Sheet and Google Drive API's in python

Functionalities:
1. Traverse into a Google Drive or a Team Drive containing subfolders with Google objects like a doc, sheet etc
2. Downloads data from multiple google sheets into your local machine as a csv which then can be used to load into a database
3. Given similar google sheets, if all of them are in the same format i.e tab names, this script will traverse along all of sheets downloading those specific tabs and consolidating into a single csv.
For example, if there are 5 similar google sheets having 7 tabs(different sheets) and you need to download only two tabs of all of them, this script will download the tabs and consolidate them into 2 different csvs'
4. Works on a GUID of Google Drive/team drive folder or a single google sheet too

Command to use: 
python scanDriveFolder_google_sheet.py <source_drive_id> <dest_server_path> <tab_name or prefix> <remove_flag> <archive_flag>

Examples: 
1. Single google sheet 
  - Included in the images attached 
  - Input is a GUID of that google sheet
  - The result is a Sheet1.csv stored in destination directory
   
2. Multiple google sheets stored in a drive/team drive
  - Included in the images attached
  - Input is the google drive/team drive folder ID
  - The result is a consolidated csv called Sheet1.csv which has 'Sheet1' data from all google sheets

# Important Notes: 
- This python script is meant to download data from google sheets which has multiple tabs/sheets inside a single google sheet
- You can specify what tab to download with a comma separated values like Sheet1, Sheet2 and it will create two csvs as a result i.e Sheet1.csv and Sheet2.csv
- Tab name/Sheet name is mandatory in the input since a single google sheet can contain multiple tabs and you might want to download only specific tabs' data
- The script contains creation of manifest file in a destination directory, it contains the file path with the file name of the destination, so you can automate a load of this csv to any other destination like a google cloud storage bucket or an s3 bucket. The next process can then read manifest file and pick up those files from local machine/server and move it to your ultimate destination.
