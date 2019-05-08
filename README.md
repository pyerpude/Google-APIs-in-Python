# Google-APIs-in-Python
Learn about how to use Google Sheet and Google Drive API's in python

Functionalities:
1. Traverse into a Google Drive or a Team Drive containing subfolders with Google objects like a doc, sheet etc
2. Downloads data from multiple google sheets into your local machine
3. Given similar google sheets, if all of them are in the same format i.e tab names, this script will traverse along all of sheets downloading those specific tabs and consolidating into a single csv.
For example, if there are 5 similar google sheets having 7 tabs(different sheets) and you need to download only two tabs of all of them, this script will download the tabs and consolidate them into 2 different csvs'
4. Works on a GUID of Google Drive/team drive folder or a single google sheet too

Command to use: 
python scanDriveFolder_google_sheet.py <source_drive_id> <dest_server_path> <tab_name or prefix> <remove_flag> <archive_flag>

Example is included in the image uploaded : sample1 and sample2
