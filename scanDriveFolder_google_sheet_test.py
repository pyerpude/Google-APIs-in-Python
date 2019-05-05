# -*- coding: utf-8 -*- 
from __future__ import print_function
import httplib2
import os
import sys
from googleapiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from datetime import datetime, timedelta
import time
import io
import hashlib
import numpy as np
import gspread
import shutil
import csv
import glob

try:
    from configparser import ConfigParser
except ImportError:
    from ConfigParser import ConfigParser  # ver. < 3.0

try:
    reload(sys)
    sys.setdefaultencoding('utf-8')  # ver. < 3.0
except:
    pass

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/drive-python-quickstart.json
SCOPES = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
CLIENT_SECRET_FILE = 'token.json'
APPLICATION_NAME = '####'

GOOGLE_MIME_TYPES = {
    'application/vnd.google-apps.document':
    ['application/vnd.openxmlformats-officedocument.wordprocessingml.document',
     '.docx'],
    'application/vnd.google-apps.spreadsheet':
    ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
     '.xlsx'],
    'application/vnd.google-apps.presentation':
    ['application/vnd.openxmlformats-officedocument.presentationml.presentation',
     '.pptx']
}

def get_credentials():
	store = Storage(CLIENT_SECRET_FILE)
	credentials = store.get()
	return credentials

def get_sheet_service():
	credentials = get_credentials()
	http = credentials.authorize(httplib2.Http())
	credentials.refresh(http)
	return gspread.authorize(credentials)

def get_drive_service():
	credentials = get_credentials()
	http = credentials.authorize(httplib2.Http())
	return discovery.build('drive', 'v3', http=http)
	
def list_files(drive_id):
	files = []
	q = "'" + drive_id + "' in parents and trashed != True"
	page_token = None
	while True:
		results = {}
		for retry in range(5):
			try:
				results = drive_service.files().list(q=q, fields="nextPageToken, files(id, name, mimeType, parents, modifiedTime, md5Checksum)", supportsTeamDrives=True, includeTeamDriveItems=True, pageToken=page_token).execute()
				break
			except Exception as e:
				print(e)
				print("Retrying list_files...", drive_id)
				continue
		files.extend(results.get('files', []))
		page_token = results.get('nextPageToken', None)
		if page_token is None:
			break
	return files

def csv_merge(tab_name, temp, data_dir):
	filepath = "%s/%s.csv" % (data_dir, tab_name)
	try:
		os.remove(filepath)
	except:
		pass
	fout = open(filepath, 'a')
	os.chdir(temp)
	for counter, file in enumerate(glob.glob("%s*" % tab_name)):
		if counter == 0:
			for line in open(file):
				fout.write(line)
		elif counter > 0:
			f = open(file)
			f.next()
			for line in f:
				fout.write(line)
	fout.close()
	return filepath
	
def process_folder(drive_id):
	'''Checks folders and continues to run listing all the folders and files inside them'''
	files = list_files(drive_id)
	for file in files:
		#f = open("google_drive_manifest","w+")
		if file['mimeType'] == 'application/vnd.google-apps.folder':
			process_folder(file['id'])
		elif file['mimeType'] == 'application/vnd.google-apps.spreadsheet':
			process_file(file)
		
def open_sheet(file_id):
	for retry in range(5):
		try:
			sheet = sheet_service.open_by_key(file_id)
			return sheet
		except Exception as e:
			print(e)
			print("Retrying open_sheet: " + file_id)
			time.sleep(60) # sleep for 1 min
			continue
	return None

def process_file(file):
	'''Downloads all the files from given folder if it is a first time, else checks for the 
	modified time of the file in drive and downloads it if changed'''
	#file["name"] = file["name"] + '.' + file["id"] #prefix the file name with its guid
	#print(file["name"])
	#drive_files = 'Drive_files'
	#local_path = os.path.join(local_path, drive_files)
	
	sheet = open_sheet(file['id'])
	if sheet is None:
		print("Unable to open sheet: " + file['id'])
		return
	for tab in tab_name:
		sheetId = sheet.worksheet(tab)._properties['sheetId']
		body = {
		  "requests": [
			{
			  "updateCells": {
				"range": {
				  "sheetId": sheetId,
				  "startRowIndex": 0,
				  "startColumnIndex": 0
				},
				"rows": [
				  {
					"values": [
					  {
						"userEnteredFormat": {
						  "wrapStrategy": "OVERFLOW_STRATEGY"
						}
					  }
					]
				  }
				],
				"fields": "userEnteredFormat.wrapStrategy"
			  }
			}
		  ]
		}
		res = sheet.batch_update(body)
		csvfile = os.path.join(temp_folder, tab+".csv")
		full_path = os.path.join(csvfile, file['id'])
		print("Processing tab", full_path)
		for retry in range(5):
			try:
				worksheet = sheet.worksheet(tab)
				all_values = worksheet.get_all_values()
				if os.path.isfile(csvfile):
					all_values.pop(0) # remove header
				with open(csvfile, 'a') as f:
					writer = csv.writer(f, delimiter='\325')
					writer.writerows(all_values)
				break
			except Exception as e:
				print(e)
				if retry == 4:
					print("Unable to process tab: " + tab + " of sheet: " + file['id'])
				else:
					print("Retrying tab: " + tab + " of sheet: " + file['id'])
				time.sleep(60)
				sheet_service.login()
				sheet = open_sheet(file['id'])
				continue
		time.sleep(2)
	print("Sheet processed " + file['id'])
	
	if remove_flag == 'Y':
		remove_file_from_gdrive(file)
		print(file['name'] + ' removed')
	if archive_flag == 'Y':
		archive_file_from_gdrive(file)
		print(file['name'] + ' archived')
		
def remove_file_from_gdrive(file):
	request = drive_service.files().delete(fileId=file["id"],supportsTeamDrives=True).execute()
	
def archive_file_from_gdrive(file):
	folder_id = "1SpVyKpXjjz_7i1FejiNvSrgFLJvcm4Yk"
	file_id = file['id']
	#Retrieve the existing parents to remove
	#file = service.files().get(fileId=file_id,fields='parents',supportsTeamDrives=True).execute()
	previous_parents = ",".join(file.get('parents'))
	#Move the file to the new folder
	file = drive_service.files().update(fileId=file_id,addParents=folder_id,removeParents=previous_parents,supportsTeamDrives=True,fields='id, parents').execute()

def main():
	global local_path
	global tab_name
	global remove_flag
	global archive_flag
	global drive_service
	global sheet_service
	global fileHandler
	global data_dir
	global temp_folder
	drive_id = sys.argv[1] # folder or file ID from Google Team Drive
	local_path = sys.argv[2] #"C:\\Users\\bfm763\\Downloaded" #sys.argv[2] # local OS path
	tab_name = sys.argv[3] #Enter the file name or prefix of file name to be downloaded
	tab_name = tab_name.split(',')
	print(tab_name)
	remove_flag = sys.argv[4] #Remove flag either 'Y' or 'N' - If 'Y', it will remove the files from drive
	archive_flag = sys.argv[5] #To move the file in archive folder of drive
	drive_service = get_drive_service()
	sheet_service = get_sheet_service()
	print ('---------------Starting to check at %s--------------' % (datetime.now()))
	#manifest =  os.path.join(local_path, '/google_drive_manifest.txt')
	data_dir = "/data/temp/Google_Drive_Metadata/Google_Sheet_Metadata/"
	temp_folder = "/data/temp/Google_Drive_Metadata/Google_Sheet_Metadata/temp"
	#shutil.rmtree(temp_folder)
	#os.makedirs(temp_folder)
	manifest = local_path + '/google_drive_manifest.txt'
	for csv_existing in glob.glob(data_dir + "*.csv"):
		os.remove(csv_existing)
	if os.path.isfile(manifest):	
		os.remove(manifest)
	print(manifest)
	
	print("Manifest file created")
	#To check if the input ID is a file or a folder in google team drive
	result = drive_service.files().get(fileId=drive_id, supportsTeamDrives=True, fields="id, name, mimeType, parents, modifiedTime, md5Checksum").execute()
	if result['mimeType'] == 'application/vnd.google-apps.folder':
		process_folder(drive_id)
	elif result['mimeType'] == 'application/vnd.google-apps.spreadsheet':
		process_file(result)
	files = os.listdir(temp_folder)
	fileHandler = open(manifest, "a")
	for csv_file in files:
		shutil.move(os.path.join(temp_folder, csv_file), data_dir)
		fileHandler.write(os.path.join(data_dir, csv_file) + "\n")
	fileHandler.close()

if __name__ == '__main__':
	print(sys.argv)
	if len(sys.argv) == 6:
		main()
	else:
		print("\nCorrect command: python scanDriveFolder_google_sheet.py <source_drive_id> <dest_server_path> <tab_name or prefix> <remove_flag> <archive_flag>\n")
