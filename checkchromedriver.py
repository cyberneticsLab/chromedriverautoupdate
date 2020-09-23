#########################################
# Python Windows chromedriver installer #
# Created by Ben Chaney - Version 1.0   #
#########################################
# This is used for Windows based OS and #
# chromedriver for Chrome automation.   #
#########################################

# Imports #
import os
from os import path
import zipfile
import requests #pip install requests
import socket 
from win32com.client import Dispatch #pip3 install pypiwin32

# Definitions #

def get_version_via_com(filename): #Used to grab file version number
    parser = Dispatch("Scripting.FileSystemObject") #powershell
    try:
        version = parser.GetFileVersion(filename) #same command as Scripting.FileSystemObject.GetFileVersion(filename)
        print('Chrome version: '+version)
    except Exception:
        return None
    return version

def download(address):
    print('Downloading chromedriver from \n - '+address)
    get_response = requests.get(address, stream=False)
    file_name  = address.split("/")[-1] #makes the filename store as the actual filename listed on the webpage
    with open(file_name, 'wb') as f:
        for chunk in get_response.iter_content(chunk_size=1024):
            if chunk: # filter out keep-alive new chunks
                f.write(chunk)
    print('[+]Download complete! File location: \n  ' + os.getcwd())
    return None
                    
def chromedriverurl(num):
    print('Getting chromedriver URL.')
    url = 'https://chromedriver.storage.googleapis.com/LATEST_RELEASE_' + num #grabs version number from other def which passes through the latest chromedriver version
    r = requests.get(url, allow_redirects=True)
    url = 'https://chromedriver.storage.googleapis.com/' + r.text + '/chromedriver_win32.zip' #grabs only the windows versions.
    print('[+]Success!')
    return url

def unzipChrome ():
    print('Extracting zipfile to: ' +os.getcwd())
    with zipfile.ZipFile('chromedriver_win32.zip',"r") as zip_ref:
        zip_ref.extractall(os.getcwd()) #extracts everything to cwd
    return None  

# Variables #

### Default location for Chrome application
paths = [r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"]      

### checks version number of Chrome application, needed for chromedriver application search
version = str(list(filter(None, [get_version_via_com(p) for p in paths]))[0])[:2]

# Attempt #

try:
    download(chromedriverurl(version))
    unzipChrome()
    print('[+]Extraction complete!')
except Exception as e:
    print('[-]An error has occured: \n' + str(e))
