#import Python libraries
import os
from os.path import isfile, join
from os import listdir
import sys
import pyshorteners as sh
import pandas as pd
import time
from datetime import datetime
'''
import win32gui
import win32process
import win32api
import win32com.client as win32
import glob
import shutil

import pywinauto
import pywinauto
from pynput.mouse import Button, Controller
mouse = Controller()

from pynput.keyboard import Key, Controller
keyboard = Controller()
'''
import base64
import re

import streamlit as st
import tkinter as tk
from tkinter import filedialog

# Set up tkinter
root = tk.Tk()
root.withdraw()

# Make folder picker dialog appear on top of other windows
root.wm_attributes('-topmost', 1)

# Folder picker button
st.title('Folder Picker')
st.write('Please select a folder:')
clicked = st.button('Folder Picker')
dirname = ''
#month = ''
if clicked:
    dirname = st.text_input('Selected folder:', filedialog.askdirectory(master=root))

filelist=[]
es=[]
td_date = time.strftime("%d-%b-%Y")
url = "https://miscbhd.sharepoint.com/:b:/r/sites/ES-FIN/Shared Documents/14 AR & Treasury/COPY OF INVOICES/"
for root, dirs, files in os.walk(dirname):
      for file in files:
             filename=os.path.join(root, file)
             filelist.append(filename)
             mon = root.split("/")[-1]
             month = mon.split('\\')[0]
             full_url = url +file[14:18]+"/"+ mon +"/"+ file
             link = full_url.replace('%','%25').replace('&','%26').replace(' ','%20')
             if len(link)>=245:
                 #print("yes")
                 s = sh.Shortener()
                 #print(s.tinyurl.short(link))
                 es.append([file[9:13],file[0:8],file[14:18],file,s.tinyurl.short(link)])
             else:
                 #es.append([second_char,first_char,third_char,x,link])
                 #st.write(link)
                 es.append([file[9:13],file[0:8],file[14:18],file,link])
             #st.write(file[0:8])

#st.write(root.split("/"))
 
st.write(filelist)
st.write(es)

df = pd.DataFrame(es)
df.columns = ['Company code','Document no','Fiscal year','URL Name','URL']
excel_name = 'ZFI_UPLOAD_URL '+month+' '+ td_date + '.xlsx'
df.to_excel(excel_name, index=False)

#print('Data extract successfully')
st.write('Data extract successfully')
cwd = os.getcwd()
file_dir = cwd + "\\" + excel_name
#print(file_dir)
'''
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = "mhazlan.hamdan@miscbhd.com;syazwan.naim@miscbhd.com;nurulain.mohdyasin@miscbhd.com;hadi.mdnor@miscbhd.com"
mail.Subject = 'ES Saving Invoice SAP URL'
mail.Body = 'Message body'
mail.HTMLBody = '<h2>Number of Invoice been Process : '+str(len(es))+'</h2>' #this field is optional

# To attach a file to the email (optional):
attachment  = file_dir
mail.Attachments.Add(attachment)

mail.Send()

def windowEnumerationHandler(hwnd, top_windows):
    top_windows.append((hwnd, win32gui.GetWindowText(hwnd)))


top_windows = []
win32gui.EnumWindows(windowEnumerationHandler, top_windows)


for il in top_windows:
    if ('SAP LOGON' in il[1].upper()) or (il[1].upper() == 'SAP'):
        print (il[1].upper())
        win32gui.SetForegroundWindow(il[0])
        time.sleep(1)
        with keyboard.pressed(Key.alt):
            keyboard.press(Key.f4)
            keyboard.release(Key.f4)

    if ('ZPDEV' in il[1].upper()):
        print (il[1].upper())
        win32gui.SetForegroundWindow(il[0])
        time.sleep(1)
        with keyboard.pressed(Key.alt):
            keyboard.press(Key.f4)
            keyboard.release(Key.f4)
        time.sleep(1)
        keyboard.press(Key.tab)
        time.sleep(1)
        keyboard.press(Key.enter)


class WindowMgr:
    '''Encapsulates some calls to the winapi for window management'''

    def __init__ (self):
        '''Constructor'''
        self._handle = None

    def find_window(self, class_name, window_name=None):
        '''find a window by its class_name'''
        self._handle = win32gui.FindWindow(class_name, window_name)

    def _window_enum_callback(self, hwnd, wildcard):
        '''Pass to win32gui.EnumWindows() to check all the opened windows'''
        if re.match(wildcard, str(win32gui.GetWindowText(hwnd))) is not None:
            self._handle = hwnd

    def find_window_wildcard(self, wildcard):
        '''find a window whose title matches the wildcard regex'''
        self._handle = None
        win32gui.EnumWindows(self._window_enum_callback, wildcard)

    def set_foreground(self):
        '''put the window in the foreground'''
        win32gui.SetForegroundWindow(self._handle)

keyboard = Controller()

#uid = 'MHAZLAN'
uid = 'junitaak'
#uid = sapid
#pxd = base64.b64decode(b'QWRhLURnYUAwNDIwMjI=').decode('utf-8')
pxd = 'year2022-MEIMEI'
#pxd = base64.b64decode(b'eWVhcjIwMjItTUVJTUVJ').decode('utf-8')

#pxd = base64.b64encode(b'year2022-MEIMEI').decode('utf-8')

w = WindowMgr()

os.startfile(r'C:/Program Files (x86)/SAP/FrontEnd/SAPgui/saplogon.exe')
time.sleep(5)

with keyboard.pressed(Key.alt):
    keyboard.press(Key.space)
    keyboard.release(Key.space)
time.sleep(3)

keyboard.press('x')         
keyboard.release('x')
time.sleep(4) 

for i in range(11):
   keyboard.press(Key.down)
   time.sleep(0.15)

keyboard.press(Key.enter)
time.sleep(3)
keyboard.type(uid)
time.sleep(2)
keyboard.press(Key.tab)
time.sleep(2)
keyboard.type(pxd)
time.sleep(2)
keyboard.press(Key.enter)
time.sleep(1)
keyboard.press(Key.down)
time.sleep(1)
keyboard.release(Key.space)
time.sleep(1)
keyboard.type('ZFI_UPLOAD_URL')
time.sleep(1)

keyboard.press(Key.enter)
time.sleep(1)
keyboard.type(file_dir)
time.sleep(2)
keyboard.press(Key.tab)
time.sleep(2)
keyboard.press(Key.space)
time.sleep(2)
keyboard.press(Key.f8)
time.sleep(2)

with keyboard.pressed(Key.shift):
    keyboard.press(Key.tab)
    keyboard.release(Key.tab)
time.sleep(1)
keyboard.press(Key.enter)
time.sleep(1)
'''
