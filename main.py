#import Python libraries
import os
from os.path import isfile, join
from os import listdir
import sys
import pyshorteners as sh
import pandas as pd
import time
from datetime import datetime

import base64
import re
import streamlit as st

from pathlib import Path

#str_path = "my_path"
#path = Path(str_path)


dirname = r'D:\Users\mhazlan.hamdan\MISC Group\EagleStar Finance - COPY OF INVOICES\2022\07. JUL 2022'
path = Path(dirname)
st.write(path)
filelist=[]
es=[]
td_date = time.strftime("%d-%b-%Y")
url = "https://miscbhd.sharepoint.com/:b:/r/sites/ES-FIN/Shared Documents/14 AR & Treasury/COPY OF INVOICES/"
for root, dirs, files in os.walk(path):
      for file in files:
             filename=os.path.join(root, file)
             filelist.append(filename)
             #st.write(root)
             mon = root.split("\\")[-2]
             com = root.split("\\")[-1]
             #st.write(mon)
             month = mon.split('\\')[0]
             full_url = url +file[14:18]+"/"+ mon +"/"+com+"/"+ file
             link = full_url.replace('%','%25').replace('&','%26').replace(' ','%20')
             es.append([file[9:13],file[0:8],file[14:18],file,link])
                          

st.write(filelist)
st.write(es)
