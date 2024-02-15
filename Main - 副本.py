import numpy as np
import time
import pandas as pd
import win32api
import win32con
import win32gui
import os.path as op
from PIL import Image
import sys
import datetime
from pptx import Presentation
from datetime import date

day = [0, 0, 0, 0, 0, 0, 0]
sub = [[0] * 100 for _ in range(100)]
te = [[0] * 100 for _ in range(100)]
ts = [[0] * 100 for _ in range(100)]

def get_col_num(ws):
    return ws.shape[1]

ti = time.ctime()
from datetime import date
dayOfWeek = date.today().weekday()


for i in range(7):
    day[i] = pd.read_excel("ClassSchedule.xlsx", sheet_name=i)
    rows, columns = day[i].shape
    for j in range(rows):
        ts[i][j] = day[i].iloc[j, 1]
        ts[i][j] = str(ts[i][j])
        te[i][j] = day[i].iloc[j, 2]
        sub[i][j] = day[i].iloc[j, 3]
ppt = Presentation("PPT.pptx")
import os  # 添加这一行来导入os模块
import ctypes  
import win32api  
import win32con  
import win32gui   
while True:
    now = datetime.datetime.now()
    xingqi = dayOfWeek + 1
    print(now.strftime("%H:%M:%S"))
    AAAAA = now.strftime("%H:%M:%S")
    scheduled_time = datetime.datetime.strptime(AAAAA, "%H:%M:%S").time()
    schetime = str(scheduled_time)
    print (schetime)
    print (dayOfWeek)
    print (ts[dayOfWeek][2])
    for j in range(rows):
        print (j)
        if schetime == ts[dayOfWeek][1]:
            print("AAAAAAA")
            rownum = j + 1
            image_path = os.path.abspath(f'table\\pptjpg\\{xingqi}\\幻灯片{rownum}.JPG')
            SPI_SETDESKWALLPAPER = 20
            ctypes.windll.user32.SystemParametersInfoW(SPI_SETDESKWALLPAPER, 0, image_path, win32con.SPIF_UPDATEINIFILE | win32con.SPIF_SENDWININICHANGE)
            win32gui.SendMessageTimeout(win32con.HWND_BROADCAST, win32con.WM_SETTINGCHANGE, 0, 'SPI_SETDESKWALLPAPER', win32con.SMTO_ABORTIFHUNG, 1000)
            AAAAA = 0
            break
        print (schetime+"AADSASD")
    time.sleep(0.63)
