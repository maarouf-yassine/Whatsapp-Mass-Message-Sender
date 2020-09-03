# -*- coding: utf-8 -*-
"""
Created on Sat Aug 22 20:28:27 2020

@author: maarouf
"""
#################
#Make Sure your phone numbers are stored on and excel sheet

#You May Need To Install The Libraries Below

#################
import time
import pyperclip
from keyboard import press
import pynput
from pynput.keyboard import Key, Controller
keyboard = Controller()

from openpyxl import load_workbook
wb = load_workbook(r'C:\Users\maaro\Desktop\Book1.xlsx')  # Work Book
ws = wb.get_sheet_by_name('Sheet1')  # Work Sheet
column = ws['A']  # Column
column_list = [str(column[x].value) for x in range(len(column))]
print(column_list)

import webbrowser

url = 'https://api.whatsapp.com/send?phone='
webbrowser.register('firefox',
	None,
	webbrowser.BackgroundBrowser("C://Program Files//Mozilla Firefox//firefox.exe"))

for x in column_list:
    webbrowser.get('firefox').open(url+x)
    time.sleep(10)
    keyboard.press(Key.ctrl)
    keyboard.press('v')
    keyboard.release('v')
    keyboard.release(Key.ctrl)
    press('enter')
