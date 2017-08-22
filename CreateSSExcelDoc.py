#Need these packages
import openpyxl #For interacting with Excel
import sys #For interacting with the OS
import time #For logging
import logging #For logging
import os
from tkinter import filedialog
from tkinter import *

#Logging set-up information, check if the log directory exists, if not create it then set up the log
now=time.strftime("%c")
os.makedirs("logs",exist_ok=True)
logging.basicConfig(filename='logs/SSExcelDocCreate.log',level=logging.DEBUG)
logging.info('\nDate and Time:\n' + now + '\n-------------------------------\n')

#Create the excel file
wb=openpyxl.Workbook()
f = Tk()
f = filedialog.asksaveasfilename(initialdir = "%USERPROFILE%\Desktop",title = "Select save location",filetypes=[("Excel Workbook",".xlsx")],initialfile="Steady State Heating and Cooling Loads",defaultextension=".xlsx")

#Try to save, but catch the error if the user hits cancel.
try:
	wb.save(f)
except Exception:
	logging.critical("SaveAs dialogue closed with 'Cancel', cannot continue")
	exit(5)

