import templateReader
import openpyxl
import os
import time
import logging

#Logging set-up information, check if the log directory exists, if not create it then set up the log
now=time.strftime("%c")
os.makedirs("logs",exist_ok=True)
logging.basicConfig(filename='logs/ExcelReader.log',level=logging.DEBUG)
logging.info('\nDate and Time:\n' + now + '\n-------------------------------\n')

	try:
		os.remove("template.db")
	except OSError:
		pass

wb = openpyxl.load_workbook(filename='Sample Excel Files/TM_Heating & Cooling.xlsx')

for ws in wb.worksheets:
	templateReader.templateSheet(ws)

