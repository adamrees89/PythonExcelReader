import sqlite3
import openpyxl
import sys
import time
import logging
import os

#Logging set-up information, check if the log directory exists, if not create it then set up the log
now=time.strftime("%c")
os.makedirs("logs",exist_ok=True)
logging.basicConfig(filename='logs/templatereader.log',level=logging.DEBUG)
logging.info('\nDate and Time:\n' + now + '\n-------------------------------\n')

#Create the cell function
def templateCell(s,sn,col,r):
	cellRef = col + r
	val = s[cellRef].value
	fname = s[cellRef].font.name
	fsize = s[cellRef].font.size
	fbold = int(s[cellRef].font.bold == 'true')
	fital = int(s[cellRef].font.italic == 'true')
	ccolour = s[cellRef].fill.start_color.index
	data=[cellRef,val,fname,fsize,fbold,fital,ccolour]
	c.execute("INSERT INTO "+sn+" VALUES (?,?,?,?,?,?,?)",data)

def templateSheet(s):
	#This class will call the template cell class from its method, and from init will create the sql table
	sn = s.title
	sn = sn.replace(" ","").replace("+","").replace("-","").replace("/","").replace("_","").replace("&","").replace("%","")
	column2 = 'Cell'
	column3 = 'Value'
	column4 = 'Font_Name'
	column5 = 'Font_Size'
	column6 = 'Font_Bold'
	column7 = 'Font_Italic'
	column8 = 'Cell_Colour'
	fieldtype1 = 'INTEGER' #INTEGER, TEXT, NULL, REAL, BLOB
	fieldtype2 = 'TEXT'
	try:
		c.execute('CREATE TABLE {tn} ({c2} {ft2}, {c3} {ft2}, {c4} {ft2}, {c5} {ft2}, {c6} {ft1}, {c7} {ft1}, {c8} {ft1})'.format(tn=sn,ft1=fieldtype1,ft2=fieldtype2,c2=column2,c3=column3,c4=column4,c5=column5,c6=column6,c7=column7,c8=column8))
	except sqlite3.Error:
		logging.critical('Adding data threw an unexpected error, does the table exist within the database?\nCannot continue\nExit(5)')
		print('Is the database already open?')
		sys.exit(5)
	ExtentRow = s.max_row
	ExtentColumn = s.max_column
	rows = list(range(1,ExtentRow))
	column = list(range(1,ExtentColumn))
	for co in column:
		coL = openpyxl.utils.get_column_letter(co)
		for ro in rows:
			templateCell(s,sn,str(coL),str(ro))


#Initialise the sqllite3 file, and delete the old one if it exists
try:
    os.remove("template.db")
except OSError:
    pass

conn = sqlite3.connect("template.db")
c = conn.cursor()

#Get the workbook, and create the table for the first sheet
wb = openpyxl.load_workbook(filename='Sample Excel Files/testExcel.xlsx')

for ws in wb.worksheets:
	templateSheet(ws)
	conn.commit()

conn.close()
