import unittest
import templateReader
import openpyxl

class templateReaderFunctionTest(unittest.TestCase):
	""" Class doc """
	
	def test_Excelfile(self):
		""" Class initialiser """
		f='Sample Excel Files/TM_Heating & Cooling.xlsx'
		try:
			wb = openpyxl.load_workbook(filename=f)
			for ws in wb.worksheets:
				templateReader.templateSheet(ws)
		except ExceptionType:
			self.fail("templateReader.templateSheet() threw an exception with the sample .xlsx file:".format(f))

	def test_Wordfile(self):
		""" Class initialiser """
		f='Sample Excel Files/sample.docx'
		try:
			wb = openpyxl.load_workbook(filename=f)
			for ws in wb.worksheets:
				templateReader.templateSheet(ws)
		except ExceptionType:
			self.fail("templateReader.templateSheet() threw an exception with the sample .docx file:".format(f))


def main():
	unittest.main()

if __name__ == '__main__':
	main()

